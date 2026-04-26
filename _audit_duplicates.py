r"""Read-only audit: detect duplicate messages between 'Imported PST' and
the rest of a mailbox, before running _flatten_imported.py.

For each mailbox processed, walks two subtrees:

  - SOURCE: 'Imported PST' and all its descendants
  - DESTINATION: every other folder in the mailbox (Inbox, Sent Items,
    Drafts, custom folders, etc., recursively)

Two overlap checks are run:

  1. Message-ID match (strict): same internetMessageId on both sides.
     Catches duplicates when the other importer preserved SMTP headers.

  2. Fuzzy match (subject + sent-minute + from address): catches
     duplicates even when the other importer rewrote Message-IDs (Outlook
     drag/drop and some MAPI-based tools do this).

A non-zero "fuzzy_only" column indicates messages that look like the same
logical email but have different or missing Message-IDs - which means
--skip-duplicates would need to dedupe by the fuzzy key, not the Message-ID.

This script makes NO changes - all GET requests, no POST/PATCH/DELETE.

Usage from Graph-Migrate/ (Windows):
  .\.venv\Scripts\python.exe _audit_duplicates.py -c config.toml --mailbox UPN
  .\.venv\Scripts\python.exe _audit_duplicates.py -c config.toml -m mapping.csv
  .\.venv\Scripts\python.exe _audit_duplicates.py -c config.toml --mailbox UPN --show 50
"""
from __future__ import annotations

import argparse
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import quote

from loguru import logger

from jtet_pstmigrate.auth import AppPool
from jtet_pstmigrate.config import AppConfig, expand_user_paths
from jtet_pstmigrate.graph_client import GraphClient, GraphError
from jtet_pstmigrate.orchestrator import load_mapping

ROOT_FOLDER_NAME = "Imported PST"
PAGE_SIZE = 999  # Graph max for /messages collection


FuzzyKey = tuple[str, str, str]  # (subject_norm, sent_at_minute, from_addr_lower)

# Strip RFC-style reply/forward prefixes that some clients add when
# importing or replying. Repeated to peel "RE: FW: RE:" chains.
_REPLY_PREFIX_RE = re.compile(r"^\s*(?:re|fw|fwd|aw|sv|tr|wg|antwort|antw)\s*[:\[\(]?\s*", re.IGNORECASE)


def _norm_subject(s: str) -> str:
    s = (s or "").strip()
    while True:
        new = _REPLY_PREFIX_RE.sub("", s, count=1)
        if new == s:
            break
        s = new
    return " ".join(s.split()).lower()


@dataclass(slots=True)
class Msg:
    """Lightweight record of a message for the audit."""
    message_id: str           # internetMessageId, lowercased ("" if missing)
    folder_path: str          # human-readable folder breadcrumb
    subject: str
    from_addr: str
    sent_at: str              # ISO string from sentDateTime; "" if missing
    fuzzy: FuzzyKey           # (subject_norm, sent_at_minute, from_addr_lower)


@dataclass(slots=True)
class Side:
    """One side of the audit: either the Imported PST tree or everything else."""
    name: str
    folders_walked: int = 0
    messages_with_id: int = 0
    messages_without_id: int = 0
    by_id: dict[str, list[Msg]] = field(default_factory=dict)        # internet msg id -> messages
    by_fuzzy: dict[FuzzyKey, list[Msg]] = field(default_factory=dict)  # fuzzy key -> messages


def _strip_base(url: str) -> str:
    base = "https://graph.microsoft.com/v1.0"
    return url[len(base):] if url.startswith(base) else url


def list_child_folders(graph: GraphClient, mailbox: str, parent_id: str) -> list[dict]:
    out: list[dict] = []
    path = (
        f"/users/{quote(mailbox)}/mailFolders/{parent_id}/childFolders"
        f"?$top=100&$select=id,displayName,childFolderCount,totalItemCount"
    )
    while path:
        resp = graph.get(path, expect_status=(200,))
        body = resp.json()
        out.extend(body.get("value", []))
        next_link = body.get("@odata.nextLink")
        path = _strip_base(next_link) if next_link else None
    return out


def iter_messages(graph: GraphClient, mailbox: str, folder_id: str):
    """Yield message dicts (with internetMessageId, subject, from, sentDateTime) from a folder."""
    path = (
        f"/users/{quote(mailbox)}/mailFolders/{folder_id}/messages"
        f"?$top={PAGE_SIZE}&$select=id,internetMessageId,subject,from,sentDateTime"
    )
    while path:
        resp = graph.get(path, expect_status=(200,))
        body = resp.json()
        for m in body.get("value", []):
            yield m
        next_link = body.get("@odata.nextLink")
        path = _strip_base(next_link) if next_link else None


def _from_addr(msg: dict) -> str:
    f = msg.get("from") or {}
    eb = f.get("emailAddress") or {}
    return eb.get("address") or eb.get("name") or ""


def _record_messages(side: Side, graph: GraphClient, mailbox: str, folder_id: str, folder_path: str) -> None:
    side.folders_walked += 1
    for m in iter_messages(graph, mailbox, folder_id):
        mid_raw = (m.get("internetMessageId") or "").strip()
        sent_at = (m.get("sentDateTime") or "")[:19]
        from_addr = _from_addr(msg=m)
        subject = (m.get("subject") or "")[:120]
        # Minute-precision timestamp tolerates the few-second drift some
        # importers introduce when re-serializing dates.
        sent_minute = sent_at[:16]
        fuzzy: FuzzyKey = (_norm_subject(subject), sent_minute, from_addr.lower())

        rec = Msg(
            message_id=mid_raw.lower(),
            folder_path=folder_path,
            subject=subject,
            from_addr=from_addr,
            sent_at=sent_at,
            fuzzy=fuzzy,
        )
        if mid_raw:
            side.messages_with_id += 1
            side.by_id.setdefault(rec.message_id, []).append(rec)
        else:
            side.messages_without_id += 1
        # Only index for fuzzy matching when we have at least subject+sender or
        # subject+sent time; otherwise it produces noisy ("", "", "") collisions.
        if (fuzzy[0] or fuzzy[2]) and fuzzy[1]:
            side.by_fuzzy.setdefault(fuzzy, []).append(rec)


def walk_subtree(graph: GraphClient, mailbox: str, root_folder: dict, root_path: str, side: Side) -> None:
    """DFS through `root_folder` and below, recording every message into `side`."""
    _record_messages(side, graph, mailbox, root_folder["id"], root_path)
    if int(root_folder.get("childFolderCount") or 0) <= 0:
        return
    for child in list_child_folders(graph, mailbox, root_folder["id"]):
        child_path = f"{root_path}/{child['displayName']}"
        walk_subtree(graph, mailbox, child, child_path, side)


def find_imported_root(graph: GraphClient, mailbox: str) -> dict | None:
    for c in list_child_folders(graph, mailbox, "msgFolderRoot"):
        if c["displayName"] == ROOT_FOLDER_NAME:
            return c
    return None


def audit_mailbox(graph: GraphClient, mailbox: str, *, sample_size: int) -> dict:
    log = logger.bind(ctx=f"audit[{mailbox}]")
    result: dict = {"mailbox": mailbox, "skipped": False}

    imported = find_imported_root(graph, mailbox)
    if imported is None:
        log.info("No '{}' folder found; nothing to audit.", ROOT_FOLDER_NAME)
        result["skipped"] = True
        return result

    src = Side(name="source (Imported PST subtree)")
    dst = Side(name="destination (live mailbox excl. Imported PST)")

    log.info("Walking '{}' subtree...", ROOT_FOLDER_NAME)
    walk_subtree(graph, mailbox, imported, ROOT_FOLDER_NAME, src)
    log.info(
        "  -> {} folders, {} messages w/ ID, {} without",
        src.folders_walked, src.messages_with_id, src.messages_without_id,
    )

    log.info("Walking rest of mailbox...")
    for top in list_child_folders(graph, mailbox, "msgFolderRoot"):
        if top["displayName"] == ROOT_FOLDER_NAME:
            continue
        walk_subtree(graph, mailbox, top, top["displayName"], dst)
    log.info(
        "  -> {} folders, {} messages w/ ID, {} without",
        dst.folders_walked, dst.messages_with_id, dst.messages_without_id,
    )

    # ---- Message-ID overlap ----
    overlap_ids = set(src.by_id.keys()) & set(dst.by_id.keys())
    src_msgs_in_overlap = sum(len(src.by_id[i]) for i in overlap_ids)
    dst_msgs_in_overlap = sum(len(dst.by_id[i]) for i in overlap_ids)

    # ---- Fuzzy overlap (subject + sent-minute + from) ----
    overlap_fuzzy = set(src.by_fuzzy.keys()) & set(dst.by_fuzzy.keys())
    src_msgs_in_fuzzy = sum(len(src.by_fuzzy[k]) for k in overlap_fuzzy)
    dst_msgs_in_fuzzy = sum(len(dst.by_fuzzy[k]) for k in overlap_fuzzy)

    # Fuzzy hits where Message-IDs disagree -> evidence the other importer
    # rewrote IDs. Pick one representative pair per fuzzy key for sampling.
    fuzzy_only_pairs: list[tuple[Msg, Msg]] = []
    for k in overlap_fuzzy:
        s_rec = src.by_fuzzy[k][0]
        d_rec = dst.by_fuzzy[k][0]
        # If either side has no id, or the ids differ, this is "fuzzy-only".
        if not s_rec.message_id or not d_rec.message_id or s_rec.message_id != d_rec.message_id:
            fuzzy_only_pairs.append((s_rec, d_rec))

    result.update({
        "src_folders": src.folders_walked,
        "src_msgs_with_id": src.messages_with_id,
        "src_msgs_no_id": src.messages_without_id,
        "dst_folders": dst.folders_walked,
        "dst_msgs_with_id": dst.messages_with_id,
        "dst_msgs_no_id": dst.messages_without_id,
        "overlap_ids": len(overlap_ids),
        "src_msgs_in_overlap": src_msgs_in_overlap,
        "dst_msgs_in_overlap": dst_msgs_in_overlap,
        "overlap_fuzzy": len(overlap_fuzzy),
        "src_msgs_in_fuzzy": src_msgs_in_fuzzy,
        "dst_msgs_in_fuzzy": dst_msgs_in_fuzzy,
        "fuzzy_only_pairs": len(fuzzy_only_pairs),
    })

    log.info(
        "Message-ID overlap: {} unique IDs - {} src / {} dst messages collide.",
        len(overlap_ids), src_msgs_in_overlap, dst_msgs_in_overlap,
    )
    log.info(
        "Fuzzy overlap   : {} unique (subject+sent+from) keys - {} src / {} dst messages collide.",
        len(overlap_fuzzy), src_msgs_in_fuzzy, dst_msgs_in_fuzzy,
    )
    if fuzzy_only_pairs:
        log.warning(
            "{} fuzzy matches have differing/missing Message-IDs -> the other importer "
            "likely rewrote IDs; --skip-duplicates would need to use the fuzzy key.",
            len(fuzzy_only_pairs),
        )

    # ---- Sample: Message-ID matches (full duplicates) ----
    if overlap_ids and sample_size > 0:
        log.info("Sample (up to {} Message-ID matches):", sample_size)
        for i, mid in enumerate(sorted(overlap_ids)):
            if i >= sample_size:
                break
            sample_src = src.by_id[mid][0]
            sample_dst = dst.by_id[mid][0]
            subj = sample_src.subject or sample_dst.subject or "(no subject)"
            log.info(
                "  ID  {} | from {!r} | sent {} | src={!r} ({}x) | dst={!r} ({}x) | subject: {}",
                mid, sample_src.from_addr or sample_dst.from_addr,
                sample_src.sent_at or sample_dst.sent_at,
                sample_src.folder_path, len(src.by_id[mid]),
                sample_dst.folder_path, len(dst.by_id[mid]),
                subj,
            )

    # ---- Sample: fuzzy matches with different/missing Message-IDs ----
    if fuzzy_only_pairs and sample_size > 0:
        log.info("Sample (up to {} fuzzy matches w/ differing Message-IDs):", sample_size)
        for i, (s_rec, d_rec) in enumerate(fuzzy_only_pairs):
            if i >= sample_size:
                break
            log.info(
                "  FUZ from {!r} | sent {} | src={!r} | dst={!r} | subject: {}",
                s_rec.from_addr or d_rec.from_addr,
                s_rec.sent_at or d_rec.sent_at,
                s_rec.folder_path,
                d_rec.folder_path,
                s_rec.subject or d_rec.subject or "(no subject)",
            )
            log.info(
                "      src-id={!r}  dst-id={!r}",
                s_rec.message_id or "(missing)",
                d_rec.message_id or "(missing)",
            )
    return result


def main() -> int:
    ap = argparse.ArgumentParser(description="Audit Message-ID overlap between Imported PST and rest of mailbox.")
    ap.add_argument("-c", "--config", required=True)
    ap.add_argument("-m", "--mapping", help="mapping.csv path")
    ap.add_argument("--mailbox", action="append", help="UPN of a single mailbox (repeatable)")
    ap.add_argument(
        "--show", type=int, default=20,
        help="how many sample duplicate Message-IDs to print per mailbox (0 = none, default 20)",
    )
    ns = ap.parse_args()

    cfg = expand_user_paths(AppConfig.load(Path(ns.config)))

    if ns.mailbox:
        mailboxes = sorted(set(ns.mailbox))
    elif ns.mapping:
        mailboxes = sorted({r.target_mailbox for r in load_mapping(Path(ns.mapping))})
    else:
        ap.error("Provide either --mapping or one or more --mailbox UPNs.")

    logger.remove()
    logger.add(sys.stderr, level="INFO", format="{time:HH:mm:ss} | {level: <7} | {extra[ctx]} | {message}")

    pool = AppPool(cfg.apps)
    parallelism = min(cfg.migration.max_parallel_mailboxes, len(mailboxes)) or 1

    logger.bind(ctx="audit").info(
        "Auditing {} mailbox(es) with {} parallel workers, {} app(s).",
        len(mailboxes), parallelism, len(cfg.apps),
    )

    results: list[dict] = []
    with GraphClient(pool, cfg.throttle) as graph:
        with ThreadPoolExecutor(max_workers=parallelism, thread_name_prefix="audit") as ex:
            futs = {
                ex.submit(audit_mailbox, graph, m, sample_size=ns.show): m
                for m in mailboxes
            }
            for fut in as_completed(futs):
                m = futs[fut]
                try:
                    results.append(fut.result())
                except GraphError as e:
                    logger.bind(ctx=f"audit[{m}]").error("Graph error: {}", e)
                    results.append({"mailbox": m, "error": str(e)})
                except Exception as e:
                    logger.bind(ctx=f"audit[{m}]").exception("Crashed: {}", e)
                    results.append({"mailbox": m, "error": str(e)})

    # Summary table
    print()
    hdr = (
        f"{'mailbox':<46} {'src_msgs':>10} {'dst_msgs':>10} "
        f"{'id_match':>10} {'fuzzy':>10} {'fuzzy_only':>11}  status"
    )
    print(hdr)
    print("-" * len(hdr))
    for r in sorted(results, key=lambda x: x["mailbox"]):
        if "error" in r:
            print(
                f"{r['mailbox']:<46} {'-':>10} {'-':>10} {'-':>10} {'-':>10} {'-':>11}  "
                f"ERROR: {r['error'][:30]}"
            )
        elif r.get("skipped"):
            print(
                f"{r['mailbox']:<46} {'-':>10} {'-':>10} {'-':>10} {'-':>10} {'-':>11}  "
                f"no Imported PST"
            )
        else:
            print(
                f"{r['mailbox']:<46} "
                f"{r['src_msgs_with_id'] + r['src_msgs_no_id']:>10} "
                f"{r['dst_msgs_with_id'] + r['dst_msgs_no_id']:>10} "
                f"{r['src_msgs_in_overlap']:>10} "
                f"{r['src_msgs_in_fuzzy']:>10} "
                f"{r['fuzzy_only_pairs']:>11}  ok"
            )
    print()
    print("Columns: src_msgs/dst_msgs = total messages in each subtree;")
    print("         id_match  = src messages whose internetMessageId is also present in the live mailbox;")
    print("         fuzzy     = src messages whose (subject, sent-minute, from) is also present in the live mailbox;")
    print("         fuzzy_only= matches found by fuzzy key but with differing/missing Message-IDs")
    print("                     (high values here suggest the other importer rewrote IDs).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
