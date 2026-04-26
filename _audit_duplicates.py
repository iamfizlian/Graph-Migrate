r"""Read-only audit: find Message-ID overlap between 'Imported PST' and the
rest of a mailbox, before running _flatten_imported.py.

For each mailbox processed, walks two subtrees and collects every message
with an internetMessageId:

  - SOURCE: 'Imported PST' and all its descendants
  - DESTINATION: every other folder in the mailbox (Inbox, Sent Items,
    Drafts, custom folders, etc., recursively)

Reports the size of each set, their intersection, and a sample of the
overlapping messages so you can decide whether to:
  - just run _flatten_imported.py (no overlap, no dup risk)
  - run with --skip-duplicates (overlap exists; not implemented here yet)
  - investigate further before flattening

This script makes NO changes - all GET requests, no POST/PATCH/DELETE.

Usage from Graph-Migrate/ (Windows):
  .\.venv\Scripts\python.exe _audit_duplicates.py -c config.toml --mailbox UPN
  .\.venv\Scripts\python.exe _audit_duplicates.py -c config.toml -m mapping.csv
  .\.venv\Scripts\python.exe _audit_duplicates.py -c config.toml --mailbox UPN --show 50
"""
from __future__ import annotations

import argparse
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


@dataclass(slots=True)
class Msg:
    """Lightweight record of a message for the audit."""
    message_id: str           # internetMessageId, lowercased
    folder_path: str          # human-readable folder breadcrumb
    subject: str
    from_addr: str
    sent_at: str              # ISO string from sentDateTime; "" if missing


@dataclass(slots=True)
class Side:
    """One side of the audit: either the Imported PST tree or everything else."""
    name: str
    folders_walked: int = 0
    messages_with_id: int = 0
    messages_without_id: int = 0
    by_id: dict[str, list[Msg]] = field(default_factory=dict)  # id -> messages with that id


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
        mid = m.get("internetMessageId") or ""
        if not mid:
            side.messages_without_id += 1
            continue
        rec = Msg(
            message_id=mid.strip().lower(),
            folder_path=folder_path,
            subject=(m.get("subject") or "")[:120],
            from_addr=_from_addr(msg=m),
            sent_at=(m.get("sentDateTime") or "")[:19],
        )
        side.messages_with_id += 1
        side.by_id.setdefault(rec.message_id, []).append(rec)


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

    overlap_ids = set(src.by_id.keys()) & set(dst.by_id.keys())
    src_msgs_in_overlap = sum(len(src.by_id[i]) for i in overlap_ids)
    dst_msgs_in_overlap = sum(len(dst.by_id[i]) for i in overlap_ids)

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
    })

    log.info(
        "Overlap: {} unique Message-IDs - {} source messages, {} dest messages collide.",
        len(overlap_ids), src_msgs_in_overlap, dst_msgs_in_overlap,
    )

    # Print a sample so the user can sanity-check the matches.
    if overlap_ids and sample_size > 0:
        log.info("Sample (up to {} duplicate Message-IDs):", sample_size)
        for i, mid in enumerate(sorted(overlap_ids)):
            if i >= sample_size:
                break
            src_list = src.by_id[mid]
            dst_list = dst.by_id[mid]
            sample_src = src_list[0]
            sample_dst = dst_list[0]
            subj = sample_src.subject or sample_dst.subject or "(no subject)"
            log.info(
                "  {} | from {!r} | sent {} | src={!r} ({}x) | dst={!r} ({}x) | subject: {}",
                mid, sample_src.from_addr or sample_dst.from_addr,
                sample_src.sent_at or sample_dst.sent_at,
                sample_src.folder_path, len(src_list),
                sample_dst.folder_path, len(dst_list),
                subj,
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
    hdr = f"{'mailbox':<46} {'src_msgs':>10} {'dst_msgs':>10} {'overlap_ids':>12} {'src_dups':>10} {'dst_dups':>10}  status"
    print(hdr)
    print("-" * len(hdr))
    for r in sorted(results, key=lambda x: x["mailbox"]):
        if "error" in r:
            print(f"{r['mailbox']:<46} {'-':>10} {'-':>10} {'-':>12} {'-':>10} {'-':>10}  ERROR: {r['error'][:30]}")
        elif r.get("skipped"):
            print(f"{r['mailbox']:<46} {'-':>10} {'-':>10} {'-':>12} {'-':>10} {'-':>10}  no Imported PST")
        else:
            print(
                f"{r['mailbox']:<46} {r['src_msgs_with_id']:>10} {r['dst_msgs_with_id']:>10} "
                f"{r['overlap_ids']:>12} {r['src_msgs_in_overlap']:>10} {r['dst_msgs_in_overlap']:>10}  ok"
            )
    return 0


if __name__ == "__main__":
    sys.exit(main())
