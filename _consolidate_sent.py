r"""Move every message out of non-canonical 'Sent'-named folders into the
mailbox's actual Sent Items, then delete the now-empty source folders.

Why this exists: earlier runs of _flatten_imported.py only applied the
'Sent' -> Sent Items well-known rule at the top level of 'Imported PST'.
PSTs that nested their sent folder under Inbox (e.g. Imported PST/Inbox/Sent)
ended up creating an Inbox/Sent subfolder in the live mailbox alongside the
real Sent Items, splitting outgoing mail across two folders. The flatten
script was patched to apply that rule recursively on subsequent runs, but
mailboxes that were processed before the patch are stuck with the misplaced
folder. This script consolidates them.

Detection rule (intentionally narrow):

  - Walk every folder in the mailbox EXCEPT the 'Imported PST' subtree
    (handled by _flatten_imported.py) and the 'Deleted Items' subtree
    (the user's trash; leave it alone).
  - For each folder whose displayName is one of the well-known sent
    aliases ('Sent', 'Sent Items', 'Sentitems', 'Sent Mail',
    'Sent Messages')...
  - ...EXCEPT the actual well-known 'sentitems' folder itself...
  - ...move every message in it (and its subtree) into the live
    Sent Items folder, then delete the (now-empty) source.

User-created folders with arbitrary 'Sent <something>' names (e.g.
'Sent To Auditor', 'Sent - Q3 Reports') are NOT touched.

Read-only by default (--dry-run). All GETs + POST /move + DELETE only.

Run from Graph-Migrate/:

  .\.venv\Scripts\python.exe _consolidate_sent.py -c config.toml --mailbox UPN --dry-run
  .\.venv\Scripts\python.exe _consolidate_sent.py -c config.toml --mailbox UPN
  .\.venv\Scripts\python.exe _consolidate_sent.py -c config.toml -m mapping.csv
"""
from __future__ import annotations

import argparse
import os
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote

from loguru import logger

from jtet_pstmigrate.auth import AppPool
from jtet_pstmigrate.config import AppConfig, expand_user_paths
from jtet_pstmigrate.graph_client import GraphClient, GraphError
from jtet_pstmigrate.orchestrator import load_mapping

PAGE_SIZE = 100

# Top-level folder names whose entire subtree we deliberately ignore:
#  - Imported PST: flatten handles those Sent variants via WELL_KNOWN_RULES
#    (recursive since 53a66a0). Doing it here too would race with flatten
#    and pull messages out of folders the user is mid-migrating.
#  - Deleted Items: that's the user's trash. If they soft-deleted Imported PST
#    or some old Sent-named folder, leave it where it is.
SKIP_TOP_LEVEL = frozenset({"Imported PST", "Deleted Items"})

# Known aliases for the canonical Sent Items folder. See the matching
# comment in _flatten_imported.py for why this is a closed set rather than
# a "Sent <anything>" prefix match.
_SENT_VARIANT_NAMES = frozenset({
    "sent",
    "sent items",
    "sentitems",
    "sent mail",
    "sent messages",
})


def _is_sent_variant(name: str) -> bool:
    return (name or "").strip().lower() in _SENT_VARIANT_NAMES


def _strip_base(url: str) -> str:
    base = "https://graph.microsoft.com/v1.0"
    return url[len(base):] if url.startswith(base) else url


def list_child_folders(graph: GraphClient, mailbox: str, parent_id: str) -> list[dict]:
    out: list[dict] = []
    path = (
        f"/users/{quote(mailbox)}/mailFolders/{parent_id}/childFolders"
        f"?$top={PAGE_SIZE}&$select=id,displayName,childFolderCount,totalItemCount"
    )
    while path:
        resp = graph.get(path, expect_status=(200,))
        body = resp.json()
        out.extend(body.get("value", []))
        next_link = body.get("@odata.nextLink")
        path = _strip_base(next_link) if next_link else None
    return out


def list_message_page(graph: GraphClient, mailbox: str, folder_id: str) -> list[dict]:
    path = (
        f"/users/{quote(mailbox)}/mailFolders/{folder_id}/messages"
        f"?$top={PAGE_SIZE}&$select=id"
    )
    resp = graph.get(path, expect_status=(200,))
    return resp.json().get("value", [])


def snapshot_all_message_ids(
    graph: GraphClient,
    mailbox: str,
    folder_id: str,
) -> list[dict]:
    """Page through every (visible) message in the folder and return
    [{id, subject}, ...].

    Snapshotting up front (instead of re-listing after each move) avoids two
    classes of bug:
      - pagination drift, where moving items shifts the page window and we
        skip or revisit messages
      - infinite loops, where a silently-failing move would leave the same
        item visible on page 1 forever

    Note on residual items: Graph v1.0's /messages collection only exposes
    IPM.Note items in the regular table. It does NOT expose Folder Associated
    Information (FAI) items -- folder views, custom rules, search criteria,
    forms -- which PST imports commonly carry. Those still count toward
    mailFolder.totalItemCount, so after this drain the source folder may
    look non-empty by that metric. delete_folder_if_drained handles that
    case: as long as no visible messages remain, the FAI residual goes away
    with the folder when we delete it.
    """
    path = (
        f"/users/{quote(mailbox)}/mailFolders/{folder_id}/messages"
        f"?$top={PAGE_SIZE}&$select=id,subject"
    )
    out: list[dict] = []
    while path:
        resp = graph.get(path, expect_status=(200,))
        body = resp.json()
        out.extend(body.get("value", []))
        next_link = body.get("@odata.nextLink")
        path = _strip_base(next_link) if next_link else None
    return out


def get_folder(graph: GraphClient, mailbox: str, folder_id: str) -> dict:
    path = (
        f"/users/{quote(mailbox)}/mailFolders/{folder_id}"
        f"?$select=id,displayName,childFolderCount,totalItemCount"
    )
    resp = graph.get(path, expect_status=(200,))
    return resp.json()


def get_well_known_folder(graph: GraphClient, mailbox: str, name: str) -> dict | None:
    path = f"/users/{quote(mailbox)}/mailFolders/{name}?$select=id,displayName"
    try:
        resp = graph.get(path, expect_status=(200,))
        return resp.json()
    except GraphError as e:
        if e.status == 404:
            return None
        raise


def move_message(graph: GraphClient, mailbox: str, msg_id: str, dest_id: str) -> None:
    path = f"/users/{quote(mailbox)}/messages/{msg_id}/move"
    graph.post(path, json={"destinationId": dest_id}, expect_status=(200, 201))


def delete_folder_if_drained(graph: GraphClient, mailbox: str, folder: dict, log) -> bool:
    """Delete the folder if it's effectively empty for the user.

    Refuses to delete if:
      - any subfolders remain, OR
      - any visible message remains in the folder

    Allows deletion when totalItemCount > 0 only if GET /messages returns
    nothing. In that case the residual items are FAI (folder-associated
    information) entries -- views, rules, custom search criteria, forms --
    that Graph v1.0 doesn't enumerate or move individually. They go away
    when the folder itself is deleted.
    """
    fresh = get_folder(graph, mailbox, folder["id"])
    items = int(fresh.get("totalItemCount") or 0)
    subs = int(fresh.get("childFolderCount") or 0)
    name = fresh.get("displayName")

    if subs > 0:
        log.warning("Refusing to delete {!r}: still has {} subfolder(s).", name, subs)
        return False

    if items > 0:
        # Verify those items aren't visible messages we somehow missed.
        visible = list_message_page(graph, mailbox, folder["id"])
        if visible:
            log.warning(
                "Refusing to delete {!r}: still has {} visible message(s).",
                name, len(visible),
            )
            return False
        log.info(
            "{!r}: {} residual FAI item(s) (rules/views/config) will be removed with the folder.",
            name, items,
        )

    try:
        graph.delete(
            f"/users/{quote(mailbox)}/mailFolders/{folder['id']}",
            expect_status=(204,),
        )
        return True
    except GraphError as e:
        log.warning("Could not delete {!r} ({}); leaving it.", name, e.status)
        return False


def find_misplaced_sent_folders(
    graph: GraphClient,
    mailbox: str,
    sent_items_id: str,
    *,
    skip_top_level: frozenset[str] = SKIP_TOP_LEVEL,
) -> list[tuple[dict, str]]:
    """Return [(folder, breadcrumb_path)] for every folder in the mailbox
    whose name matches the sent-variant predicate, EXCEPT the canonical
    Sent Items folder itself.

    Top-level folders whose displayName is in `skip_top_level` are not
    descended into at all -- by default that excludes 'Imported PST'
    (handled by flatten) and 'Deleted Items' (the user's trash).
    """
    found: list[tuple[dict, str]] = []

    def walk(folder: dict, prefix: str) -> None:
        path = f"{prefix}/{folder['displayName']}" if prefix else folder["displayName"]
        if folder["id"] != sent_items_id and _is_sent_variant(folder["displayName"]):
            found.append((folder, path))
        if int(folder.get("childFolderCount") or 0) > 0:
            for sub in list_child_folders(graph, mailbox, folder["id"]):
                walk(sub, path)

    for top in list_child_folders(graph, mailbox, "msgFolderRoot"):
        if top.get("displayName") in skip_top_level:
            continue
        walk(top, "")
    return found


def consolidate_subtree_into_sent(
    graph: GraphClient,
    mailbox: str,
    folder: dict,
    sent_items_id: str,
    log,
    *,
    src_path: str,
    dry_run: bool,
) -> int:
    """Move every message in `folder` and its descendants into Sent Items.
    Returns the number of messages moved (0 in dry-run)."""
    moved = 0

    if not dry_run:
        # Snapshot all visible messages, then move. (Pagination-while-mutating
        # is fragile; see snapshot_all_message_ids docstring.) FAI residual
        # is handled at folder-delete time by delete_folder_if_drained.
        for m in snapshot_all_message_ids(graph, mailbox, folder["id"]):
            try:
                move_message(graph, mailbox, m["id"], sent_items_id)
                moved += 1
                if moved % 200 == 0:
                    log.info("  {!r}: moved {} messages so far", src_path, moved)
            except GraphError as e:
                subj = (m.get("subject") or "(no subject)")[:80]
                log.warning(
                    "  {!r}: could not move item {!r} ({}); leaving in place.",
                    src_path, subj, e.status,
                )
    else:
        moved += int(folder.get("totalItemCount") or 0)

    # Recurse into children. With the sent rule applied recursively, any
    # nested folder that ALSO matches a Sent variant gets folded too;
    # everything else gets folded into Sent Items as well (the user has
    # already declared this whole subtree as misplaced sent mail).
    if int(folder.get("childFolderCount") or 0) > 0:
        for sub in list_child_folders(graph, mailbox, folder["id"]):
            sub_path = f"{src_path}/{sub['displayName']}"
            moved += consolidate_subtree_into_sent(
                graph, mailbox, sub, sent_items_id, log,
                src_path=sub_path, dry_run=dry_run,
            )
            if not dry_run:
                delete_folder_if_drained(graph, mailbox, sub, log)
    return moved


def consolidate_mailbox(graph: GraphClient, mailbox: str, *, dry_run: bool) -> dict:
    log = logger.bind(ctx=f"sent[{mailbox}]")
    stats = {"mailbox": mailbox, "folders_found": 0, "messages_moved": 0, "skipped": False}

    sent_items = get_well_known_folder(graph, mailbox, "sentitems")
    if sent_items is None:
        log.warning("Couldn't resolve well-known 'sentitems' folder; skipping mailbox.")
        stats["skipped"] = True
        return stats

    misplaced = find_misplaced_sent_folders(graph, mailbox, sent_items["id"])
    if not misplaced:
        log.info("No misplaced Sent folders found.")
        return stats

    log.info("Found {} misplaced Sent folder(s):", len(misplaced))
    for folder, path in misplaced:
        items = int(folder.get("totalItemCount") or 0)
        subs = int(folder.get("childFolderCount") or 0)
        log.info("  {!r}  ({} direct items, {} subfolders)", path, items, subs)
    stats["folders_found"] = len(misplaced)

    if dry_run:
        est = 0
        for folder, path in misplaced:
            est += consolidate_subtree_into_sent(
                graph, mailbox, folder, sent_items["id"], log,
                src_path=path, dry_run=True,
            )
        log.info("DRY RUN -- would move ~{} messages into Sent Items, then delete the source folders.", est)
        stats["messages_moved"] = est
        return stats

    total_moved = 0
    for folder, path in misplaced:
        log.info("Consolidating {!r} -> Sent Items", path)
        moved = consolidate_subtree_into_sent(
            graph, mailbox, folder, sent_items["id"], log,
            src_path=path, dry_run=False,
        )
        total_moved += moved
        delete_folder_if_drained(graph, mailbox, folder, log)
        log.info("  -> moved {} messages from {!r}", moved, path)

    log.info("Done. {} messages moved into Sent Items across {} folder(s).", total_moved, len(misplaced))
    stats["messages_moved"] = total_moved
    return stats


def _write_jtet_live_status(**fields: str) -> None:
    path = os.environ.get("JTET_LIVE_STATUS_FILE")
    if not path:
        return
    lines = [f"{k}: {v}" for k, v in fields.items()]
    lines.append(
        f"updated_utc: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%SZ')}"
    )
    try:
        Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8")
    except OSError:
        pass


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Move misplaced 'Sent'-named folders into the actual Sent Items.",
    )
    ap.add_argument("-c", "--config", required=True)
    ap.add_argument("-m", "--mapping", help="mapping.csv path")
    ap.add_argument("--mailbox", action="append", help="UPN of a single mailbox (repeatable)")
    ap.add_argument("--dry-run", action="store_true", help="Walk + report, no moves or deletes.")
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

    if ns.dry_run:
        logger.bind(ctx="sent").warning("DRY RUN -- nothing will be moved or deleted.")

    pool = AppPool(cfg.apps)
    parallelism = min(cfg.migration.max_parallel_mailboxes, len(mailboxes)) or 1

    n_mb = len(mailboxes)
    logger.bind(ctx="sent").info(
        "Processing {} mailbox(es) with {} parallel workers, {} app(s).",
        n_mb, parallelism, len(cfg.apps),
    )
    logger.bind(ctx="sent").info(
        "STEP 2 PROGRESS: watch for  PROGRESS: k/{} mailboxes complete  (fires when each mailbox's consolidate finishes).",
        n_mb,
    )

    _write_jtet_live_status(
        pipeline_step="2 of 3",
        run_step="consolidate Sent Items (Graph)",
        mailboxes_total=str(n_mb),
        mailboxes_finished_in_this_step=f"0/{n_mb}",
    )

    results: list[dict] = []
    with GraphClient(pool, cfg.throttle) as graph:
        with ThreadPoolExecutor(max_workers=parallelism, thread_name_prefix="sent") as ex:
            futs = {
                ex.submit(consolidate_mailbox, graph, m, dry_run=ns.dry_run): m
                for m in mailboxes
            }
            done = 0
            for fut in as_completed(futs):
                m = futs[fut]
                try:
                    results.append(fut.result())
                    done += 1
                    logger.bind(ctx="sent").info(
                        "PROGRESS: {}/{} mailboxes complete (finished: {}, ok).",
                        done, n_mb, m,
                    )
                    _write_jtet_live_status(
                        pipeline_step="2 of 3",
                        run_step="consolidate Sent Items (Graph)",
                        mailboxes_total=str(n_mb),
                        mailboxes_finished_in_this_step=f"{done}/{n_mb}",
                        last_mailbox_just_completed=m,
                        not_finished_yet_in_this_step=str(n_mb - done),
                    )
                except GraphError as e:
                    logger.bind(ctx=f"sent[{m}]").error("Graph error: {}", e)
                    results.append({"mailbox": m, "error": str(e)})
                    done += 1
                    logger.bind(ctx="sent").warning(
                        "PROGRESS: {}/{} mailboxes complete (finished: {}, Graph error).",
                        done, n_mb, m,
                    )
                    _write_jtet_live_status(
                        pipeline_step="2 of 3",
                        run_step="consolidate Sent Items",
                        mailboxes_finished_in_this_step=f"{done}/{n_mb}",
                        last_mailbox_just_completed=m,
                        last_result="Graph error",
                    )
                except Exception as e:
                    logger.bind(ctx=f"sent[{m}]").exception("Crashed: {}", e)
                    results.append({"mailbox": m, "error": str(e)})
                    done += 1
                    logger.bind(ctx="sent").warning(
                        "PROGRESS: {}/{} mailboxes complete (finished: {}, ERROR).",
                        done, n_mb, m,
                    )
                    _write_jtet_live_status(
                        pipeline_step="2 of 3",
                        run_step="consolidate Sent Items",
                        mailboxes_finished_in_this_step=f"{done}/{n_mb}",
                        last_mailbox_just_completed=m,
                        last_result="ERROR",
                    )

    _write_jtet_live_status(
        pipeline_step="2 of 3 (complete)",
        run_step="consolidate Sent Items",
        status="Step 2 finished for this process.",
    )

    print()
    hdr = f"{'mailbox':<46} {'folders':>8} {'messages':>10}  status"
    print(hdr)
    print("-" * len(hdr))
    for r in sorted(results, key=lambda x: x["mailbox"]):
        if "error" in r:
            print(f"{r['mailbox']:<46} {'-':>8} {'-':>10}  ERROR: {r['error'][:30]}")
        elif r.get("skipped"):
            print(f"{r['mailbox']:<46} {'-':>8} {'-':>10}  no Sent Items folder")
        else:
            print(f"{r['mailbox']:<46} {r['folders_found']:>8} {r['messages_moved']:>10}  ok")
    return 0


if __name__ == "__main__":
    sys.exit(main())
