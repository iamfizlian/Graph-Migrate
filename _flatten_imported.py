r"""Flatten 'Imported PST' folder into the mailbox root.

For each mailbox in mapping.csv (or --mailbox), walks the 'Imported PST'
folder. For each immediate child folder:

  - If a folder with the same displayName exists at the mailbox root
    (e.g. Inbox, Sent Items, Drafts), recursively moves messages and
    subfolders from the imported copy into the existing one, then
    deletes the now-empty imported folder.

  - If no matching folder exists at the mailbox root (e.g. a custom
    'Customers' folder), moves the folder wholesale to the mailbox
    root with a single Graph call.

After all children are processed, deletes the 'Imported PST' container.

Idempotent: if the script is interrupted, re-running picks up wherever
it left off. Already-moved messages and folders are simply absent from
the imported tree on the next pass.

Run from Graph-Migrate/:

  .\.venv\Scripts\python.exe _flatten_imported.py -c config.toml -m mapping.csv
  .\.venv\Scripts\python.exe _flatten_imported.py -c config.toml --mailbox johnd@jteatono365.onmicrosoft.com
  .\.venv\Scripts\python.exe _flatten_imported.py -c config.toml -m mapping.csv --dry-run
"""
from __future__ import annotations

import argparse
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import quote

from loguru import logger

from jtet_pstmigrate.auth import AppPool
from jtet_pstmigrate.config import AppConfig, expand_user_paths
from jtet_pstmigrate.graph_client import GraphClient, GraphError
from jtet_pstmigrate.orchestrator import load_mapping

ROOT_FOLDER_NAME = "Imported PST"
PAGE_SIZE = 100

# Top-level Imported PST folders whose displayName matches one of these
# predicates get merged into the corresponding Graph well-known folder
# (resolved via /users/{m}/mailFolders/{wellKnownName}, which is locale-safe),
# instead of going through exact-name matching at the mailbox root.
#
# This handles PSTs that named the sent folder "Sent" instead of the Outlook
# canonical "Sent Items", and similar drift.  Add more rules as you encounter
# other variants worth normalizing (e.g. "Trash" -> deleteditems).
# Order matters; first match wins.


def _is_sent_variant(name: str) -> bool:
    n = name.strip().lower()
    return n == "sent" or n == "sentitems" or n.startswith("sent ")


WELL_KNOWN_RULES: list[tuple] = [
    (_is_sent_variant, "sentitems"),
]


# ---------------------------------------------------------------------- helpers


def _strip_base(url: str) -> str:
    """Strip Graph base URL from @odata.nextLink to feed back into GraphClient."""
    base = "https://graph.microsoft.com/v1.0"
    if url.startswith(base):
        return url[len(base):]
    return url


def list_child_folders(graph: GraphClient, mailbox: str, parent_id: str) -> list[dict]:
    """Return all immediate child folders of `parent_id`."""
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
    """Return one page of messages from a folder."""
    path = (
        f"/users/{quote(mailbox)}/mailFolders/{folder_id}/messages"
        f"?$top={PAGE_SIZE}&$select=id"
    )
    resp = graph.get(path, expect_status=(200,))
    return resp.json().get("value", [])


def move_message(graph: GraphClient, mailbox: str, msg_id: str, dest_id: str) -> None:
    path = f"/users/{quote(mailbox)}/messages/{msg_id}/move"
    graph.post(path, json={"destinationId": dest_id}, expect_status=(200, 201))


def move_folder(graph: GraphClient, mailbox: str, folder_id: str, dest_parent_id: str) -> None:
    path = f"/users/{quote(mailbox)}/mailFolders/{folder_id}/move"
    graph.post(path, json={"destinationId": dest_parent_id}, expect_status=(200, 201))


def delete_folder(graph: GraphClient, mailbox: str, folder_id: str) -> None:
    path = f"/users/{quote(mailbox)}/mailFolders/{folder_id}"
    try:
        graph.delete(path, expect_status=(204,))
    except GraphError as e:
        logger.bind(ctx=f"flatten[{mailbox}]").warning(
            "Could not delete folder {} ({}). Likely not empty; leaving it.", folder_id, e.status
        )


def find_imported_root(graph: GraphClient, mailbox: str) -> dict | None:
    for c in list_child_folders(graph, mailbox, "msgFolderRoot"):
        if c["displayName"] == ROOT_FOLDER_NAME:
            return c
    return None


def find_child_named(graph: GraphClient, mailbox: str, parent_id: str, name: str) -> dict | None:
    name_lower = name.lower()
    for c in list_child_folders(graph, mailbox, parent_id):
        if c["displayName"].lower() == name_lower:
            return c
    return None


def get_well_known_folder(graph: GraphClient, mailbox: str, well_known_name: str) -> dict | None:
    """Resolve a Graph well-known folder (e.g. 'sentitems', 'inbox') to its
    actual folder object. Locale-safe: works regardless of the user's
    Outlook display language.
    """
    path = f"/users/{quote(mailbox)}/mailFolders/{well_known_name}?$select=id,displayName"
    try:
        resp = graph.get(path, expect_status=(200,))
        return resp.json()
    except GraphError as e:
        if e.status == 404:
            return None
        raise


def resolve_top_level_target(graph: GraphClient, mailbox: str, child_name: str) -> dict | None:
    """Decide which root-level folder a top-level Imported PST child should
    merge into. Returns the target folder dict, or None if no destination
    exists at the mailbox root (caller should move the folder wholesale).
    """
    for predicate, well_known in WELL_KNOWN_RULES:
        if predicate(child_name):
            return get_well_known_folder(graph, mailbox, well_known)
    return find_child_named(graph, mailbox, "msgFolderRoot", child_name)


# -------------------------------------------------------- merge / move / flatten


def plan_merge(
    graph: GraphClient,
    mailbox: str,
    src: dict,
    dst: dict,
    dst_path: str,
    *,
    log,
    depth: int,
) -> tuple[int, int]:
    """Walk subfolders of `src` and log the plan for each, recursively.

    The caller is responsible for logging the src -> dst header line for
    `src` itself; this function only emits lines for sub-entries.

    Returns (planned_messages, planned_folders) where folders counts
    src plus all descendant folders. Pure logging — no side effects
    beyond the GET requests needed to walk the tree.
    """
    indent = "  " * depth
    msgs = int(src.get("totalItemCount") or 0)
    folders = 1  # the src folder itself
    for sub in list_child_folders(graph, mailbox, src["id"]):
        existing = find_child_named(graph, mailbox, dst["id"], sub["displayName"])
        sub_path = f"{dst_path}/{sub['displayName']}"
        sub_items = int(sub.get("totalItemCount") or 0)
        sub_subs = int(sub.get("childFolderCount") or 0)
        if existing:
            log.info(
                "{}[merge]  {!r} ({} items, {} subfolders) -> {!r}",
                indent, sub["displayName"], sub_items, sub_subs, sub_path,
            )
            sm, sf = plan_merge(graph, mailbox, sub, existing, sub_path, log=log, depth=depth + 1)
            msgs += sm
            folders += sf
        else:
            log.info(
                "{}[move]   {!r} ({} items) -> {!r} (new, single folder-move)",
                indent, sub["displayName"], sub_items, sub_path,
            )
            msgs += sub_items
            folders += 1
    return msgs, folders


def execute_merge(
    graph: GraphClient,
    mailbox: str,
    src_id: str,
    dst_id: str,
    *,
    log,
    depth: int,
) -> tuple[int, int]:
    """Actually move messages + subfolders from src into dst, recursively.

    Pre: dst exists. Post: src is empty (caller deletes it).
    Returns (messages_moved, subfolders_processed).
    """
    indent = "  " * depth
    moved = 0
    while True:
        msgs = list_message_page(graph, mailbox, src_id)
        if not msgs:
            break
        for m in msgs:
            move_message(graph, mailbox, m["id"], dst_id)
            moved += 1
            if moved % 200 == 0:
                log.info("{}... moved {} messages so far", indent, moved)

    subs = 0
    for sub in list_child_folders(graph, mailbox, src_id):
        existing = find_child_named(graph, mailbox, dst_id, sub["displayName"])
        if existing:
            sm, ss = execute_merge(graph, mailbox, sub["id"], existing["id"], log=log, depth=depth + 1)
            delete_folder(graph, mailbox, sub["id"])
            moved += sm
            subs += ss + 1
        else:
            move_folder(graph, mailbox, sub["id"], dst_id)
            subs += 1

    return moved, subs


def flatten_mailbox(graph: GraphClient, mailbox: str, *, dry_run: bool) -> dict:
    log = logger.bind(ctx=f"flatten[{mailbox}]")
    stats = {"mailbox": mailbox, "messages_moved": 0, "folders_processed": 0, "skipped": False}

    imported = find_imported_root(graph, mailbox)
    if imported is None:
        log.info("No '{}' folder found; nothing to do.", ROOT_FOLDER_NAME)
        stats["skipped"] = True
        return stats

    log.info(
        "Found '{}' ({} direct items, {} top-level subfolders)",
        ROOT_FOLDER_NAME,
        imported.get("totalItemCount", 0),
        imported.get("childFolderCount", 0),
    )

    children = list_child_folders(graph, mailbox, imported["id"])
    log.info("Plan:")

    # Build the plan first (always — for dry-run AND real run, so the user
    # sees the intent before any writes happen).
    actions: list[dict] = []
    plan_msgs = 0
    plan_folders = 0
    for child in children:
        name = child["displayName"]
        items = int(child.get("totalItemCount") or 0)
        subs = int(child.get("childFolderCount") or 0)
        target = resolve_top_level_target(graph, mailbox, name)
        if target:
            tgt_name = target.get("displayName", "?")
            note = " (well-known rule)" if any(p(name) for p, _ in WELL_KNOWN_RULES) else ""
            log.info("  [merge]  {!r} ({} items, {} subfolders) -> {!r}{}", name, items, subs, tgt_name, note)
            sm, sf = plan_merge(
                graph, mailbox, child, target,
                dst_path=tgt_name, log=log, depth=2,
            )
            actions.append({"kind": "merge", "src": child, "dst": target})
            # plan_merge already counted the top-level child (folders += 1) and its messages,
            # but we logged the top line ourselves; re-add to keep totals consistent.
            plan_msgs += sm
            plan_folders += sf
        else:
            log.info("  [move]   {!r} ({} items, {} subfolders) -> mailbox root (new top-level folder)", name, items, subs)
            actions.append({"kind": "move-root", "src": child})
            plan_msgs += items
            plan_folders += 1 + subs

    log.info(
        "Plan total: ~{} messages to move, {} folder operations.",
        plan_msgs, plan_folders,
    )

    if dry_run:
        log.info("DRY RUN — no changes made.")
        return stats

    # Execute
    log.info("Executing plan...")
    total_msgs = 0
    total_subs = 0
    for action in actions:
        src = action["src"]
        if action["kind"] == "merge":
            dst = action["dst"]
            log.info("merge {!r} -> {!r}", src["displayName"], dst.get("displayName"))
            m, s = execute_merge(graph, mailbox, src["id"], dst["id"], log=log, depth=2)
            delete_folder(graph, mailbox, src["id"])
            total_msgs += m
            total_subs += s + 1
        else:
            log.info("move {!r} -> mailbox root", src["displayName"])
            move_folder(graph, mailbox, src["id"], "msgFolderRoot")
            total_subs += 1

    delete_folder(graph, mailbox, imported["id"])

    stats["messages_moved"] = total_msgs
    stats["folders_processed"] = total_subs
    log.info("Done. Moved {} messages across {} folders.", total_msgs, total_subs)
    return stats


# ----------------------------------------------------------------------- driver


def main() -> int:
    ap = argparse.ArgumentParser(description="Merge 'Imported PST' folders into mailbox root.")
    ap.add_argument("-c", "--config", required=True, help="config.toml path")
    ap.add_argument("-m", "--mapping", help="mapping.csv path (process every mailbox in it)")
    ap.add_argument(
        "--mailbox",
        action="append",
        help="UPN of a single mailbox to process (repeatable). Overrides -m.",
    )
    ap.add_argument("--dry-run", action="store_true", help="Walk the tree and log what WOULD happen, no writes.")
    ns = ap.parse_args()

    cfg = expand_user_paths(AppConfig.load(Path(ns.config)))

    # Resolve mailbox list
    mailboxes: list[str]
    if ns.mailbox:
        mailboxes = sorted(set(ns.mailbox))
    elif ns.mapping:
        rows = load_mapping(Path(ns.mapping))
        mailboxes = sorted({r.target_mailbox for r in rows})
    else:
        ap.error("Provide either --mapping or one or more --mailbox UPNs.")

    logger.remove()
    logger.add(sys.stderr, level="INFO", format="{time:HH:mm:ss} | {level: <7} | {extra[ctx]} | {message}")

    if ns.dry_run:
        logger.bind(ctx="flatten").warning("DRY RUN — no folder/message moves or deletes will happen.")

    pool = AppPool(cfg.apps)
    parallelism = min(cfg.migration.max_parallel_mailboxes, len(mailboxes)) or 1

    logger.bind(ctx="flatten").info(
        "Processing {} mailbox(es) with {} parallel workers and {} app(s).",
        len(mailboxes), parallelism, len(cfg.apps),
    )

    all_stats: list[dict] = []
    with GraphClient(pool, cfg.throttle) as graph:
        with ThreadPoolExecutor(max_workers=parallelism, thread_name_prefix="flatten") as ex:
            futures = {
                ex.submit(flatten_mailbox, graph, m, dry_run=ns.dry_run): m
                for m in mailboxes
            }
            for fut in as_completed(futures):
                m = futures[fut]
                try:
                    all_stats.append(fut.result())
                except Exception as e:
                    logger.bind(ctx=f"flatten[{m}]").exception("Crashed: {}", e)
                    all_stats.append({"mailbox": m, "error": str(e)})

    # Summary
    print()
    print(f"{'mailbox':<46} {'messages':>10} {'folders':>8}  status")
    print("-" * 80)
    for s in sorted(all_stats, key=lambda x: x["mailbox"]):
        if "error" in s:
            print(f"{s['mailbox']:<46} {'-':>10} {'-':>8}  ERROR: {s['error'][:40]}")
        elif s.get("skipped"):
            print(f"{s['mailbox']:<46} {'-':>10} {'-':>8}  no Imported PST folder")
        else:
            print(f"{s['mailbox']:<46} {s['messages_moved']:>10} {s['folders_processed']:>8}  ok")
    return 0


if __name__ == "__main__":
    sys.exit(main())
