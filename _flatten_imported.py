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

When --skip-duplicates is set, the script first scans the destination
mailbox (everything except 'Imported PST') and builds a fuzzy-key index
of (normalized subject, sent-minute, from address). Any source message
whose key is already present in that index is LEFT in 'Imported PST'
instead of being moved. This prevents re-introducing duplicates when
some of the mailbox content was placed there by a separate importer
that may have rewritten Message-IDs (run _audit_duplicates.py first to
diagnose). With this flag, wholesale folder-moves are disabled - every
source message gets an individual duplicate check.

Run from Graph-Migrate/:

  .\.venv\Scripts\python.exe _flatten_imported.py -c config.toml -m mapping.csv
  .\.venv\Scripts\python.exe _flatten_imported.py -c config.toml --mailbox johnd@jteatono365.onmicrosoft.com
  .\.venv\Scripts\python.exe _flatten_imported.py -c config.toml -m mapping.csv --dry-run
  .\.venv\Scripts\python.exe _flatten_imported.py -c config.toml --mailbox johnd@... --skip-duplicates --dry-run
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
PAGE_SIZE = 100

# ----------------------------------------------------------- duplicate detection
# We dedupe by a "fuzzy key": (normalized subject, sent-minute, from address).
# This catches duplicates even when the SMTP Message-ID was rewritten by
# whatever tool did the prior import (Exchange Online minted server IDs of
# the form <by3pr...@by3pr....prod.outlook.com> instead of preserving the
# original PST headers). See _audit_duplicates.py for the diagnosis.

FuzzyKey = tuple[str, str, str]  # (subject_norm, sent_at_minute, from_addr_lower)

_REPLY_PREFIX_RE = re.compile(r"^\s*(?:re|fw|fwd|aw|sv|tr|wg|antwort|antw)\s*[:\[\(]?\s*", re.IGNORECASE)


def _norm_subject(s: str) -> str:
    s = (s or "").strip()
    while True:
        new = _REPLY_PREFIX_RE.sub("", s, count=1)
        if new == s:
            break
        s = new
    return " ".join(s.split()).lower()


def _from_addr(msg: dict) -> str:
    f = msg.get("from") or {}
    eb = f.get("emailAddress") or {}
    return (eb.get("address") or eb.get("name") or "").lower()


def fuzzy_key(msg: dict) -> FuzzyKey | None:
    """Build the dedup key for a message dict, or None if there isn't enough
    signal to compare (no sent timestamp, or both subject + from missing)."""
    subj = _norm_subject(msg.get("subject") or "")
    sent_minute = (msg.get("sentDateTime") or "")[:16]
    addr = _from_addr(msg)
    if not sent_minute:
        return None
    if not subj and not addr:
        return None
    return (subj, sent_minute, addr)


@dataclass(slots=True)
class DedupIndex:
    """Set of fuzzy keys already present in the destination mailbox."""
    keys: set[FuzzyKey] = field(default_factory=set)
    skipped: int = 0  # running count of source messages skipped as duplicates

    def has(self, k: FuzzyKey | None) -> bool:
        return k is not None and k in self.keys

    def add(self, k: FuzzyKey | None) -> None:
        if k is not None:
            self.keys.add(k)

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


def list_message_page(
    graph: GraphClient,
    mailbox: str,
    folder_id: str,
    *,
    with_dedup_fields: bool = False,
) -> list[dict]:
    """Return one page of messages from a folder.

    `with_dedup_fields` controls whether subject/from/sentDateTime are
    fetched too. Off by default to keep response sizes small in the
    common (no-dedup) path.
    """
    select = "id,subject,sentDateTime,from" if with_dedup_fields else "id"
    path = (
        f"/users/{quote(mailbox)}/mailFolders/{folder_id}/messages"
        f"?$top={PAGE_SIZE}&$select={select}"
    )
    resp = graph.get(path, expect_status=(200,))
    return resp.json().get("value", [])


def iter_all_messages(graph: GraphClient, mailbox: str, folder_id: str):
    """Yield every message in a folder (paged), with dedup fields populated."""
    path = (
        f"/users/{quote(mailbox)}/mailFolders/{folder_id}/messages"
        f"?$top=999&$select=id,subject,sentDateTime,from"
    )
    while path:
        resp = graph.get(path, expect_status=(200,))
        body = resp.json()
        for m in body.get("value", []):
            yield m
        next_link = body.get("@odata.nextLink")
        path = _strip_base(next_link) if next_link else None


def move_message(graph: GraphClient, mailbox: str, msg_id: str, dest_id: str) -> None:
    path = f"/users/{quote(mailbox)}/messages/{msg_id}/move"
    graph.post(path, json={"destinationId": dest_id}, expect_status=(200, 201))


def move_folder(graph: GraphClient, mailbox: str, folder_id: str, dest_parent_id: str) -> None:
    path = f"/users/{quote(mailbox)}/mailFolders/{folder_id}/move"
    graph.post(path, json={"destinationId": dest_parent_id}, expect_status=(200, 201))


def create_subfolder(graph: GraphClient, mailbox: str, parent_id: str, name: str) -> dict:
    """Create a child folder under `parent_id`. Returns the new folder dict.

    Used in --skip-duplicates mode where we can't fall back to wholesale
    folder-moves (we need to walk every message). If a sibling with the
    same name already exists, returns it instead of creating a duplicate.
    """
    existing = find_child_named(graph, mailbox, parent_id, name)
    if existing:
        return existing
    path = f"/users/{quote(mailbox)}/mailFolders/{parent_id}/childFolders"
    resp = graph.post(path, json={"displayName": name}, expect_status=(201,))
    return resp.json()


def build_dedup_index(graph: GraphClient, mailbox: str, *, skip_root_id: str, log) -> DedupIndex:
    """Walk every folder under `msgFolderRoot` except `skip_root_id` and
    collect fuzzy keys for every message. Returns a DedupIndex.
    """
    idx = DedupIndex()
    folders_walked = 0
    msgs_indexed = 0

    def _walk(folder: dict) -> None:
        nonlocal folders_walked, msgs_indexed
        folders_walked += 1
        for m in iter_all_messages(graph, mailbox, folder["id"]):
            k = fuzzy_key(m)
            if k:
                idx.keys.add(k)
                msgs_indexed += 1
        if int(folder.get("childFolderCount") or 0) > 0:
            for sub in list_child_folders(graph, mailbox, folder["id"]):
                _walk(sub)

    for top in list_child_folders(graph, mailbox, "msgFolderRoot"):
        if top["id"] == skip_root_id:
            continue
        _walk(top)

    log.info(
        "Dedup index: {} fuzzy keys from {} folders ({} messages indexed).",
        len(idx.keys), folders_walked, msgs_indexed,
    )
    return idx


def get_folder(graph: GraphClient, mailbox: str, folder_id: str) -> dict:
    """Fetch a single folder by id with the fields we care about."""
    path = (
        f"/users/{quote(mailbox)}/mailFolders/{folder_id}"
        f"?$select=id,displayName,childFolderCount,totalItemCount"
    )
    resp = graph.get(path, expect_status=(200,))
    return resp.json()


def delete_folder_if_empty(graph: GraphClient, mailbox: str, folder_id: str) -> bool:
    """Refetch the folder and only DELETE if it's truly empty.

    Why this exists: per Microsoft's docs, DELETE on a non-empty mailFolder
    succeeds and sends contents to Deleted Items. That's recoverable but not
    desirable - if we ever reach a delete step with a folder that still has
    content (orphan messages, a partially-failed merge, etc.), we'd rather
    leave it in place and let the user see the leftover than soft-delete it.

    Returns True if the folder was deleted, False if it was preserved.
    """
    log = logger.bind(ctx=f"flatten[{mailbox}]")
    try:
        fresh = get_folder(graph, mailbox, folder_id)
    except GraphError as e:
        log.warning("Could not refetch folder {} pre-delete ({}); skipping delete.", folder_id, e.status)
        return False

    items = int(fresh.get("totalItemCount") or 0)
    subs = int(fresh.get("childFolderCount") or 0)
    if items > 0 or subs > 0:
        log.warning(
            "Refusing to delete {!r}: still has {} items, {} subfolders. Re-run to clean up.",
            fresh.get("displayName"), items, subs,
        )
        return False

    path = f"/users/{quote(mailbox)}/mailFolders/{folder_id}"
    try:
        graph.delete(path, expect_status=(204,))
        return True
    except GraphError as e:
        log.warning("Could not delete folder {} ({}); leaving it.", folder_id, e.status)
        return False


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


def recursive_item_count(graph: GraphClient, mailbox: str, folder: dict) -> int:
    """Sum totalItemCount across `folder` and all its descendants.

    Graph's totalItemCount only counts items directly inside a folder; to
    get a true subtree size we have to walk. Cost: one GET per non-empty
    branch. Used for accurate dry-run totals on wholesale folder-moves.
    """
    total = int(folder.get("totalItemCount") or 0)
    if int(folder.get("childFolderCount") or 0) > 0:
        for sub in list_child_folders(graph, mailbox, folder["id"]):
            total += recursive_item_count(graph, mailbox, sub)
    return total


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
            if sub_subs > 0:
                deep = recursive_item_count(graph, mailbox, sub)
                log.info(
                    "{}[move]   {!r} ({} direct items, {} subfolders, {} total in subtree) -> {!r} (new, single folder-move)",
                    indent, sub["displayName"], sub_items, sub_subs, deep, sub_path,
                )
            else:
                deep = sub_items
                log.info(
                    "{}[move]   {!r} ({} items) -> {!r} (new, single folder-move)",
                    indent, sub["displayName"], sub_items, sub_path,
                )
            msgs += deep
            folders += 1
    return msgs, folders


def execute_merge(
    graph: GraphClient,
    mailbox: str,
    src: dict,
    dst_id: str,
    *,
    src_path: str,
    log,
    depth: int,
    dedup: DedupIndex | None = None,
) -> tuple[int, int]:
    """Actually move messages + subfolders from src into dst, recursively.

    Pre: dst exists. Post: src is empty IF every message was moved (the
    caller decides whether to delete it).

    `src_path` is the human-readable source path (e.g. 'Imported PST/Inbox')
    used only for progress logging so the user can see which merge each
    progress tick belongs to.

    `dedup`, if provided, is checked for every message before moving. Any
    source message whose fuzzy key already exists in the destination is
    skipped (left in `src`). Skipped count goes into dedup.skipped.

    With dedup enabled, wholesale folder-moves are disabled: we create a
    matching folder under dst when one doesn't exist, then merge into it
    item-by-item, so every message gets a duplicate check.

    Returns (messages_moved_individually, subfolders_processed).
    """
    indent = "  " * depth
    src_id = src["id"]
    moved = 0

    # ------ messages directly in `src` ------
    while True:
        msgs = list_message_page(graph, mailbox, src_id, with_dedup_fields=dedup is not None)
        if not msgs:
            break
        # Track messages we *visit* this page; if dedup leaves them all in
        # place, listing the same page again would loop forever, so when
        # dedup is on we page using $skip-style by tracking processed ids.
        all_skipped_this_page = True
        for m in msgs:
            if dedup is not None:
                k = fuzzy_key(m)
                if dedup.has(k):
                    dedup.skipped += 1
                    continue
                # Move it, and remember its key so duplicates among the
                # imported messages themselves don't all flood through.
                move_message(graph, mailbox, m["id"], dst_id)
                dedup.add(k)
                moved += 1
                all_skipped_this_page = False
            else:
                move_message(graph, mailbox, m["id"], dst_id)
                moved += 1
            if moved and moved % 200 == 0:
                log.info("{}{!r}: moved {} messages so far", indent, src_path, moved)
        if dedup is not None and all_skipped_this_page:
            # Every message in this page is staying put. Re-listing would
            # return the same items forever; break out and let them stay
            # in the source folder for the user to review.
            break

    # ------ subfolders ------
    subs = 0
    for sub in list_child_folders(graph, mailbox, src_id):
        existing = find_child_named(graph, mailbox, dst_id, sub["displayName"])
        sub_path = f"{src_path}/{sub['displayName']}"
        if existing:
            sm, ss = execute_merge(
                graph, mailbox, sub, existing["id"],
                src_path=sub_path, log=log, depth=depth + 1, dedup=dedup,
            )
            delete_folder_if_empty(graph, mailbox, sub["id"])
            moved += sm
            subs += ss + 1
        elif dedup is not None:
            # Can't wholesale-move when dedup is on: create the folder and
            # merge into it so each message is checked.
            new_dst = create_subfolder(graph, mailbox, dst_id, sub["displayName"])
            sm, ss = execute_merge(
                graph, mailbox, sub, new_dst["id"],
                src_path=sub_path, log=log, depth=depth + 1, dedup=dedup,
            )
            delete_folder_if_empty(graph, mailbox, sub["id"])
            moved += sm
            subs += ss + 1
        else:
            move_folder(graph, mailbox, sub["id"], dst_id)
            subs += 1

    return moved, subs


def preview_dedup_skips(
    graph: GraphClient,
    mailbox: str,
    folder: dict,
    dedup: DedupIndex,
) -> int:
    """Walk `folder` and all descendants, count how many messages would be
    skipped by dedup. Pure GETs, no writes. Used in dry-run mode."""
    skipped = 0
    for m in iter_all_messages(graph, mailbox, folder["id"]):
        if dedup.has(fuzzy_key(m)):
            skipped += 1
    if int(folder.get("childFolderCount") or 0) > 0:
        for sub in list_child_folders(graph, mailbox, folder["id"]):
            skipped += preview_dedup_skips(graph, mailbox, sub, dedup)
    return skipped


def flatten_mailbox(
    graph: GraphClient,
    mailbox: str,
    *,
    dry_run: bool,
    skip_duplicates: bool = False,
) -> dict:
    log = logger.bind(ctx=f"flatten[{mailbox}]")
    stats = {
        "mailbox": mailbox,
        "messages_moved": 0,
        "folders_processed": 0,
        "duplicates_skipped": 0,
        "skipped": False,
    }

    imported = find_imported_root(graph, mailbox)
    if imported is None:
        log.info("No '{}' folder found; nothing to do.", ROOT_FOLDER_NAME)
        stats["skipped"] = True
        return stats

    # Build dedup index up front. Done before any moves so we capture the
    # destination state before we start modifying it.
    dedup: DedupIndex | None = None
    if skip_duplicates:
        log.info("Building dedup index (walking destination mailbox)...")
        dedup = build_dedup_index(graph, mailbox, skip_root_id=imported["id"], log=log)

    root_direct_items = int(imported.get("totalItemCount") or 0)
    log.info(
        "Found '{}' ({} direct items, {} top-level subfolders)",
        ROOT_FOLDER_NAME,
        root_direct_items,
        imported.get("childFolderCount", 0),
    )

    # Direct items sitting in 'Imported PST' itself (above all subfolders) get
    # moved into the live Inbox -- that's the closest semantic match for "loose"
    # imported messages. We resolve Inbox via the well-known endpoint so it works
    # in non-English tenants too.
    inbox_for_loose: dict | None = None
    if root_direct_items > 0:
        inbox_for_loose = get_well_known_folder(graph, mailbox, "inbox")
        if inbox_for_loose is None:
            log.warning(
                "{} direct items in '{}' but couldn't resolve Inbox; will leave them and skip the final delete.",
                root_direct_items, ROOT_FOLDER_NAME,
            )

    children = list_child_folders(graph, mailbox, imported["id"])
    log.info("Plan:")

    if root_direct_items > 0 and inbox_for_loose is not None:
        log.info(
            "  [move-msgs] {} direct items in '{}' -> 'Inbox' (item-by-item)",
            root_direct_items, ROOT_FOLDER_NAME,
        )

    # Build the plan first (always - for dry-run AND real run, so the user
    # sees the intent before any writes happen).
    actions: list[dict] = []
    plan_msgs = root_direct_items if inbox_for_loose is not None else 0
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
            plan_msgs += sm
            plan_folders += sf
        else:
            if subs > 0:
                deep = recursive_item_count(graph, mailbox, child)
                log.info(
                    "  [move]   {!r} ({} direct items, {} subfolders, {} total in subtree) -> mailbox root (new top-level folder)",
                    name, items, subs, deep,
                )
            else:
                deep = items
                log.info(
                    "  [move]   {!r} ({} items) -> mailbox root (new top-level folder)",
                    name, items,
                )
            actions.append({"kind": "move-root", "src": child})
            plan_msgs += deep
            plan_folders += 1

    log.info(
        "Plan total: ~{} messages affected (item-by-item + folder-move subtrees), {} folder operations.",
        plan_msgs, plan_folders,
    )

    if dedup is not None:
        # In dedup mode we always go item-by-item -- no wholesale moves --
        # so the user shouldn't see "[move] (single folder-move)" promises.
        log.info(
            "--skip-duplicates is on: wholesale folder-moves above will instead be created+merged "
            "item-by-item so every message is checked against the dedup index ({} keys).",
            len(dedup.keys),
        )
        # Dry-run estimate: how many messages would be skipped as duplicates?
        if dry_run:
            log.info("Counting how many source messages match the dedup index...")
            est_skipped = preview_dedup_skips(graph, mailbox, imported, dedup)
            log.info(
                "Estimated duplicates that would be SKIPPED: {} of ~{} source messages.",
                est_skipped, plan_msgs,
            )

    if dry_run:
        log.info("DRY RUN - no changes made.")
        return stats

    # Execute
    log.info("Executing plan...")
    item_moves = 0  # messages moved one at a time (counted exactly)
    folder_moves = 0  # whole-folder moves (each one Graph call, many messages relocated)
    merges_done = 0  # number of folder merges completed (parents of item-moves)

    if root_direct_items > 0 and inbox_for_loose is not None:
        log.info("moving {} direct items from {!r} -> Inbox", root_direct_items, ROOT_FOLDER_NAME)
        moved_loose = 0
        while True:
            msgs = list_message_page(
                graph, mailbox, imported["id"],
                with_dedup_fields=dedup is not None,
            )
            if not msgs:
                break
            all_skipped = True
            for m in msgs:
                if dedup is not None:
                    k = fuzzy_key(m)
                    if dedup.has(k):
                        dedup.skipped += 1
                        continue
                    move_message(graph, mailbox, m["id"], inbox_for_loose["id"])
                    dedup.add(k)
                else:
                    move_message(graph, mailbox, m["id"], inbox_for_loose["id"])
                moved_loose += 1
                all_skipped = False
                if moved_loose % 200 == 0:
                    log.info("  '{}' (loose root items): moved {} so far", ROOT_FOLDER_NAME, moved_loose)
            if dedup is not None and all_skipped:
                break
        item_moves += moved_loose

    for action in actions:
        src = action["src"]
        if action["kind"] == "merge":
            dst = action["dst"]
            src_path = f"{ROOT_FOLDER_NAME}/{src['displayName']}"
            dst_name = dst.get("displayName", "?")
            log.info("merge {!r} -> {!r}", src_path, dst_name)
            m, s = execute_merge(
                graph, mailbox, src, dst["id"],
                src_path=src_path, log=log, depth=2, dedup=dedup,
            )
            delete_folder_if_empty(graph, mailbox, src["id"])
            item_moves += m
            # `s` counts every subfolder the merge touched, including ones that
            # were folder-moved (single API calls). Track the folder-move count
            # separately so the summary can show both.
            folder_moves += s  # subfolder ops (mostly folder-moves)
            merges_done += 1
        else:
            # move-root: hoist a top-level folder under Imported PST to the
            # mailbox root.
            if dedup is None:
                log.info("move {!r} -> mailbox root", src["displayName"])
                move_folder(graph, mailbox, src["id"], "msgFolderRoot")
                folder_moves += 1
            else:
                # With dedup on we can't wholesale-move (no per-message check).
                # Create the folder at root and merge into it item-by-item.
                src_path = f"{ROOT_FOLDER_NAME}/{src['displayName']}"
                log.info(
                    "move {!r} -> mailbox root (item-by-item due to --skip-duplicates)",
                    src["displayName"],
                )
                new_root = create_subfolder(
                    graph, mailbox, "msgFolderRoot", src["displayName"],
                )
                m, s = execute_merge(
                    graph, mailbox, src, new_root["id"],
                    src_path=src_path, log=log, depth=2, dedup=dedup,
                )
                delete_folder_if_empty(graph, mailbox, src["id"])
                item_moves += m
                folder_moves += s + 1

    # Only delete Imported PST if every prior step actually cleared it.
    # delete_folder_if_empty refuses to delete a folder that still has
    # contents (items or subfolders), so leftover data stays visible to
    # the user under 'Imported PST' rather than getting soft-deleted.
    delete_folder_if_empty(graph, mailbox, imported["id"])

    stats["messages_moved"] = item_moves
    stats["folders_processed"] = folder_moves + merges_done
    if dedup is not None:
        stats["duplicates_skipped"] = dedup.skipped
        log.info(
            "Done. {} messages moved individually, {} duplicates skipped, "
            "{} top-level merges, {} folder ops. Plan estimated ~{} total messages affected.",
            item_moves, dedup.skipped, merges_done, folder_moves, plan_msgs,
        )
    else:
        log.info(
            "Done. {} messages moved individually, {} top-level merges, {} folder-move ops "
            "(each carrying an entire subtree). Plan estimated ~{} total messages affected.",
            item_moves, merges_done, folder_moves, plan_msgs,
        )
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
    ap.add_argument(
        "--skip-duplicates",
        action="store_true",
        help=(
            "Before flattening, scan the destination mailbox and build a fuzzy-key "
            "index (subject + sent-minute + from). Any source message whose key is "
            "already present is left in 'Imported PST' instead of being moved. "
            "Disables wholesale folder-moves (every message is checked individually). "
            "Use this when an earlier import populated the mailbox and may overlap "
            "with the PST contents -- see _audit_duplicates.py output."
        ),
    )
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
    if ns.skip_duplicates:
        logger.bind(ctx="flatten").info(
            "--skip-duplicates ON: messages whose (subject, sent-minute, from) match an "
            "existing destination message will be left in Imported PST."
        )

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
                ex.submit(
                    flatten_mailbox, graph, m,
                    dry_run=ns.dry_run, skip_duplicates=ns.skip_duplicates,
                ): m
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
    hdr = f"{'mailbox':<46} {'messages':>10} {'dup_skip':>9} {'folders':>8}  status"
    print(hdr)
    print("-" * len(hdr))
    for s in sorted(all_stats, key=lambda x: x["mailbox"]):
        if "error" in s:
            print(f"{s['mailbox']:<46} {'-':>10} {'-':>9} {'-':>8}  ERROR: {s['error'][:40]}")
        elif s.get("skipped"):
            print(f"{s['mailbox']:<46} {'-':>10} {'-':>9} {'-':>8}  no Imported PST folder")
        else:
            dup = s.get("duplicates_skipped", 0)
            print(
                f"{s['mailbox']:<46} {s['messages_moved']:>10} "
                f"{dup if ns.skip_duplicates else '-':>9} {s['folders_processed']:>8}  ok"
            )
    return 0


if __name__ == "__main__":
    sys.exit(main())
