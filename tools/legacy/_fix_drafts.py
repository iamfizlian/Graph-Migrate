r"""Clear the isDraft flag on every imported message that's stuck as a draft.

Why this exists
---------------
Earlier versions of uploader.py created messages by POST-ing raw MIME to
``/users/{id}/messages``, which lands them in Drafts with isDraft=true.
The follow-up ``/move`` correctly relocated them to Inbox / Sent Items / etc.
but did NOT clear the underlying MAPI MSGFLAG_UNSENT bit. The result is
that Outlook on the web shows every imported message as a draft, displays
'Draft' badges on conversations, and loses date-sort.

This script walks every folder in a mailbox and PATCHes
PR_MESSAGE_FLAGS = 0x01 (MSGFLAG_READ; clears UNSENT and SUBMIT) on
every message Graph still reports as isDraft=true.

Scope
-----
By default the walk skips Drafts itself (real drafts the user wrote should
stay drafts) and the Deleted Items subtree (their state doesn't matter).
The Imported PST subtree IS walked, so this can be run before or after
flatten/consolidate without coordinating order.

Idempotent: re-running on a clean mailbox finds 0 matches.

Read-only with --dry-run; counts and reports per-folder how many drafts
would be fixed without making any changes.

Run from Graph-Migrate/:

  .\.venv\Scripts\python.exe _fix_drafts.py -c config.toml --mailbox UPN --dry-run
  .\.venv\Scripts\python.exe _fix_drafts.py -c config.toml --mailbox UPN
  .\.venv\Scripts\python.exe _fix_drafts.py -c config.toml -m mapping.csv
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

PAGE_SIZE = 100

# Top-level folders we deliberately do not descend into:
#  - Drafts: real drafts the user authored should stay drafts
#  - Deleted Items: trash; their state doesn't matter for the user
SKIP_TOP_LEVEL = frozenset({"Drafts", "Deleted Items"})

# PR_MESSAGE_FLAGS, PT_LONG.  Bits we care about:
#   MSGFLAG_READ      0x01  - has been read
#   MSGFLAG_UNMODIFIED 0x02 - no pending edits
#   MSGFLAG_SUBMIT    0x04  - was successfully submitted (i.e. NOT a draft)
#   MSGFLAG_UNSENT    0x08  - is a draft awaiting submission  <-- WANT 0
# We set 0x05 (READ | SUBMIT) so the message looks like a delivered/sent
# item. Setting just 0x01 (the previous attempt) cleared UNSENT and READ
# applied, but Graph kept reporting isDraft=true on at least some tenants
# -- it appears to also consider PR_SUBMIT_FLAGS / SUBMITTED bit.
PROP_MESSAGE_FLAGS = "Integer 0x0E07"
MSG_FLAGS_VALUE = "5"  # MSGFLAG_READ | MSGFLAG_SUBMIT

# PR_SUBMIT_FLAGS, PT_LONG.  Should be 0 for a delivered/sent message
# (any nonzero value indicates "submission in progress" workflow).
PROP_SUBMIT_FLAGS = "Integer 0x0E14"
SUBMIT_FLAGS_VALUE = "0"

EXT_PROP_PAYLOAD = {
    "singleValueExtendedProperties": [
        {"id": PROP_MESSAGE_FLAGS, "value": MSG_FLAGS_VALUE},
        {"id": PROP_SUBMIT_FLAGS, "value": SUBMIT_FLAGS_VALUE},
    ]
}

# Direct write of the high-level isDraft property is undocumented (the
# Graph reference page lists isDraft without a 'writeable' annotation),
# but in practice some tenants accept it. We try this first because if
# it works it's the cleanest possible PATCH; on rejection we fall back
# to the MAPI-flag bundle above.
DIRECT_PAYLOAD = {"isDraft": False}


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


def iter_draft_message_ids(graph: GraphClient, mailbox: str, folder_id: str):
    """Page through every draft in the folder, yielding message id strings.

    Uses Graph's $filter=isDraft eq true to avoid pulling non-draft messages.
    """
    path = (
        f"/users/{quote(mailbox)}/mailFolders/{folder_id}/messages"
        f"?$top={PAGE_SIZE}&$select=id&$filter=isDraft eq true"
    )
    while path:
        resp = graph.get(path, expect_status=(200,))
        body = resp.json()
        for m in body.get("value", []):
            yield m["id"]
        next_link = body.get("@odata.nextLink")
        path = _strip_base(next_link) if next_link else None


def patch_one(graph: GraphClient, mailbox: str, msg_id: str) -> str:
    """Try PATCH strategies in order; return the strategy name that worked.

    Strategies:
      'direct'  : PATCH {"isDraft": false}.  Undocumented; some tenants accept
                  it. Cleanest fix when it works.
      'mapi'    : PATCH PR_MESSAGE_FLAGS=5 + PR_SUBMIT_FLAGS=0 via
                  singleValueExtendedProperties. Reliable mechanism but
                  Graph's isDraft derivation has been observed to lag this
                  on some tenants.

    On success returns one of the strategy names. On total failure raises
    the last GraphError encountered.
    """
    path = f"/users/{quote(mailbox)}/messages/{msg_id}"
    last_err: GraphError | None = None
    try:
        graph.patch(path, json=DIRECT_PAYLOAD, expect_status=(200,))
        return "direct"
    except GraphError as e:
        last_err = e
    try:
        graph.patch(path, json=EXT_PROP_PAYLOAD, expect_status=(200,))
        return "mapi"
    except GraphError as e:
        last_err = e
    raise last_err  # type: ignore[misc]


def fix_mailbox(graph: GraphClient, mailbox: str, *, dry_run: bool) -> dict:
    log = logger.bind(ctx=f"fix-drafts[{mailbox}]")
    stats = {
        "mailbox": mailbox,
        "folders_walked": 0,
        "drafts_found": 0,
        "fixed": 0,
        "failed": 0,
        "via_direct": 0,
        "via_mapi": 0,
    }

    def walk(folder: dict, prefix: str) -> None:
        path_label = f"{prefix}/{folder['displayName']}" if prefix else folder["displayName"]
        stats["folders_walked"] += 1
        # Collect ids first so we don't paginate while mutating.
        ids = list(iter_draft_message_ids(graph, mailbox, folder["id"]))
        if ids:
            stats["drafts_found"] += len(ids)
            log.info("  {!r}: {} draft(s)", path_label, len(ids))
            if not dry_run:
                for mid in ids:
                    try:
                        via = patch_one(graph, mailbox, mid)
                        stats["fixed"] += 1
                        stats[f"via_{via}"] = stats.get(f"via_{via}", 0) + 1
                    except GraphError as e:
                        stats["failed"] += 1
                        log.warning("    couldn't patch {} ({})", mid, e.status)
                    if stats["fixed"] and stats["fixed"] % 500 == 0:
                        log.info("    ... fixed {} so far", stats["fixed"])
        if int(folder.get("childFolderCount") or 0) > 0:
            for sub in list_child_folders(graph, mailbox, folder["id"]):
                walk(sub, path_label)

    for top in list_child_folders(graph, mailbox, "msgFolderRoot"):
        if top.get("displayName") in SKIP_TOP_LEVEL:
            continue
        walk(top, "")

    if stats["drafts_found"] == 0:
        log.info("No imported-as-draft messages found; mailbox is clean.")
    elif dry_run:
        log.info("DRY RUN -- would PATCH {} message(s) across {} folder(s).",
                 stats["drafts_found"], stats["folders_walked"])
    else:
        log.info(
            "Fixed {} of {} draft(s); {} failed.  (via direct: {}, via mapi: {})",
            stats["fixed"], stats["drafts_found"], stats["failed"],
            stats["via_direct"], stats["via_mapi"],
        )
    return stats


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Clear isDraft on imported messages by PATCHing PR_MESSAGE_FLAGS=1.",
    )
    ap.add_argument("-c", "--config", required=True)
    ap.add_argument("-m", "--mapping", help="mapping.csv path")
    ap.add_argument("--mailbox", action="append", help="UPN of a single mailbox (repeatable)")
    ap.add_argument("--dry-run", action="store_true", help="Walk + count, no PATCH calls.")
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
        logger.bind(ctx="fix-drafts").warning("DRY RUN -- no PATCHes will be issued.")

    pool = AppPool(cfg.apps)
    parallelism = min(cfg.migration.max_parallel_mailboxes, len(mailboxes)) or 1

    logger.bind(ctx="fix-drafts").info(
        "Processing {} mailbox(es) with {} parallel workers, {} app(s).",
        len(mailboxes), parallelism, len(cfg.apps),
    )

    results: list[dict] = []
    with GraphClient(pool, cfg.throttle) as graph:
        with ThreadPoolExecutor(max_workers=parallelism, thread_name_prefix="fix-drafts") as ex:
            futs = {
                ex.submit(fix_mailbox, graph, m, dry_run=ns.dry_run): m
                for m in mailboxes
            }
            for fut in as_completed(futs):
                m = futs[fut]
                try:
                    results.append(fut.result())
                except GraphError as e:
                    logger.bind(ctx=f"fix-drafts[{m}]").error("Graph error: {}", e)
                    results.append({"mailbox": m, "error": str(e)})
                except Exception as e:
                    logger.bind(ctx=f"fix-drafts[{m}]").exception("Crashed: {}", e)
                    results.append({"mailbox": m, "error": str(e)})

    print()
    hdr = f"{'mailbox':<46} {'folders':>8} {'drafts':>8} {'fixed':>8}  status"
    print(hdr)
    print("-" * len(hdr))
    for r in sorted(results, key=lambda x: x["mailbox"]):
        if "error" in r:
            print(f"{r['mailbox']:<46} {'-':>8} {'-':>8} {'-':>8}  ERROR: {r['error'][:30]}")
        else:
            note = "ok" if not r["failed"] else f"{r['failed']} failed"
            print(
                f"{r['mailbox']:<46} {r['folders_walked']:>8} "
                f"{r['drafts_found']:>8} {r['fixed']:>8}  {note}"
            )
    return 0


if __name__ == "__main__":
    sys.exit(main())
