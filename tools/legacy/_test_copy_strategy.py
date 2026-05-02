r"""Quick experiment: does Graph's /copy create a non-draft message?

Last resort Graph idea before going to EWS.  Picks one draft message in
the mailbox (in a non-Drafts folder), copies it to the same folder via
``POST /users/{user}/messages/{id}/copy``, then reports what isDraft /
PR_MESSAGE_FLAGS look like on the copy.  Always deletes the copy at the
end so we don't leave a duplicate regardless of outcome.

Three things to check on the copy:

  1. isDraft from the copy response payload (Graph echoes it back)
  2. isDraft re-queried via GET on the new id (in case the response is
     stale)
  3. PR_MESSAGE_FLAGS via $expand on the new id (definitive proof of
     whether MSGFLAG_UNSENT 0x08 is set or cleared)

If isDraft=false / msgFlg 0x08 clear on the copy, copy+delete becomes
the Graph-only fix and we wire it into _fix_drafts.py.  If isDraft=true
/ msgFlg has 0x08 set, we go EWS.

Usage:

  .\.venv\Scripts\python.exe _test_copy_strategy.py -c config.toml \
      --mailbox UPN [--folder sentitems]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from urllib.parse import quote

from loguru import logger

from jtet_pstmigrate.auth import AppPool
from jtet_pstmigrate.config import AppConfig, expand_user_paths
from jtet_pstmigrate.graph_client import GraphClient, GraphError


def _find_one_draft(graph: GraphClient, mailbox: str, folder_id: str | None) -> dict | None:
    """Return the first draft message we can find, with id + parentFolderId.

    If ``folder_id`` is provided, search only that folder.  Otherwise
    walk the top-level folders until we find a draft.
    """
    select = "id,subject,parentFolderId,isDraft"
    if folder_id:
        path = f"/users/{quote(mailbox)}/mailFolders/{folder_id}/messages"
        params = {
            "$top": "1",
            "$select": select,
            "$filter": "isDraft eq true",
        }
        resp = graph.get(path, params=params, expect_status=(200,))
        items = resp.json().get("value", [])
        return items[0] if items else None

    # No folder pinned: search the whole mailbox for any draft outside the
    # real Drafts folder.  (We don't want to test on a legitimate draft.)
    path = f"/users/{quote(mailbox)}/messages"
    params = {
        "$top": "25",
        "$select": select,
        "$filter": "isDraft eq true",
    }
    resp = graph.get(path, params=params, expect_status=(200,))
    drafts = resp.json().get("value", [])
    if not drafts:
        return None

    # Look up the displayName of each parent folder so we can skip Drafts.
    seen_folders: dict[str, str] = {}
    for m in drafts:
        pf = m["parentFolderId"]
        if pf not in seen_folders:
            fr = graph.get(
                f"/users/{quote(mailbox)}/mailFolders/{pf}",
                params={"$select": "displayName"},
                expect_status=(200,),
            )
            seen_folders[pf] = fr.json().get("displayName", "")
        if seen_folders[pf] != "Drafts":
            return m
    return None


def _get_msg_flags(graph: GraphClient, mailbox: str, msg_id: str) -> tuple[bool, int | None]:
    """Re-fetch a message and return (isDraft, PR_MESSAGE_FLAGS-as-int)."""
    path = f"/users/{quote(mailbox)}/messages/{msg_id}"
    params = {
        "$select": "id,isDraft,parentFolderId",
        "$expand": "singleValueExtendedProperties($filter=id eq 'Integer 0x0E07')",
    }
    resp = graph.get(path, params=params, expect_status=(200,))
    body = resp.json()
    is_draft = bool(body.get("isDraft"))
    flags: int | None = None
    for ep in body.get("singleValueExtendedProperties") or []:
        ep_id = (ep.get("id") or "").lower().replace(" 0x0", " 0x")
        if ep_id.startswith("integer 0xe07") or ep_id.startswith("integer 0x0e07"):
            try:
                flags = int(ep.get("value") or "")
            except ValueError:
                pass
    return is_draft, flags


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("-c", "--config", required=True)
    ap.add_argument("--mailbox", required=True, help="UPN of the mailbox to test on")
    ap.add_argument(
        "--folder",
        help="Optional well-known folder id (e.g. 'sentitems', 'inbox') to "
             "pick the test message from. Default: first draft we find "
             "outside Drafts.",
    )
    ns = ap.parse_args()

    cfg = expand_user_paths(AppConfig.load(Path(ns.config)))
    logger.remove()
    logger.add(sys.stderr, level="INFO",
               format="{time:HH:mm:ss} | {level: <7} | {message}")

    pool = AppPool(cfg.apps)
    with GraphClient(pool, cfg.throttle) as graph:
        logger.info("Looking for a draft message to test on...")
        target = _find_one_draft(graph, ns.mailbox, ns.folder)
        if not target:
            logger.error("No draft messages found to test with. "
                         "(Mailbox may already be clean, or --folder was empty.)")
            return 1

        original_id = target["id"]
        parent_id = target["parentFolderId"]
        subject = target.get("subject") or "(no subject)"
        logger.info("Test target: {!r}", subject[:60])
        logger.info("  original id        = {}...", original_id[:40])
        logger.info("  parent folder id   = {}...", parent_id[:40])
        logger.info("  isDraft (original) = {}", target.get("isDraft"))

        # Step 1: copy to same folder
        logger.info("POST /messages/{{id}}/copy with destinationId = same folder")
        try:
            cresp = graph.post(
                f"/users/{quote(ns.mailbox)}/messages/{original_id}/copy",
                json={"destinationId": parent_id},
                expect_status=(200, 201),
            )
        except GraphError as e:
            logger.error("Copy failed: {}", e)
            return 2
        copy_body = cresp.json()
        copy_id = copy_body.get("id")
        if not copy_id:
            logger.error("Copy response had no id; can't continue.")
            return 2

        copy_response_isdraft = copy_body.get("isDraft")
        logger.info("  copy id (in response) = {}...", copy_id[:40])
        logger.info("  isDraft (response body) = {}", copy_response_isdraft)

        # Step 2: re-query to see what Graph really stores after the copy
        try:
            recheck_isdraft, recheck_flags = _get_msg_flags(graph, ns.mailbox, copy_id)
        except GraphError as e:
            logger.error("Could not re-fetch the copy: {}", e)
            recheck_isdraft, recheck_flags = copy_response_isdraft, None

        logger.info("  isDraft (re-fetched)    = {}", recheck_isdraft)
        flags_str = f"0x{recheck_flags:04x}" if recheck_flags is not None else "(unavailable)"
        logger.info("  PR_MESSAGE_FLAGS        = {}", flags_str)
        if recheck_flags is not None:
            unsent = bool(recheck_flags & 0x08)
            logger.info("  -> MSGFLAG_UNSENT (0x08) {}",
                        "SET (still a draft)" if unsent else "CLEAR (NOT a draft!)")

        # Step 3: clean up the copy. Always delete -- we just wanted to know
        # what state it was created in, not actually duplicate the message.
        logger.info("DELETE the test copy to clean up...")
        try:
            graph.delete(
                f"/users/{quote(ns.mailbox)}/messages/{copy_id}",
                expect_status=(204,),
            )
            logger.info("  copy deleted OK")
        except GraphError as e:
            logger.warning(
                "Could not delete the test copy ({}); manually remove "
                "message id {} from the mailbox.", e, copy_id,
            )

        print()
        if recheck_flags is not None and not (recheck_flags & 0x08) and not recheck_isdraft:
            print("RESULT: /copy creates a non-draft message.  Graph fix is viable;")
            print("        we can wire copy+delete-original into _fix_drafts.py.")
            return 0
        if recheck_isdraft:
            print("RESULT: /copy preserves isDraft=true.  Graph cannot fix this.")
            print("        EWS is the next step.")
            return 3
        print("RESULT: ambiguous (isDraft={}, flags={}).  See log above.".format(
            recheck_isdraft, flags_str))
        return 4


if __name__ == "__main__":
    sys.exit(main())
