r"""Quick experiment: does Graph's /copy operation create a non-draft message?

Background
----------
We've established that PATCH-based fixes for the imported-as-draft bug do
not work:

  * PATCH {"isDraft": false}                  -> 200 OK, silently no-op'd
  * PATCH PR_MESSAGE_FLAGS via extended props -> 200 OK, UNSENT bit
                                                 (0x08) preserved.

Per [MS-OXCMSG] mfUnsent is read-only after a message has been saved as
a draft.  That leaves one untested Graph operation: ``POST /messages/{id}/copy``.
``/copy`` creates an entirely new message resource with a fresh server-side
identity.  If Exchange treats that creation as "import" rather than as
"clone the source's draft state", the copy will have UNSENT cleared.
The Graph docs are silent on this so we test empirically.

What this script does
---------------------
  1. Find ONE draft message in --folder of --mailbox.
  2. POST /messages/{id}/copy with destinationId = its own parentFolderId.
  3. GET the copy with $expand on PR_MESSAGE_FLAGS / PR_SUBMIT_FLAGS so
     we can see exactly what landed.
  4. Print the original vs. the copy side-by-side and a verdict line.
  5. Leave both messages in place -- nothing is deleted.  You decide
     whether to clean up the duplicate copy after seeing the verdict.

Outcomes
--------
  * Copy has draft=no, UNSENT clear  -> Graph-only fix is feasible; we'll
    add a /copy + /delete strategy to _fix_drafts.py.
  * Copy still has draft=yes / UNSENT set -> /copy preserves draft state;
    EWS is the only remaining mechanism.

Usage
-----

  .\.venv\Scripts\python.exe _test_draft_copy.py -c config.toml \
      --mailbox UPN --folder sentitems
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from urllib.parse import quote

from loguru import logger

from jtet_pstmigrate.auth import AppPool
from jtet_pstmigrate.config import AppConfig, expand_user_paths
from jtet_pstmigrate.graph_client import GraphClient

WELL_KNOWN = {"inbox", "drafts", "sentitems", "deleteditems", "junkemail",
              "outbox", "archive", "scheduled", "msgfolderroot"}


def resolve_folder_id(graph: GraphClient, mailbox: str, folder: str) -> str:
    """Accept either a well-known name (sentitems / inbox / ...) or a
    display name and return Graph's folder id."""
    if folder.lower() in WELL_KNOWN:
        return folder.lower()
    path = (
        f"/users/{quote(mailbox)}/mailFolders"
        f"?$top=100&$select=id,displayName"
    )
    while path:
        resp = graph.get(path, expect_status=(200,))
        body = resp.json()
        for f in body.get("value", []):
            if f["displayName"].lower() == folder.lower():
                return f["id"]
        nl = body.get("@odata.nextLink")
        path = nl[len("https://graph.microsoft.com/v1.0"):] if nl else None
    raise SystemExit(f"Folder {folder!r} not found in {mailbox}")


def find_one_draft(graph: GraphClient, mailbox: str, folder_id: str) -> dict:
    path = f"/users/{quote(mailbox)}/mailFolders/{folder_id}/messages"
    params = {
        "$top": "1",
        "$select": "id,subject,isDraft,parentFolderId",
        "$filter": "isDraft eq true",
    }
    resp = graph.get(path, params=params, expect_status=(200,))
    items = resp.json().get("value", [])
    if not items:
        raise SystemExit(f"No drafts found in folder {folder_id}; pick another --folder")
    return items[0]


def fetch_with_flags(graph: GraphClient, mailbox: str, msg_id: str) -> dict:
    """Re-fetch a message with PR_MESSAGE_FLAGS / PR_SUBMIT_FLAGS expanded."""
    path = f"/users/{quote(mailbox)}/messages/{msg_id}"
    params = {
        "$select": "id,subject,isDraft,isRead,parentFolderId,lastModifiedDateTime",
        "$expand": (
            "singleValueExtendedProperties("
            "$filter=id eq 'Integer 0x0E07' or id eq 'Integer 0x0E14')"
        ),
    }
    resp = graph.get(path, params=params, expect_status=(200,))
    return resp.json()


def ext_int(msg: dict, prop_id: str) -> int | None:
    want = prop_id.lower().replace(" 0x0", " 0x")
    for ep in msg.get("singleValueExtendedProperties") or []:
        if (ep.get("id") or "").lower().replace(" 0x0", " 0x") == want:
            try:
                return int(ep["value"])
            except (KeyError, TypeError, ValueError):
                return None
    return None


def render(label: str, msg: dict) -> None:
    msg_flags = ext_int(msg, "Integer 0x0E07")
    sub_flags = ext_int(msg, "Integer 0x0E14")
    print(f"  {label}:")
    print(f"    id           = {msg.get('id')[:40]}{'...' if len(msg.get('id') or '') > 40 else ''}")
    print(f"    subject      = {(msg.get('subject') or '(none)')[:60]}")
    print(f"    isDraft      = {msg.get('isDraft')}")
    print(f"    isRead       = {msg.get('isRead')}")
    print(f"    msgFlg       = {('0x%04x' % msg_flags) if msg_flags is not None else '(missing)'}")
    print(f"    subFlg       = {('0x%04x' % sub_flags) if sub_flags is not None else '(missing)'}")
    print(f"    lastModified = {msg.get('lastModifiedDateTime')}")
    print(f"    parentFolder = {(msg.get('parentFolderId') or '')[:40]}...")


def main() -> int:
    ap = argparse.ArgumentParser(description="Test whether Graph /copy resets isDraft on imported messages.")
    ap.add_argument("-c", "--config", required=True)
    ap.add_argument("--mailbox", required=True, help="Target mailbox UPN")
    ap.add_argument("--folder", default="sentitems",
                    help="Folder to find a draft in (well-known name or display name)")
    ns = ap.parse_args()

    cfg = expand_user_paths(AppConfig.load(Path(ns.config)))

    logger.remove()
    logger.add(sys.stderr, level="INFO", format="{time:HH:mm:ss} | {level: <7} | {message}")

    pool = AppPool(cfg.apps)
    with GraphClient(pool, cfg.throttle) as graph:
        logger.info("Resolving folder {} in {}", ns.folder, ns.mailbox)
        folder_id = resolve_folder_id(graph, ns.mailbox, ns.folder)

        logger.info("Searching for one draft in folder...")
        draft = find_one_draft(graph, ns.mailbox, folder_id)
        logger.info("Found draft: {!r}", (draft.get("subject") or "")[:60])

        logger.info("Fetching original with extended properties...")
        original = fetch_with_flags(graph, ns.mailbox, draft["id"])

        parent_id = original.get("parentFolderId") or folder_id
        logger.info("POST /messages/{}/copy -> destinationId={}", draft["id"][:20] + "...", parent_id[:20] + "...")
        copy_path = f"/users/{quote(ns.mailbox)}/messages/{draft['id']}/copy"
        copy_resp = graph.post(copy_path, json={"destinationId": parent_id}, expect_status=(200, 201, 202))
        copy_body = copy_resp.json() if copy_resp.content else {}
        copy_id = copy_body.get("id")
        if not copy_id:
            raise SystemExit("/copy did not return a new message id")
        logger.info("Copy created with id {}", copy_id[:40] + "...")

        logger.info("Fetching copy with extended properties...")
        copy = fetch_with_flags(graph, ns.mailbox, copy_id)

    print()
    print("=" * 78)
    print("/copy experiment results")
    print("=" * 78)
    render("ORIGINAL (still in place)", original)
    print()
    render("COPY     (newly created)", copy)
    print()
    print("=" * 78)
    o_unsent = (ext_int(original, "Integer 0x0E07") or 0) & 0x08
    c_unsent = (ext_int(copy, "Integer 0x0E07") or 0) & 0x08
    if copy.get("isDraft") is False and not c_unsent:
        print("VERDICT: /copy creates a non-draft message.  Graph-only fix is viable.")
        print("         _fix_drafts.py can be updated to use /copy + /delete original.")
    elif copy.get("isDraft") is True and c_unsent:
        print("VERDICT: /copy preserves draft state.  Graph cannot fix this.")
        print("         EWS-based remediation is the only remaining option.")
    else:
        print("VERDICT: Mixed result -- isDraft and msgFlg disagree on the copy.")
        print("         Share the output above and we'll adjust strategy.")
    print()
    print(f"Cleanup: the copy is currently sitting in the same folder as the original.")
    print(f"         Delete it manually in OWA, or run:")
    print(f"           DELETE /users/{ns.mailbox}/messages/{copy_id}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
