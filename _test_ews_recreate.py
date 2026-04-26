r"""Test whether EWS CreateItem from MIME can land a non-draft message.

Background
----------
We've empirically confirmed that neither Graph nor EWS can clear
MSGFLAG_UNSENT (0x08) on already-saved messages -- Exchange Online's
store rejects every property-update path silently while reporting
'Success'.

The remaining theory: the store decides UNSENT at CREATE time, not
at UPDATE time.  CreateItem with MimeContent + SavedItemFolderId
pointing at a NON-Drafts folder might create a fresh item with UNSENT
already clear.  This is the path most third-party PST migration tools
(ShareGate, BitTitan) appear to use.

This script proves or disproves the theory on ONE message:

  1. Find a stuck draft in the source folder (default: Sent Items)
  2. GetItem with IncludeMimeContent=true to get the original MIME
  3. CreateItem with that MIME + MessageDisposition=SaveOnly +
       SavedItemFolderId=DistinguishedFolderId('sentitems')
  4. GetItem the newly-created item and check PR_MESSAGE_FLAGS
  5. Delete the newly-created item (we don't want a duplicate)

If step 4 reports UNSENT clear, we'll wire CreateItem-from-MIME +
delete-original into a real remediation script.

If step 4 reports UNSENT still set, the store is making the decision
based on something other than folder location and the only remaining
fix is Microsoft 365 Network Upload (PST import service).

Usage
-----
  .\.venv\Scripts\python.exe _test_ews_recreate.py -c config.toml \
      --mailbox UPN \
      [--folder sentitems|inbox]   # default: sentitems
"""
from __future__ import annotations

import argparse
import sys
from html import escape
from pathlib import Path
from xml.etree import ElementTree as ET

import httpx
from loguru import logger

from _fix_drafts_ews import (
    EwsClient,
    EwsError,
    NS_M,
    NS_T,
    PROP_TAG_MESSAGE_FLAGS,
    get_ews_token,
)
from jtet_pstmigrate.config import AppConfig, expand_user_paths


def _get_mime_and_subject(
    ews: EwsClient, item_id: str, change_key: str,
) -> tuple[str, str, int | None]:
    """GetItem returning (subject, base64 MIME, current PR_MESSAGE_FLAGS)."""
    body = (
        '<m:GetItem>'
        '<m:ItemShape>'
        '<t:BaseShape>IdOnly</t:BaseShape>'
        '<t:IncludeMimeContent>true</t:IncludeMimeContent>'
        '<t:AdditionalProperties>'
        '<t:FieldURI FieldURI="item:Subject"/>'
        f'<t:ExtendedFieldURI PropertyTag="{PROP_TAG_MESSAGE_FLAGS}" PropertyType="Integer"/>'
        '</t:AdditionalProperties>'
        '</m:ItemShape>'
        '<m:ItemIds>'
        f'<t:ItemId Id="{escape(item_id)}" ChangeKey="{escape(change_key)}"/>'
        '</m:ItemIds>'
        '</m:GetItem>'
    )
    resp = ews.call(body)
    rm = resp.find(f"{NS_M}GetItemResponse/{NS_M}ResponseMessages/{NS_M}GetItemResponseMessage")
    if rm is None or rm.attrib.get("ResponseClass") != "Success":
        raise EwsError(f"GetItem failed: {ET.tostring(rm or resp, encoding='unicode')[:300]}")
    msg = rm.find(f"{NS_M}Items/{NS_T}Message")
    if msg is None:
        raise EwsError("GetItem returned no Message")
    subject = msg.findtext(f"{NS_T}Subject") or "(no subject)"
    mime_el = msg.find(f"{NS_T}MimeContent")
    if mime_el is None or mime_el.text is None:
        raise EwsError("GetItem returned no MimeContent")
    mime_b64 = mime_el.text.strip()
    flags: int | None = None
    for ep in msg.findall(f"{NS_T}ExtendedProperty"):
        uri = ep.find(f"{NS_T}ExtendedFieldURI")
        val = ep.find(f"{NS_T}Value")
        if uri is None or val is None or val.text is None:
            continue
        tag = uri.attrib.get("PropertyTag", "").lower()
        if tag.endswith("0e07") or tag.endswith("e07"):
            try:
                flags = int(val.text)
            except ValueError:
                pass
    return subject, mime_b64, flags


def _create_from_mime(
    ews: EwsClient, mime_b64: str, dest_folder: str,
) -> tuple[str, str]:
    """CreateItem with MimeContent + SavedItemFolderId=dest_folder.

    Returns (item_id, change_key) of the newly-created item.
    Note: We do NOT pass any extended properties or modified MIME -- this
    is testing whether a vanilla MIME re-import lands non-draft when
    targeted at a non-Drafts folder.  If we have to mutate flags at
    create time later, that's a follow-up.
    """
    body = (
        '<m:CreateItem MessageDisposition="SaveOnly">'
        '<m:SavedItemFolderId>'
        f'<t:DistinguishedFolderId Id="{dest_folder}"/>'
        '</m:SavedItemFolderId>'
        '<m:Items>'
        '<t:Message>'
        f'<t:MimeContent CharacterSet="UTF-8">{mime_b64}</t:MimeContent>'
        '</t:Message>'
        '</m:Items>'
        '</m:CreateItem>'
    )
    resp = ews.call(body)
    rm = resp.find(f"{NS_M}CreateItemResponse/{NS_M}ResponseMessages/{NS_M}CreateItemResponseMessage")
    if rm is None or rm.attrib.get("ResponseClass") != "Success":
        code = rm.findtext(f"{NS_M}ResponseCode") if rm is not None else None
        text = rm.findtext(f"{NS_M}MessageText") if rm is not None else None
        raise EwsError(
            f"CreateItem failed: ResponseClass={rm.attrib.get('ResponseClass') if rm is not None else None} "
            f"code={code} message={text}"
        )
    msg = rm.find(f"{NS_M}Items/{NS_T}Message")
    if msg is None:
        raise EwsError("CreateItem returned no Message")
    iid_el = msg.find(f"{NS_T}ItemId")
    if iid_el is None:
        raise EwsError("CreateItem returned Message without ItemId")
    return iid_el.attrib.get("Id", ""), iid_el.attrib.get("ChangeKey", "")


def _get_flags(ews: EwsClient, item_id: str, change_key: str) -> int | None:
    body = (
        '<m:GetItem>'
        '<m:ItemShape>'
        '<t:BaseShape>IdOnly</t:BaseShape>'
        '<t:AdditionalProperties>'
        f'<t:ExtendedFieldURI PropertyTag="{PROP_TAG_MESSAGE_FLAGS}" PropertyType="Integer"/>'
        '</t:AdditionalProperties>'
        '</m:ItemShape>'
        '<m:ItemIds>'
        f'<t:ItemId Id="{escape(item_id)}" ChangeKey="{escape(change_key)}"/>'
        '</m:ItemIds>'
        '</m:GetItem>'
    )
    resp = ews.call(body)
    rm = resp.find(f"{NS_M}GetItemResponse/{NS_M}ResponseMessages/{NS_M}GetItemResponseMessage")
    if rm is None or rm.attrib.get("ResponseClass") != "Success":
        return None
    msg = rm.find(f"{NS_M}Items/{NS_T}Message")
    if msg is None:
        return None
    for ep in msg.findall(f"{NS_T}ExtendedProperty"):
        uri = ep.find(f"{NS_T}ExtendedFieldURI")
        val = ep.find(f"{NS_T}Value")
        if uri is None or val is None or val.text is None:
            continue
        tag = uri.attrib.get("PropertyTag", "").lower()
        if tag.endswith("0e07") or tag.endswith("e07"):
            try:
                return int(val.text)
            except ValueError:
                pass
    return None


def _delete_item(ews: EwsClient, item_id: str, change_key: str) -> None:
    body = (
        '<m:DeleteItem DeleteType="HardDelete">'
        '<m:ItemIds>'
        f'<t:ItemId Id="{escape(item_id)}" ChangeKey="{escape(change_key)}"/>'
        '</m:ItemIds>'
        '</m:DeleteItem>'
    )
    resp = ews.call(body)
    rm = resp.find(f"{NS_M}DeleteItemResponse/{NS_M}ResponseMessages/{NS_M}DeleteItemResponseMessage")
    if rm is None or rm.attrib.get("ResponseClass") != "Success":
        raise EwsError("DeleteItem failed")


def _find_one_draft_in_folder(
    ews: EwsClient, folder_distinguished_id: str,
) -> tuple[str, str] | None:
    """Find one draft message in the given distinguished folder."""
    folders = ews.find_all_folders()
    target_name_map = {
        "sentitems": ("Sent Items", "Sent"),
        "inbox": ("Inbox",),
    }
    candidates = target_name_map.get(folder_distinguished_id, (folder_distinguished_id,))
    for f in folders:
        if f["display_name"] in candidates:
            drafts = ews.find_drafts_in_folder(f["id"], f["change_key"])
            if drafts:
                return drafts[0]
    return None


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("-c", "--config", required=True)
    ap.add_argument("--mailbox", required=True)
    ap.add_argument(
        "--folder", default="sentitems",
        choices=["sentitems", "inbox"],
        help="Distinguished folder to look for a stuck draft in (default: sentitems).",
    )
    ap.add_argument("--app", help="App display name to use (default: first).")
    ns = ap.parse_args()

    cfg = expand_user_paths(AppConfig.load(Path(ns.config)))
    logger.remove()
    logger.add(sys.stderr, level="INFO",
               format="{time:HH:mm:ss} | {level: <7} | {message}")

    if ns.app:
        app = next((a for a in cfg.apps if a.display_name == ns.app), None)
        if app is None:
            logger.error("App {!r} not found in config.", ns.app)
            return 2
    else:
        app = cfg.apps[0]

    token = get_ews_token(app)
    timeout = httpx.Timeout(120.0, connect=30.0)
    with httpx.Client(timeout=timeout, http2=False) as http:
        ews = EwsClient(token, ns.mailbox, http)

        logger.info("Looking for a draft in {!r} of {}...", ns.folder, ns.mailbox)
        target = _find_one_draft_in_folder(ews, ns.folder)
        if not target:
            logger.error("No drafts found in {!r}.", ns.folder)
            return 1
        item_id, change_key = target

        logger.info("Found target -- fetching MIME...")
        try:
            subject, mime_b64, orig_flags = _get_mime_and_subject(ews, item_id, change_key)
        except EwsError as e:
            logger.error("GetItem with MIME failed: {}", e)
            return 3

        logger.info("Test target subject: {!r}", subject[:60])
        flags_str = f"0x{orig_flags:04x}" if orig_flags is not None else "(unavailable)"
        logger.info("Original PR_MESSAGE_FLAGS = {} (UNSENT bit {})",
                    flags_str,
                    "set" if (orig_flags is not None and (orig_flags & 0x08)) else "clear")
        logger.info("MIME size = {} bytes (base64-decoded: ~{} bytes)",
                    len(mime_b64), int(len(mime_b64) * 0.75))

        logger.info("Creating new item via EWS CreateItem -> {!r}...", ns.folder)
        try:
            new_id, new_ck = _create_from_mime(ews, mime_b64, ns.folder)
        except EwsError as e:
            logger.error("CreateItem failed: {}", e)
            return 4

        logger.info("New item id = {}...", new_id[:40])
        logger.info("Re-fetching to inspect flags on the new item...")
        new_flags = _get_flags(ews, new_id, new_ck)
        new_flags_str = f"0x{new_flags:04x}" if new_flags is not None else "(unavailable)"
        unsent_set = new_flags is not None and (new_flags & 0x08)
        logger.info("NEW item PR_MESSAGE_FLAGS = {} (UNSENT bit {})",
                    new_flags_str,
                    "set" if unsent_set else "CLEAR")

        logger.info("Cleaning up: deleting the test copy...")
        try:
            _delete_item(ews, new_id, new_ck)
            logger.info("  test copy hard-deleted OK")
        except EwsError as e:
            logger.warning("  could not delete test copy: {} -- you may need to clean up manually", e)

        print()
        if new_flags is None:
            print("RESULT: ambiguous -- could not read flags on the new item.")
            return 5
        if not unsent_set:
            print("RESULT: SUCCESS -- EWS CreateItem from MIME landed non-draft.")
            print("        We can wire 'export MIME -> CreateItem -> delete original'")
            print("        into a remediation script. Reply 'go' to write it.")
            return 0
        print("RESULT: FAILURE -- new item is also a draft (UNSENT bit set).")
        print("        Exchange Online stores all CreateItem-via-MIME items as")
        print("        drafts regardless of SavedItemFolderId.")
        print("        Only path forward is Microsoft 365 Network Upload (PST")
        print("        import service), which uses internal MRS APIs not exposed")
        print("        through Graph or EWS.")
        return 6


if __name__ == "__main__":
    sys.exit(main())
