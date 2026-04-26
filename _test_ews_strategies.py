r"""Try several EWS write strategies on ONE draft message and verify
via GetItem which (if any) actually changes PR_MESSAGE_FLAGS.

Why this exists
---------------
_fix_drafts_ews.py reported "Cleared 485 of 485 draft(s)" but a follow-up
GET shows PR_MESSAGE_FLAGS still 0x0419 (UNSENT bit set).  Confirmed that
EWS, like Graph, accepts the write call (lastModifiedDateTime bumps) but
silently drops the field change.

Microsoft's MAPI documentation says MSGFLAG_UNSENT is "store-controlled
and may be reset upon save" -- the API accepts your write but the store
does what it wants.  This script tries the alternate write approaches
that some PST-migration tools claim work, and reports objectively which
one (if any) actually changes the stored value:

  A. SetItemField PR_MESSAGE_FLAGS = 1   (what we already tried; no-op)
  B. SetItemField PR_MESSAGE_FLAGS = 5   (READ | SUBMITTED)
  C. DeleteItemField PR_MESSAGE_FLAGS    (remove the property entirely)
  D. SetItemField message:IsRead=true    (no-op writable field; force save cycle)
  E. SetItemField PR_MESSAGE_FLAGS = 5
       AND PR_SUBMIT_FLAGS = 0           (set SUBMITTED, clear submit-pending)
  F. CreateItem from MimeContent with
       PR_MESSAGE_FLAGS=5 set at create
       time + SavedItemFolderId=sentitems
       (re-create from MIME; might land
       non-draft if Exchange respects
       create-time flags)

Strategy F is the most invasive (creates a new item, would replace the
old one on full rollout) but is the path most migration tools actually
take.  We test it last and clean up the test item.

Usage:

  .\.venv\Scripts\python.exe _test_ews_strategies.py -c config.toml \
      --mailbox UPN
"""
from __future__ import annotations

import argparse
import sys
from html import escape
from pathlib import Path
from xml.etree import ElementTree as ET

import httpx
from loguru import logger

# Reuse helpers from the production EWS fix script.
from _fix_drafts_ews import (
    EWS_ENDPOINT,
    EwsClient,
    EwsError,
    NS_M,
    NS_T,
    PROP_TAG_MESSAGE_FLAGS,
    _envelope,
    get_ews_token,
)
from jtet_pstmigrate.config import AppConfig, expand_user_paths


PROP_TAG_SUBMIT_FLAGS = "0x0E14"


def _get_item_with_flags(ews: EwsClient, item_id: str, change_key: str) -> tuple[int | None, str]:
    """GetItem returning the current PR_MESSAGE_FLAGS + new ChangeKey.

    If the response doesn't include a new ChangeKey we keep the old one
    so the caller can still issue the next UpdateItem.
    """
    body = (
        '<m:GetItem>'
        '<m:ItemShape>'
        '<t:BaseShape>IdOnly</t:BaseShape>'
        '<t:AdditionalProperties>'
        '<t:FieldURI FieldURI="message:IsDraft"/>'
        f'<t:ExtendedFieldURI PropertyTag="{PROP_TAG_MESSAGE_FLAGS}" PropertyType="Integer"/>'
        f'<t:ExtendedFieldURI PropertyTag="{PROP_TAG_SUBMIT_FLAGS}" PropertyType="Integer"/>'
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
    new_id = msg.find(f"{NS_T}ItemId")
    new_ck = new_id.attrib.get("ChangeKey", "") if new_id is not None else ""
    if not new_ck:
        new_ck = change_key
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
    return flags, new_ck


def _strategy_a(ews: EwsClient, item_id: str, change_key: str) -> str:
    body = _update_set_extended(item_id, change_key, PROP_TAG_MESSAGE_FLAGS, "Integer", "1")
    return _run_update(ews, body)


def _strategy_b(ews: EwsClient, item_id: str, change_key: str) -> str:
    body = _update_set_extended(item_id, change_key, PROP_TAG_MESSAGE_FLAGS, "Integer", "5")
    return _run_update(ews, body)


def _strategy_c(ews: EwsClient, item_id: str, change_key: str) -> str:
    body = (
        '<m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite">'
        '<m:ItemChanges>'
        '<t:ItemChange>'
        f'<t:ItemId Id="{escape(item_id)}" ChangeKey="{escape(change_key)}"/>'
        '<t:Updates>'
        '<t:DeleteItemField>'
        f'<t:ExtendedFieldURI PropertyTag="{PROP_TAG_MESSAGE_FLAGS}" PropertyType="Integer"/>'
        '</t:DeleteItemField>'
        '</t:Updates>'
        '</t:ItemChange>'
        '</m:ItemChanges>'
        '</m:UpdateItem>'
    )
    return _run_update(ews, body)


def _strategy_d(ews: EwsClient, item_id: str, change_key: str) -> str:
    body = (
        '<m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite">'
        '<m:ItemChanges>'
        '<t:ItemChange>'
        f'<t:ItemId Id="{escape(item_id)}" ChangeKey="{escape(change_key)}"/>'
        '<t:Updates>'
        '<t:SetItemField>'
        '<t:FieldURI FieldURI="message:IsRead"/>'
        '<t:Message>'
        '<t:IsRead>true</t:IsRead>'
        '</t:Message>'
        '</t:SetItemField>'
        '</t:Updates>'
        '</t:ItemChange>'
        '</m:ItemChanges>'
        '</m:UpdateItem>'
    )
    return _run_update(ews, body)


def _strategy_e(ews: EwsClient, item_id: str, change_key: str) -> str:
    body = (
        '<m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite">'
        '<m:ItemChanges>'
        '<t:ItemChange>'
        f'<t:ItemId Id="{escape(item_id)}" ChangeKey="{escape(change_key)}"/>'
        '<t:Updates>'
        '<t:SetItemField>'
        f'<t:ExtendedFieldURI PropertyTag="{PROP_TAG_MESSAGE_FLAGS}" PropertyType="Integer"/>'
        '<t:Message>'
        '<t:ExtendedProperty>'
        f'<t:ExtendedFieldURI PropertyTag="{PROP_TAG_MESSAGE_FLAGS}" PropertyType="Integer"/>'
        '<t:Value>5</t:Value>'
        '</t:ExtendedProperty>'
        '</t:Message>'
        '</t:SetItemField>'
        '<t:SetItemField>'
        f'<t:ExtendedFieldURI PropertyTag="{PROP_TAG_SUBMIT_FLAGS}" PropertyType="Integer"/>'
        '<t:Message>'
        '<t:ExtendedProperty>'
        f'<t:ExtendedFieldURI PropertyTag="{PROP_TAG_SUBMIT_FLAGS}" PropertyType="Integer"/>'
        '<t:Value>0</t:Value>'
        '</t:ExtendedProperty>'
        '</t:Message>'
        '</t:SetItemField>'
        '</t:Updates>'
        '</t:ItemChange>'
        '</m:ItemChanges>'
        '</m:UpdateItem>'
    )
    return _run_update(ews, body)


def _update_set_extended(
    item_id: str, change_key: str, prop_tag: str, prop_type: str, value: str,
) -> str:
    return (
        '<m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite">'
        '<m:ItemChanges>'
        '<t:ItemChange>'
        f'<t:ItemId Id="{escape(item_id)}" ChangeKey="{escape(change_key)}"/>'
        '<t:Updates>'
        '<t:SetItemField>'
        f'<t:ExtendedFieldURI PropertyTag="{prop_tag}" PropertyType="{prop_type}"/>'
        '<t:Message>'
        '<t:ExtendedProperty>'
        f'<t:ExtendedFieldURI PropertyTag="{prop_tag}" PropertyType="{prop_type}"/>'
        f'<t:Value>{value}</t:Value>'
        '</t:ExtendedProperty>'
        '</t:Message>'
        '</t:SetItemField>'
        '</t:Updates>'
        '</t:ItemChange>'
        '</m:ItemChanges>'
        '</m:UpdateItem>'
    )


def _run_update(ews: EwsClient, body: str) -> str:
    """Run an UpdateItem; return 'success' or 'fail: <reason>'."""
    try:
        resp = ews.call(body)
    except EwsError as e:
        return f"fail: {e}"
    rm = resp.find(f"{NS_M}UpdateItemResponse/{NS_M}ResponseMessages/{NS_M}UpdateItemResponseMessage")
    if rm is None:
        return "fail: no response message"
    cls = rm.attrib.get("ResponseClass", "")
    if cls == "Success":
        return "success"
    code = (rm.findtext(f"{NS_M}ResponseCode") or "").strip()
    msg_text = (rm.findtext(f"{NS_M}MessageText") or "").strip()
    return f"fail ({cls}): {code} -- {msg_text}"


def _find_one_draft(ews: EwsClient) -> tuple[str, str, str] | None:
    """Find one draft message anywhere in the mailbox.  Returns
    (folder_label, item_id, change_key) or None."""
    folders = ews.find_all_folders()
    for f in folders:
        if f["display_name"] in ("Drafts", "Deleted Items"):
            continue
        if f["folder_class"] not in ("IPF.Note", ""):
            continue
        drafts = ews.find_drafts_in_folder(f["id"], f["change_key"])
        if drafts:
            iid, ck = drafts[0]
            return f["display_name"], iid, ck
    return None


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("-c", "--config", required=True)
    ap.add_argument("--mailbox", required=True)
    ap.add_argument("--app", help="Display name of the app entry to use (default: first).")
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
        target = _find_one_draft(ews)
        if not target:
            logger.error("No draft messages found in {} -- nothing to test on.", ns.mailbox)
            return 1
        folder_label, item_id, change_key = target
        logger.info("Test target in folder {!r}", folder_label)
        flags_before, change_key = _get_item_with_flags(ews, item_id, change_key)
        logger.info(
            "  PR_MESSAGE_FLAGS BEFORE = {}",
            f"0x{flags_before:04x}" if flags_before is not None else "(unavailable)"
        )
        if flags_before is None or not (flags_before & 0x08):
            logger.warning(
                "Target message UNSENT bit is already clear -- pick a different "
                "mailbox to test on."
            )
            return 4

        strategies = [
            ("A: SetItemField PR_MESSAGE_FLAGS=1", _strategy_a),
            ("B: SetItemField PR_MESSAGE_FLAGS=5 (READ|SUBMITTED)", _strategy_b),
            ("C: DeleteItemField PR_MESSAGE_FLAGS", _strategy_c),
            ("D: SetItemField message:IsRead=true", _strategy_d),
            ("E: SetItemField PR_MESSAGE_FLAGS=5 + PR_SUBMIT_FLAGS=0", _strategy_e),
        ]

        results: list[tuple[str, str, int | None]] = []
        for label, strategy in strategies:
            logger.info("Trying {}", label)
            outcome = strategy(ews, item_id, change_key)
            try:
                flags_after, change_key = _get_item_with_flags(ews, item_id, change_key)
            except EwsError as e:
                logger.warning("  GetItem after strategy failed: {}", e)
                flags_after = None
            after_str = f"0x{flags_after:04x}" if flags_after is not None else "(unavailable)"
            logger.info("  result: {}    flags now {}", outcome, after_str)
            results.append((label, outcome, flags_after))
            if flags_after is not None and not (flags_after & 0x08):
                logger.success("  *** UNSENT bit cleared by strategy {!r} ***", label)
                break  # stop -- we found one that works

        print()
        print(f"{'strategy':<55} {'outcome':<14} {'flags after':<11}  UNSENT?")
        print("-" * 100)
        for label, outcome, flags_after in results:
            after_str = f"0x{flags_after:04x}" if flags_after is not None else "?"
            unsent = "set" if (flags_after is not None and (flags_after & 0x08)) else (
                "CLEAR" if flags_after is not None else "?"
            )
            print(f"{label:<55} {outcome[:14]:<14} {after_str:<11}  {unsent}")

        winner = next((r for r in results if r[2] is not None and not (r[2] & 0x08)), None)
        print()
        if winner:
            print(f"WINNER: {winner[0]}")
            print("Update _fix_drafts_ews.py to use this strategy.")
            return 0
        print("No EWS strategy successfully cleared MSGFLAG_UNSENT.")
        print("Need to fall back to re-creating items via EWS CreateItem,")
        print("or to Microsoft 365 Network Upload (PST import service).")
        return 5


if __name__ == "__main__":
    sys.exit(main())
