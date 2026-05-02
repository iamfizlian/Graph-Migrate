r"""Read-only diagnostic: dump every relevant date field for the newest
messages in a folder so we can tell whether timestamps were rewritten or
just appear different in Outlook.

Five date sources per message:

  - sentDateTime              : Graph property, set by sender
  - receivedDateTime          : Graph property, set by receiving server
  - createdDateTime           : when the item was created in this mailbox
  - lastModifiedDateTime      : last time anything in the item changed --
                                THIS UPDATES ON MOVE
  - SMTP 'Date:' header       : pulled from internetMessageHeaders;
                                cannot be rewritten by Graph; reflects
                                what the sender actually stamped on the
                                message at send time

Interpreting the output:

  - SMTP Date matches receivedDateTime + both are old           -> normal
  - SMTP Date is old, receivedDateTime is old, lastModified
    is "today"                                                  -> cosmetic
                                                                  (move
                                                                   touched
                                                                   modified
                                                                   stamp;
                                                                   Outlook
                                                                   may be
                                                                   showing
                                                                   that
                                                                   column)
  - SMTP Date is old, receivedDateTime is "today"              -> something
                                                                  rewrote
                                                                  receive
                                                                  time
                                                                  during
                                                                  import
  - SMTP Date itself is "today" or in the future               -> the PST
                                                                  source
                                                                  had bad
                                                                  metadata

Run from Graph-Migrate/:

  .\.venv\Scripts\python.exe _inspect_dates.py -c config.toml ^
      --mailbox allysonp@jteatono365.onmicrosoft.com ^
      --folder sentitems --top 10
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


WELL_KNOWN = {"inbox", "sentitems", "drafts", "deleteditems", "junkemail", "outbox", "archive"}


def resolve_folder_id(graph: GraphClient, mailbox: str, folder: str) -> tuple[str, str]:
    """Return (folder_id, display_name). `folder` may be a well-known name
    (sentitems, inbox, ...) or the displayName of a top-level folder."""
    folder_l = folder.strip().lower()
    if folder_l in WELL_KNOWN:
        path = f"/users/{quote(mailbox)}/mailFolders/{folder_l}?$select=id,displayName"
        resp = graph.get(path, expect_status=(200,))
        f = resp.json()
        return f["id"], f.get("displayName") or folder

    path = (
        f"/users/{quote(mailbox)}/mailFolders/msgFolderRoot/childFolders"
        f"?$top=200&$select=id,displayName"
    )
    while path:
        resp = graph.get(path, expect_status=(200,))
        body = resp.json()
        for f in body.get("value", []):
            if (f.get("displayName") or "").lower() == folder_l:
                return f["id"], f.get("displayName")
        next_link = body.get("@odata.nextLink")
        path = next_link[len("https://graph.microsoft.com/v1.0"):] if next_link else None
    raise SystemExit(f"folder not found: {folder!r}")


def fetch_messages(graph: GraphClient, mailbox: str, folder_id: str, top: int) -> list[dict]:
    select = (
        "id,subject,sentDateTime,receivedDateTime,"
        "createdDateTime,lastModifiedDateTime,from,internetMessageId,"
        "isDraft,isRead"
    )
    # $expand pulls back PR_MESSAGE_FLAGS (0x0E07) and PR_SUBMIT_FLAGS (0x0E14)
    # alongside the message so we can compare what Graph stores in MAPI to
    # what it reports via the high-level isDraft / isRead properties.
    #
    # We pass the query via params= rather than baking it into the path so
    # httpx URL-encodes the embedded single quotes / spaces.  An earlier
    # version embedded the whole thing in the path and Graph silently
    # dropped singleValueExtendedProperties from the response (parser
    # confused by the unencoded quotes / '=' inside $filter()).
    path = f"/users/{quote(mailbox)}/mailFolders/{folder_id}/messages"
    params = {
        "$top": str(top),
        "$select": select,
        "$orderby": "receivedDateTime desc",
        "$expand": (
            "singleValueExtendedProperties("
            "$filter=id eq 'Integer 0x0E07' or id eq 'Integer 0x0E14')"
        ),
    }
    resp = graph.get(path, params=params, expect_status=(200,))
    return resp.json().get("value", [])


def _ext_prop_int(msg: dict, prop_id: str) -> int | None:
    """Pull the integer value of a single extended property out of the
    expanded singleValueExtendedProperties array on a message.

    Graph normalises the stored ID -- e.g. ``Integer 0x0E07`` may come
    back as ``Integer 0xE07`` (no leading zero in the tag).  Compare on
    a normalised form so we don't miss matches.
    """
    want = _norm_prop_id(prop_id)
    for ep in msg.get("singleValueExtendedProperties") or []:
        if _norm_prop_id(ep.get("id") or "") == want:
            v = ep.get("value")
            try:
                return int(v) if v is not None else None
            except (TypeError, ValueError):
                return None
    return None


def _norm_prop_id(pid: str) -> str:
    """Normalise an extended-property id like 'Integer 0x0E07' -> 'integer 0xe07'.
    Strips leading zeros from the hex tag so Graph's quirky echo of the
    same property under a slightly different spelling still matches."""
    pid = pid.strip().lower()
    if " 0x" in pid:
        head, tag = pid.split(" 0x", 1)
        tag = tag.lstrip("0") or "0"
        return f"{head} 0x{tag}"
    return pid


def fetch_smtp_date_header(graph: GraphClient, mailbox: str, msg_id: str) -> str | None:
    """Read the SMTP 'Date:' header off the message. This header is set by
    the original sender and is preserved verbatim through move/copy. If
    receivedDateTime disagrees with this, something is rewriting metadata."""
    path = (
        f"/users/{quote(mailbox)}/messages/{msg_id}"
        f"?$select=internetMessageHeaders"
    )
    try:
        resp = graph.get(path, expect_status=(200,))
        for h in resp.json().get("internetMessageHeaders") or []:
            if (h.get("name") or "").lower() == "date":
                return h.get("value")
    except GraphError as e:
        return f"<error {e.status}>"
    return None


def fmt(s: str | None, width: int) -> str:
    if not s:
        return "-".ljust(width)
    return (s if len(s) <= width else s[: width - 1] + "…").ljust(width)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n", 1)[0])
    ap.add_argument("-c", "--config", required=True)
    ap.add_argument("--mailbox", required=True, help="UPN of the mailbox to inspect")
    ap.add_argument(
        "--folder",
        default="sentitems",
        help="Well-known name (sentitems, inbox, drafts, deleteditems) "
             "or the displayName of a top-level folder. Default: sentitems",
    )
    ap.add_argument("--top", type=int, default=10, help="How many messages to inspect (default 10)")
    ap.add_argument(
        "--smtp",
        action="store_true",
        help="Also fetch and print the SMTP 'Date:' header per message "
             "(one extra Graph call per message; slower).",
    )
    ns = ap.parse_args()

    cfg = expand_user_paths(AppConfig.load(Path(ns.config)))

    logger.remove()
    logger.add(sys.stderr, level="INFO", format="{time:HH:mm:ss} | {level: <7} | {message}")

    pool = AppPool(cfg.apps)
    with GraphClient(pool, cfg.throttle) as graph:
        folder_id, folder_name = resolve_folder_id(graph, ns.mailbox, ns.folder)
        logger.info("Inspecting {!r} ({}) in {}", folder_name, folder_id[:18] + "…", ns.mailbox)
        messages = fetch_messages(graph, ns.mailbox, folder_id, ns.top)
        if not messages:
            logger.warning("Folder is empty.")
            return 0

        smtp_dates: dict[str, str | None] = {}
        if ns.smtp:
            for m in messages:
                smtp_dates[m["id"]] = fetch_smtp_date_header(graph, ns.mailbox, m["id"])

    def flag(v: object) -> str:
        # Compact yes/no/? for boolean columns. Helps spot isDraft=true at a glance.
        if v is True:
            return "yes"
        if v is False:
            return "no"
        return "?"

    def hex_or_dash(v: int | None) -> str:
        # Show the raw MAPI flag value as hex so the bits are easy to read
        # against the legend below the table.
        return "-" if v is None else f"0x{v:04x}"

    print()
    base_hdr = (
        f"{'subject':<40} {'sentDateTime':<22} {'receivedDateTime':<22} "
        f"{'lastMod':<22} {'draft':<5} {'read':<5} {'msgFlg':<8} {'subFlg':<8}"
    )
    hdr = base_hdr + (f" {'SMTP Date:':<32}" if ns.smtp else "")
    print(hdr)
    print("-" * len(hdr))

    for m in messages:
        subj = (m.get("subject") or "(no subject)").strip()
        msg_flags = _ext_prop_int(m, "Integer 0x0E07")
        sub_flags = _ext_prop_int(m, "Integer 0x0E14")
        row = (
            f"{fmt(subj, 40)} "
            f"{fmt(m.get('sentDateTime'), 22)} "
            f"{fmt(m.get('receivedDateTime'), 22)} "
            f"{fmt(m.get('lastModifiedDateTime'), 22)} "
            f"{flag(m.get('isDraft')):<5} "
            f"{flag(m.get('isRead')):<5} "
            f"{hex_or_dash(msg_flags):<8} "
            f"{hex_or_dash(sub_flags):<8}"
        )
        if ns.smtp:
            row += f" {fmt(smtp_dates.get(m['id']), 32)}"
        print(row)

    print()
    print(
        "How to read this:\n"
        "  draft / read are Graph's high-level booleans.\n"
        "  msgFlg = PR_MESSAGE_FLAGS (0x0E07).  Bits: 0x01 READ, 0x02 UNMODIFIED,\n"
        "           0x04 SUBMITTED, 0x08 UNSENT (the draft bit), 0x10 HASATTACH,\n"
        "           0x20 FROMME.\n"
        "  subFlg = PR_SUBMIT_FLAGS (0x0E14).  0x01 LOCKED, 0x02 PREPROCESS.\n"
        "  If draft=yes but msgFlg has UNSENT (0x08) clear, Graph is deriving\n"
        "  isDraft from something other than PR_MESSAGE_FLAGS (likely subFlg or\n"
        "  a cached field).  If msgFlg still has UNSENT set after _fix_drafts.py,\n"
        "  the PATCH didn't fully overwrite the property and we need a different\n"
        "  approach.\n"
        "  Date column meanings:\n"
        "    sent/received/SMTP all old, lastMod=today  -> fine, just move bumped lastMod\n"
        "    receivedDateTime=today, SMTP=old           -> import rewrote receive time\n"
        "    SMTP itself today/future                   -> bad PST source metadata"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
