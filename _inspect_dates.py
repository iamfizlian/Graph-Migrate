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
        "createdDateTime,lastModifiedDateTime,from,internetMessageId"
    )
    path = (
        f"/users/{quote(mailbox)}/mailFolders/{folder_id}/messages"
        f"?$top={top}&$select={select}&$orderby=receivedDateTime desc"
    )
    resp = graph.get(path, expect_status=(200,))
    return resp.json().get("value", [])


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

    print()
    if ns.smtp:
        hdr = (
            f"{'subject':<40} {'sentDateTime':<22} {'receivedDateTime':<22} "
            f"{'lastModifiedDateTime':<22} {'created':<22} {'SMTP Date:':<32}"
        )
    else:
        hdr = (
            f"{'subject':<40} {'sentDateTime':<22} {'receivedDateTime':<22} "
            f"{'lastModifiedDateTime':<22} {'created':<22}"
        )
    print(hdr)
    print("-" * len(hdr))

    for m in messages:
        subj = (m.get("subject") or "(no subject)").strip()
        row = (
            f"{fmt(subj, 40)} "
            f"{fmt(m.get('sentDateTime'), 22)} "
            f"{fmt(m.get('receivedDateTime'), 22)} "
            f"{fmt(m.get('lastModifiedDateTime'), 22)} "
            f"{fmt(m.get('createdDateTime'), 22)}"
        )
        if ns.smtp:
            row += f" {fmt(smtp_dates.get(m['id']), 32)}"
        print(row)

    print()
    print(
        "If 'sentDateTime'/'receivedDateTime'/SMTP Date are old but "
        "'lastModifiedDateTime' is today, the messages themselves are\n"
        "fine -- only the modified-stamp moved. If receivedDateTime is "
        "today/future while SMTP Date is old, something rewrote the\n"
        "receive time during import. If SMTP Date itself is today/future, "
        "the PST source had bad metadata."
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
