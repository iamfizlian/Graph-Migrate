"""Upload a single .eml message into a Graph mailFolder.

We always POST a JSON message document directly to
  /users/{upn}/mailFolders/{folderId}/messages
with ``singleValueExtendedProperties`` populated at create time so the
message lands in the destination folder, not in Drafts. Attachments are
uploaded individually after the parent message is created; small ones
inline, large ones via ``createUploadSession``.

We deliberately do not use the MIME-import endpoint (POST /users/{id}/messages
with text/plain MIME). That path always lands the message in Drafts first,
so the only way to clear the draft flag is a PATCH after the move. That
PATCH is silently rejected by some tenants, which is exactly the bug the
``Fix-DraftsViaOutlook.ps1`` and ``_fix_drafts_*.py`` scripts in this repo
exist to clean up.

Date preservation
-----------------
Graph stamps ``receivedDateTime`` / ``sentDateTime`` / ``createdDateTime`` /
``lastModifiedDateTime`` with the current server time on create. To make
Outlook show the original timestamps we attach the corresponding MAPI
extended properties to the create call:

    PR_CLIENT_SUBMIT_TIME      0x0039  -> sentDateTime
    PR_MESSAGE_DELIVERY_TIME   0x0E06  -> receivedDateTime
    PR_CREATION_TIME           0x3007  -> createdDateTime
    PR_LAST_MODIFICATION_TIME  0x3008  -> lastModifiedDateTime

Some tenants reject writes to PR_CREATION_TIME / PR_LAST_MODIFICATION_TIME
at create time. On a 400 from the create call we retry with PR_MESSAGE_FLAGS
ONLY (preserving the no-draft invariant) and PATCH the dates afterwards.

Draft-flag fix
--------------
PR_MESSAGE_FLAGS (tag 0x0E07, PT_LONG) is set to MSGFLAG_READ (0x01) in
``singleValueExtendedProperties`` on the create call, so MSGFLAG_UNSENT
(0x08) and MSGFLAG_SUBMIT (0x04) are cleared before the message ever
exists on the server. Messages created this way report ``isDraft=false``
from the moment they appear and never need a follow-up flag PATCH.
"""

from __future__ import annotations

import base64
import datetime
import email
import email.utils
from email import policy
from urllib.parse import quote

from loguru import logger

from jtet_pstmigrate.graph_client import GraphClient, GraphError
from jtet_pstmigrate.pst_reader import ExtractedMessage

# MAPI property tags (PT_SYSTIME = 0x0040, dropped from the Graph "id" form).
PROP_CLIENT_SUBMIT_TIME = "SystemTime 0x0039"
PROP_MESSAGE_DELIVERY_TIME = "SystemTime 0x0E06"
PROP_CREATION_TIME = "SystemTime 0x3007"
PROP_LAST_MODIFICATION_TIME = "SystemTime 0x3008"
# PR_MESSAGE_FLAGS, PT_LONG. Setting this to MSGFLAG_READ (0x01) clears
# MSGFLAG_UNSENT (0x08) and MSGFLAG_SUBMIT (0x04), which is what makes
# Graph stop reporting the message as isDraft=true.
PROP_MESSAGE_FLAGS = "Integer 0x0E07"
MSGFLAG_READ = "1"


class UploadResult:
    __slots__ = ("bytes_uploaded", "graph_message_id", "via")

    def __init__(self, graph_message_id: str, bytes_uploaded: int, via: str):
        self.graph_message_id = graph_message_id
        self.bytes_uploaded = bytes_uploaded
        self.via = via


class MessageUploader:
    def __init__(self, graph: GraphClient, mailbox: str, large_threshold_bytes: int):
        self._graph = graph
        self._mailbox = mailbox
        self._threshold = large_threshold_bytes
        self._log = logger.bind(ctx=f"upload[{mailbox}]")

    def upload(self, msg: ExtractedMessage, folder_id: str, *, app_id: str) -> UploadResult:
        # Always use the JSON create-in-target-folder path so PR_MESSAGE_FLAGS
        # (clear MSGFLAG_UNSENT) goes in at create time, in singleValueExtendedProperties,
        # and the message never lands under Drafts. The old MIME path created the
        # message under Drafts and tried to clear the flag with a PATCH afterwards;
        # that PATCH is unreliable across tenants -- existing _fix_drafts_*.py and
        # Fix-DraftsViaOutlook.ps1 in this repo are the empirical proof.
        raw = msg.file_path.read_bytes()
        date_props = _date_props_from_eml(raw)
        return self._upload_via_create_then_attach(
            raw, folder_id, date_props=date_props, app_id=app_id
        )

    def _patch_dates(
        self, msg_id: str, date_props: list[dict[str, str]], *, app_id: str
    ) -> None:
        """PATCH the message with original send/receive/create/modify timestamps.

        Best-effort: the message is already in the mailbox at this point; if
        the date PATCH fails we'd rather end up with a successfully-imported
        message that has the wrong sentDateTime than mark it failed (which
        would cause a duplicate on the next retry).

        Some tenants/SKUs reject writes to PR_CREATION_TIME and
        PR_LAST_MODIFICATION_TIME (they're considered server-managed). On 400
        we retry with just the user-visible Sent/Received pair, which every
        Exchange tenant accepts. On any other error we log and move on.
        """
        path = f"/users/{quote(self._mailbox)}/messages/{msg_id}"
        try:
            self._graph.patch(
                path,
                json={"singleValueExtendedProperties": date_props},
                expect_status=(200,),
                app_id=app_id,
            )
            return
        except GraphError as e:
            if e.status != 400:
                self._log.warning(
                    "Date PATCH failed for {} ({}); message kept with import-time dates",
                    msg_id, e.status,
                )
                return
            # 400: try the visible-only fallback below
        except Exception as e:
            self._log.warning(
                "Date PATCH unexpected error on {}: {}; message kept with import-time dates",
                msg_id, e,
            )
            return

        visible_only = [
            p for p in date_props
            if p["id"] in (
                PROP_MESSAGE_DELIVERY_TIME,
                PROP_CLIENT_SUBMIT_TIME,
                PROP_MESSAGE_FLAGS,
            )
        ]
        if not visible_only:
            return
        try:
            self._graph.patch(
                path,
                json={"singleValueExtendedProperties": visible_only},
                expect_status=(200,),
                app_id=app_id,
            )
        except Exception as e:
            self._log.warning(
                "Date PATCH (visible-only) failed for {}: {}; message kept with import-time dates",
                msg_id, e,
            )

    def _upload_via_create_then_attach(
        self, raw: bytes, folder_id: str, *, date_props: list[dict[str, str]], app_id: str
    ) -> UploadResult:
        """Fallback for very large messages: create the shell, then attach.

        Graph's `/createUploadSession` on attachments allows multi-MB streaming
        without hitting the 150 MB single-request cap.
        """
        parsed = email.message_from_bytes(raw, policy=policy.default)

        message_doc = {
            "subject": (parsed.get("Subject") or "")[:255],
            "body": _body_doc(parsed),
            "from": _addr(parsed.get("From")),
            "toRecipients": _addr_list(parsed.get_all("To")),
            "ccRecipients": _addr_list(parsed.get_all("Cc")),
            "bccRecipients": _addr_list(parsed.get_all("Bcc")),
            "internetMessageId": (parsed.get("Message-ID") or "").strip("<> "),
        }
        if date_props:
            # Set dates at create time — saves a PATCH round-trip on this path.
            message_doc["singleValueExtendedProperties"] = date_props
        # Drop empty keys so Graph doesn't reject odd-shaped recipients
        message_doc = {k: v for k, v in message_doc.items() if v not in (None, [], "", {})}

        path = f"/users/{quote(self._mailbox)}/mailFolders/{folder_id}/messages"
        # Two-tier fallback that preserves the no-draft invariant:
        #   1. create with msg_flags + date props together
        #   2. on 400, retry with msg_flags ONLY (drop date props -- some tenants
        #      reject server-managed PR_CREATION_TIME / PR_LAST_MODIFICATION_TIME),
        #      then PATCH the dates after.
        # We never fall back to "create with no extended props"; that path used to
        # let the message exist as a draft until a follow-up PATCH cleared the
        # flag, and the PATCH is unreliable on locked-down tenants.
        msg_flags_only = [p for p in date_props if p["id"] == PROP_MESSAGE_FLAGS]
        date_only_props = [p for p in date_props if p["id"] != PROP_MESSAGE_FLAGS]
        create_with_dates_failed = False
        try:
            resp = self._graph.post(path, json=message_doc, expect_status=(201,), app_id=app_id)
        except GraphError as e:
            if e.status == 400 and date_only_props and msg_flags_only:
                self._log.debug(
                    "Create-with-dates rejected; retrying with msg_flags only and PATCHing dates after"
                )
                create_with_dates_failed = True
                fallback_doc = dict(message_doc)
                fallback_doc["singleValueExtendedProperties"] = msg_flags_only
                resp = self._graph.post(path, json=fallback_doc, expect_status=(201,), app_id=app_id)
            else:
                raise
        msg_id = resp.json()["id"]
        if create_with_dates_failed and date_only_props:
            self._patch_dates(msg_id, date_only_props, app_id=app_id)  # best-effort

        attached_bytes = 0
        for part in parsed.walk():
            if part.is_multipart():
                continue
            payload = part.get_payload(decode=True) or b""
            if not payload or part.get_content_disposition() not in ("attachment", "inline"):
                continue
            attached_bytes += self._upload_attachment(
                msg_id,
                filename=part.get_filename() or "attachment.bin",
                content_bytes=payload,
                content_type=part.get_content_type(),
                app_id=app_id,
            )

        return UploadResult(
            graph_message_id=msg_id,
            bytes_uploaded=len(raw) + attached_bytes,
            via="json+attach",
        )

    def _upload_attachment(
        self, msg_id: str, filename: str, content_bytes: bytes, content_type: str, *, app_id: str
    ) -> int:
        if len(content_bytes) <= self._threshold:
            path = f"/users/{quote(self._mailbox)}/messages/{msg_id}/attachments"
            self._graph.post(
                path,
                json={
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": filename[:255],
                    "contentType": content_type,
                    "contentBytes": base64.b64encode(content_bytes).decode("ascii"),
                },
                expect_status=(201,),
                app_id=app_id,
            )
            return len(content_bytes)

        # Large attachment: open upload session, chunk-upload via PUT.
        session_path = f"/users/{quote(self._mailbox)}/messages/{msg_id}/attachments/createUploadSession"
        resp = self._graph.post(
            session_path,
            json={
                "AttachmentItem": {
                    "attachmentType": "file",
                    "name": filename[:255],
                    "size": len(content_bytes),
                    "contentType": content_type,
                }
            },
            expect_status=(201, 200),
            app_id=app_id,
        )
        upload_url = resp.json()["uploadUrl"]
        chunk = 5 * 1024 * 1024
        for start in range(0, len(content_bytes), chunk):
            end = min(start + chunk, len(content_bytes)) - 1
            piece = content_bytes[start : end + 1]
            self._graph.put(
                upload_url,
                content=piece,
                headers={
                    "Content-Length": str(len(piece)),
                    "Content-Range": f"bytes {start}-{end}/{len(content_bytes)}",
                    "Content-Type": "application/octet-stream",
                    # Bypass our auth header — upload URLs are pre-signed
                    "Authorization": "",
                },
                expect_status=(200, 201, 202),
                # Pre-signed URL: any app_id is fine; pass for stats attribution
                app_id=app_id,
            )
        return len(content_bytes)


def _addr(header_val: str | None) -> dict | None:
    if not header_val:
        return None
    addrs = email.utils.getaddresses([header_val])
    if not addrs:
        return None
    name, addr = addrs[0]
    if not addr:
        return None
    return {"emailAddress": {"name": name or addr, "address": addr}}


def _addr_list(headers: list[str] | None) -> list[dict]:
    if not headers:
        return []
    out: list[dict] = []
    for h in headers:
        for name, addr in email.utils.getaddresses([h]):
            if addr:
                out.append({"emailAddress": {"name": name or addr, "address": addr}})
    return out


def _body_doc(parsed) -> dict:
    plain = None
    html = None
    if parsed.is_multipart():
        for part in parsed.walk():
            if part.get_content_disposition() in ("attachment", "inline"):
                continue
            ctype = part.get_content_type()
            if ctype == "text/plain" and plain is None:
                plain = part.get_content() if hasattr(part, "get_content") else _decode_text(part)
            elif ctype == "text/html" and html is None:
                html = part.get_content() if hasattr(part, "get_content") else _decode_text(part)
    else:
        ctype = parsed.get_content_type()
        if ctype == "text/html":
            html = parsed.get_content() if hasattr(parsed, "get_content") else _decode_text(parsed)
        else:
            plain = parsed.get_content() if hasattr(parsed, "get_content") else _decode_text(parsed)

    if html:
        return {"contentType": "html", "content": html}
    if plain:
        return {"contentType": "text", "content": plain}
    return {"contentType": "text", "content": ""}


def _decode_text(part) -> str:
    payload = part.get_payload(decode=True) or b""
    charset = part.get_content_charset() or "utf-8"
    try:
        return payload.decode(charset, errors="replace")
    except LookupError:
        return payload.decode("utf-8", errors="replace")


def _date_props_from_eml(raw: bytes) -> list[dict[str, str]]:
    """Build the singleValueExtendedProperties payload to set original
    timestamps and clear the draft flag.

    The PR_MESSAGE_FLAGS entry is always returned, even when the .eml has
    no usable Date header -- the dates can fall back to import-time, but
    we never want to leave a message flagged as a draft. Date entries are
    only included when a real Date header parses successfully.
    """
    msg_flags = {"id": PROP_MESSAGE_FLAGS, "value": MSGFLAG_READ}

    try:
        parsed = email.message_from_bytes(raw, policy=policy.compat32)
    except Exception:
        return [msg_flags]

    sent_iso = _header_to_iso(parsed.get("Date"))
    received_iso = _header_to_iso(_last_received_date(parsed)) or sent_iso
    if not sent_iso and not received_iso:
        return [msg_flags]
    sent_iso = sent_iso or received_iso

    return [
        msg_flags,
        {"id": PROP_CLIENT_SUBMIT_TIME, "value": sent_iso},
        {"id": PROP_MESSAGE_DELIVERY_TIME, "value": received_iso},
        {"id": PROP_CREATION_TIME, "value": received_iso},
        {"id": PROP_LAST_MODIFICATION_TIME, "value": received_iso},
    ]


def _header_to_iso(header_val: str | None) -> str | None:
    """Parse an RFC 2822 date header into Graph-friendly ISO 8601 (UTC)."""
    if not header_val:
        return None
    try:
        dt = email.utils.parsedate_to_datetime(header_val)
    except (TypeError, ValueError):
        return None
    if dt is None:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=datetime.timezone.utc)
    return dt.astimezone(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")


def _last_received_date(parsed) -> str | None:
    """The last `Received:` header has the actual delivery date (after `;`).

    Headers are listed top-down (most-recent-first) so the *last* one in
    iteration order is the closest to the original delivery.
    """
    received_headers = parsed.get_all("Received") or []
    if not received_headers:
        return None
    last = received_headers[-1]
    if ";" in last:
        return last.rsplit(";", 1)[1].strip()
    return None
