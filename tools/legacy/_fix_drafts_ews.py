r"""Clear isDraft on imported messages via EWS (because Graph cannot).

Why this script exists
----------------------
We've verified empirically that the Microsoft Graph mail API silently
no-ops every attempt to clear MSGFLAG_UNSENT (0x08) on messages that
were created by POST-ing raw MIME to /users/{id}/messages:

  * PATCH {"isDraft": false}                              -> 200, ignored
  * PATCH PR_MESSAGE_FLAGS via singleValueExtendedProperties -> 200,
                                                              UNSENT preserved
  * POST /messages/{id}/copy                              -> copy inherits
                                                              isDraft=true

EWS UpdateItem with ConflictResolution=AlwaysOverwrite + a SetItemField
on PR_MESSAGE_FLAGS DOES clear UNSENT, because EWS exposes the underlying
MAPI store directly. This script walks every folder in a mailbox and
batches an UpdateItem per ~50 drafts.

EWS retirement
--------------
Microsoft has scheduled EWS retirement for Exchange Online for October
2026. For a one-time post-migration cleanup that runs now, that's
fine -- after we cut over the imported mailboxes we never call EWS
again. Future imports avoid the issue at the source (see "Future-proofing"
below).

Prerequisites (one-time, in Entra)
----------------------------------
At least one of your existing Entra app registrations needs an EWS
permission grant. The script uses the SAME client_id / secret / cert
from your config.toml; only the requested OAuth scope changes.

  1. Entra admin centre -> App registrations -> [your app] -> API permissions
  2. Add a permission -> APIs my organization uses
     -> "Office 365 Exchange Online"
  3. Application permissions -> "full_access_as_app"
  4. Grant admin consent

OPTIONAL: scope which mailboxes the app can touch via Exchange PowerShell:

  Connect-ExchangeOnline
  New-ApplicationAccessPolicy ``
    -AppId <client_id> ``
    -PolicyScopeGroupId <mail-group-of-migration-users> ``
    -AccessRight RestrictAccess ``
    -Description "Limit migration app to imported mailboxes"

Without the policy the app can access every mailbox in the tenant.
Use the policy in production tenants.

Usage
-----
  .\.venv\Scripts\python.exe _fix_drafts_ews.py -c config.toml --mailbox UPN --dry-run
  .\.venv\Scripts\python.exe _fix_drafts_ews.py -c config.toml --mailbox UPN
  .\.venv\Scripts\python.exe _fix_drafts_ews.py -c config.toml -m mapping.csv

Idempotent: re-running on a clean mailbox finds 0 drafts.

Future-proofing
---------------
The uploader currently creates messages by MIME-POST then /move, which
is what plants the UNSENT bit. To avoid the issue for future imports,
the upload path should be rewritten to use EWS CreateItem with
MessageDisposition=SaveOnly directly into the target folder; that lands
the message non-draft in one shot. Out of scope for this remediation
script.
"""
from __future__ import annotations

import argparse
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from html import escape
from pathlib import Path
from xml.etree import ElementTree as ET

import httpx
import msal
from loguru import logger

from jtet_pstmigrate.auth import _load_certificate
from jtet_pstmigrate.config import AppConfig, AuthConfig, expand_user_paths
from jtet_pstmigrate.orchestrator import load_mapping

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

EWS_ENDPOINT = "https://outlook.office365.com/EWS/Exchange.asmx"
EWS_SCOPE = ["https://outlook.office365.com/.default"]

# MAPI property tag for PR_MESSAGE_FLAGS, PT_LONG.
PROP_TAG_MESSAGE_FLAGS = "0x0E07"

# Bitmask we care about: MSGFLAG_UNSENT (the draft bit).
MSGFLAG_UNSENT = 0x08
# What we set PR_MESSAGE_FLAGS to: MSGFLAG_READ (clears UNSENT). Exchange
# computes HASATTACH/EverRead/etc. itself when we set just READ.
PR_MESSAGE_FLAGS_TARGET = 1

# Top-level folders we deliberately do not descend into:
#  - Drafts: real drafts the user authored should stay drafts
#  - Deleted Items: trash; their state doesn't matter
#  - Junk Email / Junk: spam quarantine
SKIP_TOP_LEVEL = frozenset({"Drafts", "Deleted Items", "Junk Email", "Junk"})

# Per-batch update size. EWS supports up to ~256 ItemChanges per call,
# but smaller batches give better progress reporting and lower retry cost.
UPDATE_BATCH_SIZE = 50

# FindItem page size. Server may cap; this is the max we ask for.
FIND_PAGE_SIZE = 500

NS_T = "{http://schemas.microsoft.com/exchange/services/2006/types}"
NS_M = "{http://schemas.microsoft.com/exchange/services/2006/messages}"
NS_S = "{http://schemas.xmlsoap.org/soap/envelope/}"


# ---------------------------------------------------------------------------
# Token acquisition (EWS scope, app-only)
# ---------------------------------------------------------------------------

def get_ews_token(app: AuthConfig) -> str:
    """Acquire an EWS-scoped token via msal client_credentials.

    The existing TokenProvider in auth.py is hardcoded to the Graph scope,
    so we acquire EWS tokens directly here. Same client_id / secret / cert
    used; only the scope differs.
    """
    authority = f"https://login.microsoftonline.com/{app.tenant_id}"
    if app.client_certificate_path:
        cert = _load_certificate(app.client_certificate_path, app.client_certificate_password)
        client = msal.ConfidentialClientApplication(
            client_id=app.client_id, authority=authority, client_credential=cert,
        )
    else:
        client = msal.ConfidentialClientApplication(
            client_id=app.client_id, authority=authority, client_credential=app.client_secret,
        )
    result = client.acquire_token_for_client(scopes=EWS_SCOPE)
    if "access_token" not in result:
        err = result.get("error_description") or result.get("error") or str(result)
        raise RuntimeError(f"EWS token acquisition failed for {app.display_name}: {err}")
    return result["access_token"]


# ---------------------------------------------------------------------------
# SOAP envelope construction
# ---------------------------------------------------------------------------

def _envelope(mailbox: str, body: str) -> str:
    """Wrap an EWS body fragment in a SOAP envelope with impersonation
    set to the target mailbox."""
    # Exchange2013_SP1 is the documented lowest-common-denominator value
    # accepted by Exchange Online; values like 'Exchange2016_SP1' are not
    # in the RequestServerVersion enum (cloud rejects them with
    # ErrorInvalidServerVersion). Everything we need (extended properties,
    # bitmask restrictions, batched UpdateItem) is exposed here.
    safe_mbx = escape(mailbox)
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<soap:Envelope '
        'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" '
        'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" '
        'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">'
        '<soap:Header>'
        '<t:RequestServerVersion Version="Exchange2013_SP1"/>'
        '<t:ExchangeImpersonation>'
        '<t:ConnectingSID>'
        f'<t:PrimarySmtpAddress>{safe_mbx}</t:PrimarySmtpAddress>'
        '</t:ConnectingSID>'
        '</t:ExchangeImpersonation>'
        '</soap:Header>'
        f'<soap:Body>{body}</soap:Body>'
        '</soap:Envelope>'
    )


def _find_folder_body(parent: str = "msgfolderroot", offset: int = 0) -> str:
    """FindFolder Traversal=Deep returns the entire folder tree under parent."""
    return (
        '<m:FindFolder Traversal="Deep">'
        '<m:FolderShape>'
        '<t:BaseShape>IdOnly</t:BaseShape>'
        '<t:AdditionalProperties>'
        '<t:FieldURI FieldURI="folder:DisplayName"/>'
        '<t:FieldURI FieldURI="folder:FolderClass"/>'
        '<t:FieldURI FieldURI="folder:TotalCount"/>'
        '</t:AdditionalProperties>'
        '</m:FolderShape>'
        f'<m:IndexedPageFolderView MaxEntriesReturned="500" Offset="{offset}" BasePoint="Beginning"/>'
        '<m:ParentFolderIds>'
        f'<t:DistinguishedFolderId Id="{parent}"/>'
        '</m:ParentFolderIds>'
        '</m:FindFolder>'
    )


def _find_drafts_body(folder_id: str, change_key: str, offset: int) -> str:
    """FindItem with NOT(Excludes(PR_MESSAGE_FLAGS, 0x08)) filters for items
    whose UNSENT bit is set -- i.e. drafts, server-side."""
    ck_attr = f' ChangeKey="{escape(change_key)}"' if change_key else ""
    return (
        '<m:FindItem Traversal="Shallow">'
        '<m:ItemShape>'
        '<t:BaseShape>IdOnly</t:BaseShape>'
        '<t:AdditionalProperties>'
        '<t:FieldURI FieldURI="item:Subject"/>'
        f'<t:ExtendedFieldURI PropertyTag="{PROP_TAG_MESSAGE_FLAGS}" PropertyType="Integer"/>'
        '</t:AdditionalProperties>'
        '</m:ItemShape>'
        f'<m:IndexedPageItemView MaxEntriesReturned="{FIND_PAGE_SIZE}" '
        f'Offset="{offset}" BasePoint="Beginning"/>'
        '<m:Restriction>'
        '<t:Not>'
        '<t:Excludes>'
        f'<t:ExtendedFieldURI PropertyTag="{PROP_TAG_MESSAGE_FLAGS}" PropertyType="Integer"/>'
        f'<t:Bitmask Value="{MSGFLAG_UNSENT}"/>'
        '</t:Excludes>'
        '</t:Not>'
        '</m:Restriction>'
        '<m:ParentFolderIds>'
        f'<t:FolderId Id="{escape(folder_id)}"{ck_attr}/>'
        '</m:ParentFolderIds>'
        '</m:FindItem>'
    )


def _update_change(item_id: str, change_key: str) -> str:
    """One <ItemChange> that sets PR_MESSAGE_FLAGS = 1 (clears UNSENT)."""
    return (
        '<t:ItemChange>'
        f'<t:ItemId Id="{escape(item_id)}" ChangeKey="{escape(change_key)}"/>'
        '<t:Updates>'
        '<t:SetItemField>'
        f'<t:ExtendedFieldURI PropertyTag="{PROP_TAG_MESSAGE_FLAGS}" PropertyType="Integer"/>'
        '<t:Message>'
        '<t:ExtendedProperty>'
        f'<t:ExtendedFieldURI PropertyTag="{PROP_TAG_MESSAGE_FLAGS}" PropertyType="Integer"/>'
        f'<t:Value>{PR_MESSAGE_FLAGS_TARGET}</t:Value>'
        '</t:ExtendedProperty>'
        '</t:Message>'
        '</t:SetItemField>'
        '</t:Updates>'
        '</t:ItemChange>'
    )


def _update_items_body(changes: list[tuple[str, str]]) -> str:
    """UpdateItem with MessageDisposition=SaveOnly + AlwaysOverwrite. SaveOnly
    is critical: without it EWS may try to SUBMIT (send!) a message whose
    flags suggest it's pending submission."""
    inner = "".join(_update_change(i, c) for i, c in changes)
    return (
        '<m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite">'
        '<m:ItemChanges>'
        f'{inner}'
        '</m:ItemChanges>'
        '</m:UpdateItem>'
    )


# ---------------------------------------------------------------------------
# EWS client (very thin wrapper)
# ---------------------------------------------------------------------------

class EwsError(Exception):
    pass


# SOAP faultcodes that are permanent client/server bugs -- retrying buys
# nothing. The faultcode comes back as e.g. 'a:ErrorInvalidServerVersion'
# in the SOAP envelope; we strip the namespace prefix before matching.
_TERMINAL_FAULTCODES = frozenset({
    "ErrorInvalidServerVersion",     # bad RequestServerVersion value
    "ErrorAccessDenied",              # missing full_access_as_app or app policy
    "ErrorInvalidUserOid",            # mailbox doesn't exist / typo
    "ErrorNonExistentMailbox",
    "ErrorImpersonateUserDenied",     # impersonation not authorised
    "ErrorSchemaValidation",          # we built a malformed SOAP body
    "ErrorInvalidIdMalformed",
    "ErrorInvalidPropertyRequest",    # asked for a non-requestable FieldURI
    "ErrorInvalidArgument",
    "ErrorMissingArgument",
    "ErrorInvalidRequest",
    "ErrorInvalidPropertyAppend",
    "ErrorInvalidPropertyDelete",
    "ErrorInvalidPropertySet",
    "ErrorInvalidPropertyUpdateSentMessage",  # KEY: cloud's lock on sent-message props
    "ErrorObjectTypeChanged",
    "ErrorIncorrectUpdatePropertyCount",
    "ErrorUnsupportedMimeConversion",
    "ErrorRequestStreamTooLarge",
    "ErrorInvalidExtendedProperty",
    "ErrorInvalidExtendedPropertyValue",
})


def _parse_soap_fault(body: bytes) -> tuple[str | None, str | None]:
    """Pull faultcode + faultstring out of a SOAP 500 response. Returns
    (None, None) if the body isn't a SOAP fault we recognise."""
    try:
        root = ET.fromstring(body)
    except ET.ParseError:
        return None, None
    fault = root.find(f"{NS_S}Body/{NS_S}Fault")
    if fault is None:
        return None, None
    code_el = fault.find("faultcode")
    msg_el = fault.find("faultstring")
    code = (code_el.text or "") if code_el is not None else ""
    # Strip 'a:' / 's:' / 'soap:' style prefixes that EWS uses on faultcodes.
    if ":" in code:
        code = code.split(":", 1)[1]
    msg = (msg_el.text or "") if msg_el is not None else ""
    return code or None, msg or None


class EwsClient:
    """One-mailbox EWS client.

    Uses ExchangeImpersonation (set in the SOAP envelope) and an X-AnchorMailbox
    header (avoids cross-forest redirect chatter for cloud Exchange). Auto-retries
    transient 5xx and 429 with exponential backoff.
    """

    def __init__(self, token: str, mailbox: str, http: httpx.Client):
        self._token = token
        self._mailbox = mailbox
        self._http = http
        self._log = logger.bind(ctx=f"ews[{mailbox}]")

    def call(self, soap_body: str) -> ET.Element:
        """POST a SOAP envelope, return the parsed <soap:Body> element.

        Raises EwsError on terminal failure (after retries).
        """
        envelope = _envelope(self._mailbox, soap_body)
        headers = {
            "Authorization": f"Bearer {self._token}",
            "Content-Type": 'text/xml; charset=utf-8',
            "X-AnchorMailbox": self._mailbox,
            "Accept": "text/xml",
        }
        attempt = 0
        while True:
            try:
                resp = self._http.post(EWS_ENDPOINT, content=envelope, headers=headers)
            except (httpx.TransportError, httpx.TimeoutException) as e:
                if attempt >= 5:
                    raise EwsError(f"transport error after retries: {e}") from e
                self._log.warning("transport error (attempt {}): {} -- backing off", attempt + 1, e)
                time.sleep(min(30.0, 1.5 ** attempt))
                attempt += 1
                continue

            if resp.status_code == 200:
                try:
                    root = ET.fromstring(resp.content)
                except ET.ParseError as e:
                    raise EwsError(f"could not parse EWS response: {e}") from e
                body = root.find(f"{NS_S}Body")
                if body is None:
                    raise EwsError("EWS response missing soap:Body")
                return body

            if resp.status_code in (429, 503) and attempt < 8:
                # Exchange Online throttles EWS too, separate from Graph.
                retry_after = resp.headers.get("Retry-After")
                wait = float(retry_after) if retry_after else min(60.0, 2.0 ** attempt)
                self._log.warning("EWS {} -- backing off {:.1f}s", resp.status_code, wait)
                time.sleep(wait)
                attempt += 1
                continue

            # 500 responses from EWS commonly carry a SOAP Fault with a
            # specific faultcode. Permanent client errors (bad server
            # version, bad scope, malformed request) are not worth
            # retrying -- surface them immediately. Anything else falls
            # through to the generic 5xx retry below.
            if resp.status_code == 500:
                fault_code, fault_msg = _parse_soap_fault(resp.content)
                if fault_code and fault_code in _TERMINAL_FAULTCODES:
                    raise EwsError(
                        f"EWS terminal error {fault_code}: {fault_msg or '(no message)'}"
                    )

            if 500 <= resp.status_code < 600 and attempt < 5:
                wait = min(30.0, 2.0 ** attempt)
                self._log.warning("EWS {} -- retrying after {:.1f}s", resp.status_code, wait)
                time.sleep(wait)
                attempt += 1
                continue

            snippet = resp.text[:500] if resp.text else ""
            raise EwsError(
                f"EWS HTTP {resp.status_code}: {snippet}"
            )

    def find_all_folders(self) -> list[dict]:
        """Walk every folder under msgfolderroot; return list of dicts with
        id / change_key / display_name / folder_class / total_count.

        Skips the well-known Drafts and Deleted Items subtrees by name.
        Pagination is supported even though Exchange Online rarely has >500
        folders in a single mailbox."""
        folders: list[dict] = []
        offset = 0
        while True:
            body = self.call(_find_folder_body("msgfolderroot", offset))
            response_messages = body.find(f"{NS_M}FindFolderResponse/{NS_M}ResponseMessages")
            if response_messages is None:
                raise EwsError("FindFolder response missing ResponseMessages")
            ff_resp = response_messages.find(f"{NS_M}FindFolderResponseMessage")
            if ff_resp is None:
                raise EwsError("FindFolder response missing FindFolderResponseMessage")
            cls = ff_resp.attrib.get("ResponseClass", "")
            if cls != "Success":
                code = (ff_resp.findtext(f"{NS_M}ResponseCode") or "").strip()
                msg = (ff_resp.findtext(f"{NS_M}MessageText") or "").strip()
                raise EwsError(f"FindFolder failed ({cls}): {code} -- {msg}")
            root_folder = ff_resp.find(f"{NS_M}RootFolder")
            if root_folder is None:
                break
            for f in root_folder.findall(f"{NS_T}Folders/{NS_T}Folder"):
                folder_id_el = f.find(f"{NS_T}FolderId")
                if folder_id_el is None:
                    continue
                folders.append({
                    "id": folder_id_el.attrib.get("Id", ""),
                    "change_key": folder_id_el.attrib.get("ChangeKey", ""),
                    "display_name": f.findtext(f"{NS_T}DisplayName") or "",
                    "folder_class": f.findtext(f"{NS_T}FolderClass") or "",
                    "total_count": int(f.findtext(f"{NS_T}TotalCount") or 0),
                })
            includes_last = root_folder.attrib.get("IncludesLastItemInRange", "true").lower() == "true"
            if includes_last:
                break
            offset = int(root_folder.attrib.get("IndexedPagingOffset", str(offset + len(folders))))
        return folders

    def find_drafts_in_folder(self, folder_id: str, folder_change_key: str) -> list[tuple[str, str]]:
        """Return [(item_id, change_key), ...] for every message in the folder
        that has MSGFLAG_UNSENT set (i.e. drafts)."""
        out: list[tuple[str, str]] = []
        offset = 0
        while True:
            body = self.call(_find_drafts_body(folder_id, folder_change_key, offset))
            response_messages = body.find(f"{NS_M}FindItemResponse/{NS_M}ResponseMessages")
            if response_messages is None:
                raise EwsError("FindItem response missing ResponseMessages")
            fi_resp = response_messages.find(f"{NS_M}FindItemResponseMessage")
            if fi_resp is None:
                raise EwsError("FindItem response missing FindItemResponseMessage")
            cls = fi_resp.attrib.get("ResponseClass", "")
            if cls != "Success":
                code = (fi_resp.findtext(f"{NS_M}ResponseCode") or "").strip()
                msg = (fi_resp.findtext(f"{NS_M}MessageText") or "").strip()
                # ErrorAccessDenied here usually means EWS scope wasn't granted.
                raise EwsError(f"FindItem failed ({cls}): {code} -- {msg}")
            root_folder = fi_resp.find(f"{NS_M}RootFolder")
            if root_folder is None:
                break
            for msg in root_folder.findall(f"{NS_T}Items/{NS_T}Message"):
                item_id_el = msg.find(f"{NS_T}ItemId")
                if item_id_el is None:
                    continue
                out.append((
                    item_id_el.attrib.get("Id", ""),
                    item_id_el.attrib.get("ChangeKey", ""),
                ))
            includes_last = root_folder.attrib.get("IncludesLastItemInRange", "true").lower() == "true"
            if includes_last:
                break
            offset = int(root_folder.attrib.get("IndexedPagingOffset", str(offset + len(out))))
        return out

    def update_drafts_batch(self, changes: list[tuple[str, str]]) -> tuple[int, list[str]]:
        """UpdateItem on a batch of (item_id, change_key). Returns (succeeded,
        errors). Each ItemChange yields its own response message; a single
        bad item doesn't fail the batch."""
        if not changes:
            return 0, []
        body = self.call(_update_items_body(changes))
        response_messages = body.find(f"{NS_M}UpdateItemResponse/{NS_M}ResponseMessages")
        if response_messages is None:
            raise EwsError("UpdateItem response missing ResponseMessages")
        succeeded = 0
        errors: list[str] = []
        for r in response_messages.findall(f"{NS_M}UpdateItemResponseMessage"):
            cls = r.attrib.get("ResponseClass", "")
            if cls == "Success":
                succeeded += 1
            else:
                code = (r.findtext(f"{NS_M}ResponseCode") or "").strip()
                msg = (r.findtext(f"{NS_M}MessageText") or "").strip()
                errors.append(f"{code}: {msg}")
        return succeeded, errors


# ---------------------------------------------------------------------------
# Per-mailbox driver
# ---------------------------------------------------------------------------

def _chunks(seq: list, n: int):
    for i in range(0, len(seq), n):
        yield seq[i:i + n]


def _is_skip_path(display_name: str, ancestry: list[str]) -> bool:
    """A folder is skipped if itself OR any of its ancestors is in the
    SKIP_TOP_LEVEL set. We don't have parent-ids in our flat folder list,
    so we infer from name; in practice these names are well-known and
    don't collide with user-named subfolders at the top level."""
    if display_name in SKIP_TOP_LEVEL:
        return True
    return any(a in SKIP_TOP_LEVEL for a in ancestry)


def fix_mailbox(app: AuthConfig, mailbox: str, *, dry_run: bool) -> dict:
    log = logger.bind(ctx=f"ews-fix[{mailbox}]")
    stats: dict = {
        "mailbox": mailbox,
        "folders_walked": 0,
        "drafts_found": 0,
        "fixed": 0,
        "failed": 0,
        "errors": [],
    }
    try:
        token = get_ews_token(app)
    except Exception as e:
        log.error("Token acquisition failed: {}", e)
        stats["errors"].append(f"token: {e}")
        return stats

    timeout = httpx.Timeout(120.0, connect=30.0)
    with httpx.Client(timeout=timeout, http2=False) as http:
        ews = EwsClient(token, mailbox, http)
        try:
            folders = ews.find_all_folders()
        except EwsError as e:
            log.error("FindFolder failed: {}", e)
            stats["errors"].append(str(e))
            return stats

        # We have a flat list. Build a quick "is this folder under a skipped
        # subtree?" check by name -- the EWS folder tree doesn't include
        # parent ids in the response unless we ask, but the well-known names
        # we skip ('Drafts', 'Deleted Items', ...) don't collide with normal
        # user folder names at the top level in practice.
        candidate_folders = [
            f for f in folders
            if f["folder_class"] in ("IPF.Note", "")  # only mail folders, not contacts/cal
            and f["display_name"] not in SKIP_TOP_LEVEL
        ]
        log.info(
            "Found {} mail folder(s) to scan ({} total folders in mailbox).",
            len(candidate_folders), len(folders),
        )

        for folder in candidate_folders:
            try:
                drafts = ews.find_drafts_in_folder(folder["id"], folder["change_key"])
            except EwsError as e:
                log.warning("FindItem failed for {!r}: {}", folder["display_name"], e)
                stats["errors"].append(f"{folder['display_name']}: {e}")
                continue
            stats["folders_walked"] += 1
            if not drafts:
                continue
            stats["drafts_found"] += len(drafts)
            log.info("  {!r}: {} draft(s)", folder["display_name"], len(drafts))
            if dry_run:
                continue
            for chunk in _chunks(drafts, UPDATE_BATCH_SIZE):
                try:
                    ok, errs = ews.update_drafts_batch(chunk)
                except EwsError as e:
                    log.warning("    UpdateItem batch failed: {}", e)
                    stats["failed"] += len(chunk)
                    stats["errors"].append(str(e))
                    continue
                stats["fixed"] += ok
                stats["failed"] += len(chunk) - ok
                if errs:
                    sample = errs[0]
                    log.warning("    {} item(s) failed in batch (first: {})",
                                len(errs), sample[:120])

    if stats["drafts_found"] == 0:
        log.info("No imported-as-draft messages found; mailbox is clean.")
    elif dry_run:
        log.info("DRY RUN -- would clear {} draft(s) across {} folder(s).",
                 stats["drafts_found"], stats["folders_walked"])
    else:
        log.info("Cleared {} of {} draft(s); {} failed.",
                 stats["fixed"], stats["drafts_found"], stats["failed"])
    return stats


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(
        description=(
            "Clear isDraft on imported messages via EWS. Requires "
            "Office 365 Exchange Online > full_access_as_app on at least "
            "one Entra app in your config."
        )
    )
    ap.add_argument("-c", "--config", required=True)
    ap.add_argument("-m", "--mapping", help="mapping.csv path")
    ap.add_argument("--mailbox", action="append", help="UPN of a single mailbox (repeatable)")
    ap.add_argument("--dry-run", action="store_true", help="Walk + count, no UpdateItem calls.")
    ap.add_argument("--app", help="Display name of the app entry to use (default: first).")
    ns = ap.parse_args()

    cfg = expand_user_paths(AppConfig.load(Path(ns.config)))

    if ns.mailbox:
        mailboxes = sorted(set(ns.mailbox))
    elif ns.mapping:
        mailboxes = sorted({r.target_mailbox for r in load_mapping(Path(ns.mapping))})
    else:
        ap.error("Provide either --mapping or one or more --mailbox UPNs.")

    logger.remove()
    logger.add(sys.stderr, level="INFO",
               format="{time:HH:mm:ss} | {level: <7} | {extra[ctx]} | {message}")

    if ns.app:
        app = next((a for a in cfg.apps if a.display_name == ns.app), None)
        if app is None:
            available = ", ".join(a.display_name for a in cfg.apps)
            logger.bind(ctx="ews-fix").error(
                "App {!r} not found in config. Available: {}", ns.app, available,
            )
            return 2
    else:
        app = cfg.apps[0]
    logger.bind(ctx="ews-fix").info(
        "Using Entra app {!r} for EWS auth (tenant={}).",
        app.display_name, app.tenant_id,
    )

    if ns.dry_run:
        logger.bind(ctx="ews-fix").warning("DRY RUN -- no UpdateItem calls will be issued.")

    parallelism = min(cfg.migration.max_parallel_mailboxes, len(mailboxes)) or 1
    logger.bind(ctx="ews-fix").info(
        "Processing {} mailbox(es) with {} parallel worker(s).",
        len(mailboxes), parallelism,
    )

    results: list[dict] = []
    with ThreadPoolExecutor(max_workers=parallelism, thread_name_prefix="ews-fix") as ex:
        futs = {
            ex.submit(fix_mailbox, app, m, dry_run=ns.dry_run): m
            for m in mailboxes
        }
        for fut in as_completed(futs):
            m = futs[fut]
            try:
                results.append(fut.result())
            except Exception as e:
                logger.bind(ctx=f"ews-fix[{m}]").exception("Crashed: {}", e)
                results.append({"mailbox": m, "errors": [str(e)]})

    print()
    hdr = f"{'mailbox':<46} {'folders':>8} {'drafts':>8} {'fixed':>8}  status"
    print(hdr)
    print("-" * len(hdr))
    exit_code = 0
    for r in sorted(results, key=lambda x: x["mailbox"]):
        errs = r.get("errors") or []
        if errs and r.get("fixed", 0) == 0 and r.get("drafts_found", 0) == 0:
            print(f"{r['mailbox']:<46} {'-':>8} {'-':>8} {'-':>8}  ERROR: {errs[0][:40]}")
            exit_code = 1
        else:
            note = "ok" if not r.get("failed") else f"{r['failed']} failed"
            print(
                f"{r['mailbox']:<46} {r.get('folders_walked', 0):>8} "
                f"{r.get('drafts_found', 0):>8} {r.get('fixed', 0):>8}  {note}"
            )
            if r.get("failed"):
                exit_code = 1
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
