"""Upload one .vcf contact to a Graph mailbox.

We POST a single JSON contact document to ``/users/{upn}/contacts``
(the user's default contact folder). vCards from libpst are vCard 3.0
with a fairly stable subset of properties (FN/N/ORG/TITLE/EMAIL/TEL/
ADR/BDAY/NOTE/URL); we hand-roll the parse here rather than pulling in
``vobject`` because the surface we touch is small and we want explicit
control over how Outlook-isms (preamble headers, weird TYPE params)
get folded into Graph's flat shape.

What we map:
  - FN/N -> displayName / surname / givenName / middleName / title
  - ORG -> companyName (and department appended to job title if present)
  - TITLE -> jobTitle
  - EMAIL[*] -> emailAddresses[] (Graph caps at 3; we keep the first 3)
  - TEL with TYPE=HOME/WORK/CELL -> homePhones[] / businessPhones[] /
    mobilePhone (single)
  - ADR with TYPE=HOME/WORK/OTHER -> homeAddress / businessAddress /
    otherAddress
  - BDAY -> birthday (ISO datetime at midnight UTC)
  - NOTE -> personalNotes
  - URL -> businessHomePage (Graph has only one URL field on contact)

Anything we don't recognise is dropped silently -- we'd rather lose
ANNIVERSARY than crash the whole run on a quirky property.
"""

from __future__ import annotations

import datetime as _dt
import re
from urllib.parse import quote

from loguru import logger

from jtet_pstmigrate.graph_client import GraphClient
from jtet_pstmigrate.pst_reader import ExtractedContact


class ContactUploadResult:
    __slots__ = ("bytes_uploaded", "graph_contact_id", "via")

    def __init__(self, graph_contact_id: str, bytes_uploaded: int, via: str = "json"):
        self.graph_contact_id = graph_contact_id
        self.bytes_uploaded = bytes_uploaded
        self.via = via


class ContactUploadError(RuntimeError):
    """Raised when a .vcf can't be turned into a usable Graph contact."""


class ContactUploader:
    def __init__(self, graph: GraphClient, mailbox: str):
        self._graph = graph
        self._mailbox = mailbox
        self._log = logger.bind(ctx=f"contactupload[{mailbox}]")

    def upload(
        self, contact: ExtractedContact, *, app_id: str
    ) -> ContactUploadResult:
        raw = contact.file_path.read_bytes()
        doc = _vcard_to_graph_contact(raw)
        if doc is None:
            raise ContactUploadError(
                f"no usable vCard data in {contact.file_path.name}"
            )
        path = f"/users/{quote(self._mailbox)}/contacts"
        resp = self._graph.post(
            path, json=doc, expect_status=(201,), app_id=app_id
        )
        contact_id = resp.json()["id"]
        return ContactUploadResult(
            graph_contact_id=contact_id, bytes_uploaded=len(raw)
        )


# -- vCard parsing --------------------------------------------------------

# RFC 2425 line folding: a CRLF (or LF) followed by SP or TAB is a
# continuation of the previous logical line.
_FOLD_RE = re.compile(rb"\r?\n[ \t]")
_ESC_MAP = {"n": "\n", "N": "\n", ",": ",", ";": ";", ":": ":", "\\": "\\"}


def _parse_vcard(raw: bytes) -> dict[str, list[tuple[dict[str, list[str]], str]]] | None:
    """Parse a vCard blob into ``{NAME: [(params, value), ...]}``.

    Returns None when no VCARD body is present. Multiple instances of
    the same property (e.g. several EMAILs) are preserved in document
    order so the caller can pick the right one by TYPE.
    """
    unfolded = _FOLD_RE.sub(b"", raw)
    text = _decode_text(unfolded)

    props: dict[str, list[tuple[dict[str, list[str]], str]]] = {}
    in_vcard = False

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        upper = line.upper()
        if upper.startswith("BEGIN:VCARD"):
            in_vcard = True
            continue
        if upper.startswith("END:VCARD"):
            return props if (props or in_vcard) else None
        if not in_vcard:
            # Some readpst outputs have stray BOM / blank lines / a
            # leading "X-..." style header before BEGIN:VCARD; skip.
            continue
        if ":" not in line:
            continue
        head, value = line.split(":", 1)
        head_parts = head.split(";")
        name = head_parts[0].strip().upper()
        if not name:
            continue
        params: dict[str, list[str]] = {}
        for raw_param in head_parts[1:]:
            raw_param = raw_param.strip()
            if not raw_param:
                continue
            if "=" in raw_param:
                key, val = raw_param.split("=", 1)
                key = key.strip().upper()
                # TYPE=HOME,VOICE -- comma-separated values are common.
                params.setdefault(key, []).extend(
                    v.strip().upper() for v in val.split(",") if v.strip()
                )
            else:
                # vCard 2.1 bare param ("HOME" without "TYPE=") -- treat
                # as a TYPE entry so downstream sees it uniformly.
                params.setdefault("TYPE", []).append(raw_param.strip().upper())

        props.setdefault(name, []).append((params, _unescape(value)))

    # File ended without END:VCARD but we did see BEGIN -- accept what
    # we got rather than throwing the whole record away.
    return props or None


def _decode_text(raw: bytes) -> str:
    """Decode a vCard payload as text. Outlook/libpst uses UTF-8 in modern
    output but older PSTs occasionally land in cp1252; try the common
    suspects in order, then fall back to lossy UTF-8 decode."""
    for enc in ("utf-8-sig", "utf-8", "cp1252", "iso-8859-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _unescape(value: str) -> str:
    """Apply vCard value escapes: \\n -> newline, \\, \\; \\\\ literals."""
    out: list[str] = []
    i = 0
    while i < len(value):
        c = value[i]
        if c == "\\" and i + 1 < len(value):
            nxt = value[i + 1]
            out.append(_ESC_MAP.get(nxt, nxt))
            i += 2
        else:
            out.append(c)
            i += 1
    return "".join(out)


# -- Graph mapping --------------------------------------------------------

# Graph contact field length limits (from the schema; truncating instead
# of failing keeps a long-noted contact importable).
_MAX_NAME = 64
_MAX_DISPLAY = 256
_MAX_EMAIL = 256
_MAX_NOTE = 32_000
_MAX_PHONES_PER_CATEGORY = 3   # Graph caps homePhones/businessPhones at 3 each.


def _vcard_to_graph_contact(raw: bytes) -> dict | None:
    """Convert a raw vCard blob to a Graph contact JSON, or return None.

    None means "nothing useful in this file" -- the caller marks the
    item ``skipped`` so we don't keep retrying it on every run.
    """
    props = _parse_vcard(raw)
    if not props:
        return None

    doc: dict = {}

    fn = _first(props, "FN")
    n_value = _first(props, "N")
    if n_value:
        # vCard N: family;given;additional;prefixes;suffixes
        n_parts = n_value.split(";")
        family = n_parts[0].strip() if len(n_parts) > 0 else ""
        given = n_parts[1].strip() if len(n_parts) > 1 else ""
        additional = n_parts[2].strip() if len(n_parts) > 2 else ""
        prefix = n_parts[3].strip() if len(n_parts) > 3 else ""
        if family:
            doc["surname"] = family[:_MAX_NAME]
        if given:
            doc["givenName"] = given[:_MAX_NAME]
        if additional:
            doc["middleName"] = additional[:_MAX_NAME]
        if prefix:
            # Graph's `title` is the honorific prefix (Mr/Mrs/Dr); the
            # job title goes in `jobTitle` (set below from TITLE).
            doc["title"] = prefix[:_MAX_NAME]
    if fn:
        doc["displayName"] = fn[:_MAX_DISPLAY]
    elif "givenName" in doc or "surname" in doc:
        # Synthesise displayName from N when FN is missing -- otherwise
        # the contact shows up as nameless in OWA.
        doc["displayName"] = " ".join(
            x for x in (doc.get("givenName"), doc.get("surname")) if x
        )[:_MAX_DISPLAY]

    org = _first(props, "ORG")
    if org:
        # ORG: Company;Department;Sub-department -- we keep the first
        # segment as company name. Department isn't separately exposed
        # on Graph contact, so we don't do anything special with it.
        doc["companyName"] = org.split(";")[0].strip()[:_MAX_NAME]

    title = _first(props, "TITLE")
    if title:
        doc["jobTitle"] = title[:_MAX_NAME]

    emails: list[dict] = []
    for _params, value in props.get("EMAIL", []):
        addr = value.strip()
        if not addr or "@" not in addr:
            continue
        emails.append({"address": addr[:_MAX_EMAIL], "name": fn or addr})
        if len(emails) >= 3:
            # Graph rejects more than three emails per contact.
            break
    if emails:
        doc["emailAddresses"] = emails

    home_phones: list[str] = []
    business_phones: list[str] = []
    mobile: str | None = None
    for params, value in props.get("TEL", []):
        num = value.strip()
        if not num:
            continue
        types = params.get("TYPE", [])
        if "CELL" in types or "MOBILE" in types:
            if mobile is None:
                mobile = num[:64]
        elif "WORK" in types:
            if len(business_phones) < _MAX_PHONES_PER_CATEGORY:
                business_phones.append(num[:64])
        elif "HOME" in types:
            if len(home_phones) < _MAX_PHONES_PER_CATEGORY:
                home_phones.append(num[:64])
        else:
            # Untyped or unusual TYPE -- bucket into business by default
            # since most untyped vCard phones in office use are work.
            if len(business_phones) < _MAX_PHONES_PER_CATEGORY:
                business_phones.append(num[:64])
    if home_phones:
        doc["homePhones"] = home_phones
    if business_phones:
        doc["businessPhones"] = business_phones
    if mobile:
        doc["mobilePhone"] = mobile

    for params, value in props.get("ADR", []):
        addr = _adr_to_graph(value)
        if not addr:
            continue
        types = params.get("TYPE", [])
        if "WORK" in types and "businessAddress" not in doc:
            doc["businessAddress"] = addr
        elif "HOME" in types and "homeAddress" not in doc:
            doc["homeAddress"] = addr
        elif "otherAddress" not in doc:
            doc["otherAddress"] = addr

    bday = _first(props, "BDAY")
    if bday:
        iso = _bday_to_graph(bday)
        if iso:
            doc["birthday"] = iso

    note = _first(props, "NOTE")
    if note:
        doc["personalNotes"] = note[:_MAX_NOTE]

    url = _first(props, "URL")
    if url:
        # Graph contact has only one URL slot (`businessHomePage`); pick
        # the first vCard URL and put it there.
        doc["businessHomePage"] = url[:256]

    # Bail out if we don't have *any* identifying info -- a Graph
    # contact with no name and no email is just noise.
    if not (doc.get("displayName") or doc.get("givenName") or doc.get("surname")
            or doc.get("emailAddresses")):
        return None

    return doc


def _first(props: dict, name: str) -> str:
    """Return the first value for ``name`` (uppercase), or ''."""
    items = props.get(name)
    if not items:
        return ""
    return items[0][1]


def _adr_to_graph(value: str) -> dict | None:
    """vCard ADR fields: PObox;Extended;Street;City;Region;Postal;Country.

    We drop PObox + Extended (Graph contact has no equivalent fields)
    and concatenate the others into the matching Graph keys.
    """
    parts = value.split(";")
    while len(parts) < 7:
        parts.append("")
    pobox, extended, street, city, region, postal, country = (p.strip() for p in parts[:7])

    street_combined = " ".join(p for p in (extended, street) if p) or pobox
    out: dict = {}
    if street_combined:
        out["street"] = street_combined[:256]
    if city:
        out["city"] = city[:64]
    if region:
        out["state"] = region[:64]
    if postal:
        out["postalCode"] = postal[:32]
    if country:
        out["countryOrRegion"] = country[:64]
    return out or None


_BDAY_RE = re.compile(r"^(?P<y>\d{4})-?(?P<m>\d{2})-?(?P<d>\d{2})")


def _bday_to_graph(value: str) -> str | None:
    """Normalise BDAY (date-only or full ISO) to Graph's birthday shape.

    Graph stores birthday as a date-time at midnight UTC. The day part
    is what surfaces in OWA, so we discard any time component the vCard
    happened to include.
    """
    m = _BDAY_RE.match(value.strip())
    if not m:
        return None
    try:
        d = _dt.date(int(m["y"]), int(m["m"]), int(m["d"]))
    except ValueError:
        return None
    return f"{d:%Y-%m-%d}T00:00:00Z"
