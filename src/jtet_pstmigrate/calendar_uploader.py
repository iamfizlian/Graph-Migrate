"""Upload one .ics appointment to a Graph mailbox as a calendar event.

We POST a single JSON event document to ``/users/{upn}/calendar/events``
(the user's default calendar). Recurring events from the source PST
become individual single-occurrence events in the destination calendar
for v1 -- the original RRULE text is preserved in the event body so no
data is lost, but the destination calendar doesn't recreate the
recurrence pattern. Building a full RRULE -> Graph ``recurrence``
mapping is a follow-up.

Times are normalised to UTC. All-day events use ``isAllDay=true`` so
they render correctly across viewers' time zones regardless of the UTC
offset stored on the date.
"""

from __future__ import annotations

import datetime as _dt
from urllib.parse import quote

from icalendar import Calendar
from loguru import logger

from jtet_pstmigrate.graph_client import GraphClient
from jtet_pstmigrate.pst_reader import ExtractedAppointment


class CalendarUploadResult:
    __slots__ = ("bytes_uploaded", "graph_event_id", "via")

    def __init__(self, graph_event_id: str, bytes_uploaded: int, via: str = "json"):
        self.graph_event_id = graph_event_id
        self.bytes_uploaded = bytes_uploaded
        self.via = via


class CalendarUploadError(RuntimeError):
    """Raised when an .ics file can't be turned into a usable Graph event."""


class CalendarUploader:
    def __init__(self, graph: GraphClient, mailbox: str):
        self._graph = graph
        self._mailbox = mailbox
        self._log = logger.bind(ctx=f"calupload[{mailbox}]")

    def upload(
        self, appt: ExtractedAppointment, *, app_id: str
    ) -> CalendarUploadResult:
        raw = appt.file_path.read_bytes()
        event_doc = _ics_to_graph_event(raw)
        if event_doc is None:
            raise CalendarUploadError(
                f"no usable VEVENT in {appt.file_path.name}"
            )
        path = f"/users/{quote(self._mailbox)}/calendar/events"
        resp = self._graph.post(
            path, json=event_doc, expect_status=(201,), app_id=app_id
        )
        event_id = resp.json()["id"]
        return CalendarUploadResult(
            graph_event_id=event_id, bytes_uploaded=len(raw)
        )


# -- ICS parsing ----------------------------------------------------------

def _ics_to_graph_event(raw: bytes) -> dict | None:
    """Parse an .ics blob and return a Graph-shaped event document, or None."""
    try:
        cal = Calendar.from_ical(raw)
    except Exception:
        return None

    # Skip VTIMEZONE / VFREEBUSY / etc.; pick the first VEVENT.
    event = next((c for c in cal.walk() if c.name == "VEVENT"), None)
    if event is None:
        return None

    dtstart = event.get("DTSTART")
    if dtstart is None:
        return None

    summary = str(event.get("SUMMARY") or "(no subject)")[:255]
    description = str(event.get("DESCRIPTION") or "")
    location = str(event.get("LOCATION") or "")

    start_dt, start_tz, is_all_day = _to_graph_date(dtstart)
    dtend = event.get("DTEND")
    if dtend is not None:
        end_dt, end_tz, _ = _to_graph_date(dtend)
    else:
        # No DTEND: synthesise a sensible default. iCal also allows DURATION,
        # but readpst's output reliably populates DTEND so this branch is rare.
        if is_all_day:
            end_dt, end_tz = _add_days(start_dt, 1), start_tz
        else:
            end_dt, end_tz = _add_minutes(start_dt, 30), start_tz

    # Pseudo-all-day detection.
    #
    # Outlook stores all-day events as a 24h DATE-TIME range at local
    # midnight, and readpst preserves that representation: e.g., for a
    # US Eastern user, an all-day event on March 26 lands as
    # DTSTART:20190326T040000Z / DTEND:20190327T040000Z (04:00 UTC == 00:00
    # EDT). Without this fix-up we'd post that as a 24-hour TIMED event,
    # which OWA renders as "12am to 12am" rather than as a clean all-day
    # banner.
    #
    # The heuristic: if the start/end delta is exactly 24h and both have
    # the same time-of-day, treat as all-day and rewrite to date-only at
    # 00:00 UTC. We use the UTC date of DTSTART; that matches the user's
    # local date for any UTC-negative time zone (i.e. all of the Americas).
    # For UTC-positive zones the local date can be one day later than the
    # UTC date, and we'd guess wrong by a day -- left as a follow-up
    # because it requires reading mailboxSettings.timeZone per mailbox.
    if not is_all_day:
        all_day_dt = _maybe_pseudo_all_day(dtstart, dtend)
        if all_day_dt is not None:
            is_all_day = True
            start_dt = f"{all_day_dt:%Y-%m-%d}T00:00:00.0000000"
            end_dt = f"{all_day_dt + _dt.timedelta(days=1):%Y-%m-%d}T00:00:00.0000000"
            start_tz = end_tz = "UTC"

    rrule = event.get("RRULE")
    if rrule is not None:
        description = _append_rrule_note(description, rrule)

    doc: dict = {
        "subject": summary,
        "body": {"contentType": "text", "content": description},
        "start": {"dateTime": start_dt, "timeZone": start_tz},
        "end":   {"dateTime": end_dt,   "timeZone": end_tz},
        "isAllDay": is_all_day,
    }
    if location:
        doc["location"] = {"displayName": location[:255]}

    organizer = event.get("ORGANIZER")
    if organizer is not None:
        org = _parse_cal_address(organizer)
        if org:
            doc["organizer"] = {"emailAddress": org}

    raw_attendees = event.get("ATTENDEE")
    attendees: list[dict] = []
    if raw_attendees is not None:
        if not isinstance(raw_attendees, list):
            raw_attendees = [raw_attendees]
        for a in raw_attendees:
            addr = _parse_cal_address(a)
            if addr:
                attendees.append({"emailAddress": addr, "type": "required"})
    if attendees:
        doc["attendees"] = attendees

    return doc


def _to_graph_date(prop) -> tuple[str, str, bool]:
    """Convert an icalendar value to ``(dateTime, timeZone, is_all_day)``.

    Outputs are always normalised to UTC for two reasons:
      1. Graph accepts UTC universally; IANA names from icalendar's tzinfo
         can be inconsistent (sometimes they're ``zoneinfo`` names, sometimes
         Olson aliases that Graph rejects).
      2. UTC + isAllDay is the only combination that renders identically
         for all viewers regardless of their own time zone.
    """
    dt = prop.dt if hasattr(prop, "dt") else prop

    # Date-only -> all-day. iCalendar uses ``date`` (not ``datetime``) for
    # all-day events. Graph still wants a dateTime, but isAllDay=true makes
    # the time component irrelevant for display.
    if isinstance(dt, _dt.date) and not isinstance(dt, _dt.datetime):
        return (
            f"{dt.year:04d}-{dt.month:02d}-{dt.day:02d}T00:00:00.0000000",
            "UTC",
            True,
        )
    if isinstance(dt, _dt.datetime):
        if dt.tzinfo is None:
            # Floating local time. Treat as UTC -- the alternative is guessing
            # the originator's zone, which is worse than a known wrong default.
            iso = dt.strftime("%Y-%m-%dT%H:%M:%S.0000000")
        else:
            iso = dt.astimezone(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.0000000")
        return (iso, "UTC", False)
    # Unknown type -- defensive fallthrough.
    return (str(dt)[:27], "UTC", False)


def _maybe_pseudo_all_day(dtstart, dtend) -> _dt.date | None:
    """Detect Outlook's 24h DATE-TIME representation of an all-day event.

    Returns the calendar date the event falls on, or None if the start/end
    pair doesn't look like a pseudo-all-day. Both ends of the interval
    must be ``datetime`` (not ``date``), exactly 24 hours apart, and have
    identical time-of-day -- that combination means "midnight to midnight
    in some time zone", which is the canonical Outlook all-day shape.
    """
    if dtend is None:
        return None
    raw_start = getattr(dtstart, "dt", dtstart)
    raw_end = getattr(dtend, "dt", dtend)
    if not (isinstance(raw_start, _dt.datetime) and isinstance(raw_end, _dt.datetime)):
        return None
    delta = raw_end - raw_start
    if int(delta.total_seconds()) != 86400:
        return None
    if raw_start.timetz() != raw_end.timetz():
        return None
    # Require the time-of-day to be on an exact hour with no fractional
    # seconds. Every IANA time zone has an integer-hour offset from UTC
    # (the half-hour zones like India and Newfoundland still produce a
    # XX:30:00 time-of-day, which we accept), so any local-midnight
    # converted to UTC produces XX:MM:00.000000. A 10:00:23.456 boundary
    # is clearly not "midnight somewhere" -- almost certainly a real
    # 24-hour timed event we shouldn't reshape.
    t = raw_start.time()
    if t.second != 0 or t.microsecond != 0:
        return None
    if raw_start.tzinfo is not None:
        raw_start = raw_start.astimezone(_dt.timezone.utc)
    return raw_start.date()


def _add_days(iso: str, days: int) -> str:
    dt = _dt.datetime.strptime(iso[:19], "%Y-%m-%dT%H:%M:%S")
    return (dt + _dt.timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%S.0000000")


def _add_minutes(iso: str, minutes: int) -> str:
    dt = _dt.datetime.strptime(iso[:19], "%Y-%m-%dT%H:%M:%S")
    return (dt + _dt.timedelta(minutes=minutes)).strftime("%Y-%m-%dT%H:%M:%S.0000000")


def _parse_cal_address(prop) -> dict | None:
    """Convert an ATTENDEE/ORGANIZER iCal value to Graph emailAddress shape.

    Returns None when no SMTP address is recoverable -- the calling code
    drops the field entirely rather than POSTing an empty address.
    """
    if prop is None:
        return None
    try:
        value = str(prop)
    except Exception:
        return None
    addr = value
    if addr.lower().startswith("mailto:"):
        addr = addr[len("mailto:"):]
    addr = addr.strip("<>\t \n")
    if not addr or "@" not in addr:
        return None
    name = addr
    try:
        params = getattr(prop, "params", None)
        if params:
            cn = params.get("CN") or params.get("cn")
            if cn:
                name = str(cn)
    except Exception:
        pass
    return {"name": (name or addr)[:255], "address": addr}


def _append_rrule_note(description: str, rrule) -> str:
    """Stash the original RRULE text inside the event body so the recurrence
    is at least documented even when we don't recreate the pattern itself.

    icalendar exposes RRULE as a ``vRecur``-like dict; we serialise it with
    a stable separator so the result is human-readable in OWA.
    """
    parts: list[str] = []
    try:
        for key, value in rrule.items():
            if isinstance(value, list):
                parts.append(f"{key}={','.join(str(v) for v in value)}")
            else:
                parts.append(f"{key}={value}")
    except Exception:
        return description
    if not parts:
        return description
    note = (
        "[Imported from PST -- original recurrence rule: "
        + ";".join(parts)
        + "]"
    )
    if description:
        return description + "\n\n" + note
    return note
