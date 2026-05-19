"""Regression tests for RRULE -> Graph recurrence mapping."""

from datetime import datetime

import pytest
from icalendar import vDatetime, vRecur

from jtet_pstmigrate.calendar_uploader import _rrule_to_graph_recurrence


@pytest.mark.parametrize(
    ("rrule_ical", "expect_index"),
    [
        ("FREQ=MONTHLY;BYDAY=1MO", "first"),
        ("FREQ=MONTHLY;BYDAY=4WE", "fourth"),
        ("FREQ=MONTHLY;BYDAY=-1FR", "last"),
    ],
)
def test_monthly_byday_ordinals_supported_by_graph(rrule_ical: str, expect_index: str) -> None:
    rrule = vRecur.from_ical(rrule_ical)
    dtstart = vDatetime(datetime(2026, 1, 5, 12, 0, 0))
    out = _rrule_to_graph_recurrence(rrule, dtstart)
    assert out is not None
    assert out["pattern"]["index"] == expect_index


def test_monthly_fifth_weekday_not_mapped_to_last() -> None:
    """BYDAY=5TH is not the same as last Thursday; Graph has no 'fifth' index."""
    rrule = vRecur.from_ical("FREQ=MONTHLY;BYDAY=5TH")
    dtstart = vDatetime(datetime(2026, 1, 29, 10, 0, 0))
    assert _rrule_to_graph_recurrence(rrule, dtstart) is None


def test_yearly_fifth_weekday_not_mapped_to_last() -> None:
    rrule = vRecur.from_ical("FREQ=YEARLY;BYMONTH=3;BYDAY=5TU")
    dtstart = vDatetime(datetime(2026, 3, 31, 9, 0, 0))
    assert _rrule_to_graph_recurrence(rrule, dtstart) is None
