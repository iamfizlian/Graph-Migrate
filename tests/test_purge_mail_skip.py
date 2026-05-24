"""Regression tests for purge-mail Deleted Items boundary."""

from jtet_pstmigrate.orchestrator import _purge_mail_should_skip_folder


def test_skip_when_folder_id_in_skip_set() -> None:
    assert _purge_mail_should_skip_folder(
        {"id": "abc", "wellKnownName": "inbox"},
        {"abc"},
    )


def test_skip_when_well_known_deleted_items() -> None:
    assert _purge_mail_should_skip_folder(
        {"id": "any", "wellKnownName": "deletedItems"},
        set(),
    )


def test_skip_well_known_case_insensitive() -> None:
    assert _purge_mail_should_skip_folder(
        {"id": "x", "wellKnownName": "DELETEDITEMS"},
        set(),
    )


def test_do_not_skip_regular_custom_folder() -> None:
    assert not _purge_mail_should_skip_folder(
        {"id": "f1", "displayName": "Imported", "wellKnownName": None},
        set(),
    )


def test_skip_set_beats_missing_well_known() -> None:
    assert _purge_mail_should_skip_folder({"id": "di-id"}, {"di-id"})
