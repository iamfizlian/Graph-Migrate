"""Unit tests for purge-mail Deleted Items skip logic."""

from jtet_pstmigrate.orchestrator import _purge_mail_skip_deleted_items_folder


def test_skip_when_folder_id_in_skip_set() -> None:
    assert _purge_mail_skip_deleted_items_folder(
        {"id": "abc", "wellKnownName": "inbox"},
        {"abc"},
    )


def test_skip_when_well_known_name_deleteditems_even_if_skip_set_empty() -> None:
    """If GET deleteditems failed, id-based skip is empty; wellKnownName still protects."""
    assert not _purge_mail_skip_deleted_items_folder(
        {"id": "xyz", "wellKnownName": "inbox"},
        set(),
    )
    assert _purge_mail_skip_deleted_items_folder(
        {"id": "xyz", "wellKnownName": "deleteditems"},
        set(),
    )


def test_skip_case_insensitive_well_known() -> None:
    assert _purge_mail_skip_deleted_items_folder(
        {"id": "x", "wellKnownName": "DeletedItems"},
        set(),
    )
