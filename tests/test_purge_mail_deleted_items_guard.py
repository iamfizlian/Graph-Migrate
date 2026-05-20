"""Regression tests for purge-mail Deleted Items safeguards."""

from jtet_pstmigrate.orchestrator import _purge_mail_folder_is_deleted_items


def test_skip_deleted_items_by_well_known_name_without_id_set() -> None:
    folder = {"id": "folder-abc", "wellKnownFolderName": "deleteditems"}
    assert _purge_mail_folder_is_deleted_items(folder, set()) is True


def test_skip_deleted_items_by_id_in_skip_set() -> None:
    folder = {"id": "known-deleted-id", "wellKnownFolderName": None}
    assert _purge_mail_folder_is_deleted_items(folder, {"known-deleted-id"}) is True


def test_do_not_skip_regular_folder() -> None:
    folder = {"id": "inbox-id", "wellKnownFolderName": "inbox"}
    assert _purge_mail_folder_is_deleted_items(folder, set()) is False


def test_well_known_name_is_case_insensitive() -> None:
    folder = {"id": "x", "wellKnownFolderName": "DeletedItems"}
    assert _purge_mail_folder_is_deleted_items(folder, set()) is True
