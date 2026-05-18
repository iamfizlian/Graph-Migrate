"""Unit tests for purge-mail Deleted Items skip logic."""

from jtet_pstmigrate.orchestrator import _purge_mail_skip_folder_tree


def test_skip_when_folder_id_in_skip_set() -> None:
    assert _purge_mail_skip_folder_tree({"id": "abc"}, {"abc"}) is True


def test_skip_when_well_known_deleted_items_even_if_skip_set_empty() -> None:
    folder = {"id": "folder-guid", "wellKnownFolderName": "deleteditems"}
    assert _purge_mail_skip_folder_tree(folder, set()) is True


def test_well_known_name_is_case_insensitive() -> None:
    folder = {"id": "x", "wellKnownFolderName": "DeletedItems"}
    assert _purge_mail_skip_folder_tree(folder, set()) is True


def test_inbox_not_skipped() -> None:
    folder = {"id": "inbox-id", "wellKnownFolderName": "inbox"}
    assert _purge_mail_skip_folder_tree(folder, set()) is False


def test_custom_folder_not_skipped() -> None:
    folder = {"id": "q", "displayName": "Projects"}
    assert _purge_mail_skip_folder_tree(folder, set()) is False
