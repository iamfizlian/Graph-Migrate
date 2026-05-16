"""Regression: purge-mail must never drain the Deleted Items folder."""

from jtet_pstmigrate.orchestrator import _purge_mail_skip_deleted_items_folder


def test_skip_when_id_in_preflight_set() -> None:
    fid = "folder-id-1"
    folder = {"id": fid, "wellKnownName": "inbox"}
    assert _purge_mail_skip_deleted_items_folder(folder, {fid}) is True


def test_skip_by_well_known_name_when_preflight_empty() -> None:
    """If GET /deleteditems failed, skip_folder_ids is empty — still skip by Graph flag."""
    folder = {"id": "abc", "wellKnownName": "deleteditems"}
    assert _purge_mail_skip_deleted_items_folder(folder, set()) is True


def test_skip_well_known_name_case_insensitive() -> None:
    folder = {"id": "x", "wellKnownName": "DeletedItems"}
    assert _purge_mail_skip_deleted_items_folder(folder, set()) is True


def test_do_not_skip_inbox() -> None:
    folder = {"id": "inbox-id", "wellKnownName": "inbox"}
    assert _purge_mail_skip_deleted_items_folder(folder, set()) is False


def test_do_not_skip_custom_folder() -> None:
    folder = {"id": "custom", "displayName": "Imported PST"}
    assert _purge_mail_skip_deleted_items_folder(folder, set()) is False
