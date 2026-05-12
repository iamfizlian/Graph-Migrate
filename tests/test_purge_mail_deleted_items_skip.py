"""Regression: purge-mail must never treat Deleted Items as a drainable folder."""

from jtet_pstmigrate.orchestrator import _graph_mail_folder_is_deleted_items


def test_deleted_items_by_well_known_name_graph_casing() -> None:
    assert _graph_mail_folder_is_deleted_items(
        {"id": "x", "wellKnownFolderName": "deleteditems"}
    )
    assert _graph_mail_folder_is_deleted_items(
        {"id": "x", "wellKnownFolderName": "deletedItems"}
    )


def test_inbox_not_deleted_items() -> None:
    assert not _graph_mail_folder_is_deleted_items(
        {"id": "y", "wellKnownFolderName": "inbox"}
    )


def test_custom_folder_no_well_known() -> None:
    assert not _graph_mail_folder_is_deleted_items({"id": "z", "displayName": "Archive"})
