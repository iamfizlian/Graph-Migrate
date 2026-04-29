"""Guards for purge-mail: never touch Graph's Deleted Items well-known folder."""

from jtet_pstmigrate.orchestrator import _folder_is_graph_deleted_items


def test_deleted_items_detected_by_well_known_name() -> None:
    assert _folder_is_graph_deleted_items({"wellKnownFolderName": "deleteditems"})
    assert _folder_is_graph_deleted_items({"wellKnownFolderName": "DeletedItems"})
    assert _folder_is_graph_deleted_items({"wellKnownFolderName": " deleteditems "})


def test_other_folders_not_treated_as_deleted_items() -> None:
    assert not _folder_is_graph_deleted_items({})
    assert not _folder_is_graph_deleted_items({"wellKnownFolderName": "inbox"})
    assert not _folder_is_graph_deleted_items({"displayName": "Deleted Items"})
