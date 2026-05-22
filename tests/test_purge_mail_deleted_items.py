"""Regression tests for ``purge-mail`` Deleted Items handling."""

from __future__ import annotations

import pytest

from jtet_pstmigrate.orchestrator import _mail_folder_is_deleted_items


@pytest.mark.parametrize(
    ("folder", "expected"),
    [
        ({"id": "a", "wellKnownFolderName": "deletedItems"}, True),
        ({"id": "a", "wellKnownFolderName": "deleteditems"}, True),
        ({"id": "a", "wellKnownFolderName": "inbox"}, False),
        ({"id": "a", "wellKnownFolderName": "sentItems"}, False),
        ({"id": "a"}, False),
        ({"id": "a", "wellKnownFolderName": None}, False),
    ],
)
def test_mail_folder_is_deleted_items(folder: dict, expected: bool) -> None:
    assert _mail_folder_is_deleted_items(folder) is expected
