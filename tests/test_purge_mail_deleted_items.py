"""Regression tests for purge-mail Deleted Items handling."""

from __future__ import annotations

from unittest.mock import MagicMock

from jtet_pstmigrate.orchestrator import Orchestrator


class _Resp:
    __slots__ = ("_data",)

    def __init__(self, data: dict):
        self._data = data

    def json(self) -> dict:
        return self._data


def test_purge_mail_skips_deleted_items_when_well_known_get_fails():
    """Deleted Items must not be drained if GET …/mailFolders/deleteditems fails.

    Previously only ``skip_folder_ids`` (from that GET) gated the skip. When the
    GET failed, ``_walk`` still processed the Deleted Items row from the root
    ``/mailFolders`` enumeration and could delete or drain all messages there.
    """
    graph = MagicMock()
    mailbox = "user@example.com"
    message_enum_urls: list[str] = []

    def get_side_effect(path: str, **kwargs):
        if path.endswith("/mailFolders/deleteditems"):
            raise RuntimeError("simulated Graph failure")
        if "/messages" in path:
            message_enum_urls.append(path)
        if "/mailFolders?" in path and "childFolders" not in path:
            return _Resp(
                {
                    "value": [
                        {
                            "id": "deleted-folder-id",
                            "displayName": "Deleted Items",
                            "totalItemCount": 99,
                            "childFolderCount": 0,
                            "wellKnownFolderName": "deleteditems",
                        },
                        {
                            "id": "custom-folder-id",
                            "displayName": "Imported PST",
                            "totalItemCount": 0,
                            "childFolderCount": 0,
                        },
                    ]
                }
            )
        raise AssertionError(f"unexpected GET {path!r}")

    graph.get.side_effect = get_side_effect

    def delete_side_effect(path: str, **kwargs):
        if "deleted-folder-id" in path:
            raise AssertionError(f"must not DELETE Deleted Items folder: {path!r}")
        return _Resp({})

    graph.delete.side_effect = delete_side_effect

    cfg = MagicMock()
    cfg.migration.workers_per_mailbox = 2

    orch = object.__new__(Orchestrator)
    orch._cfg = cfg

    deleted, missing, errors = orch._purge_mail_one(graph, mailbox)

    assert deleted == 0
    assert missing == 0
    assert errors == 0
    assert not message_enum_urls, "must not enumerate messages under Deleted Items"
