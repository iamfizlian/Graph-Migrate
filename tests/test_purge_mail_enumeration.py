"""Regression tests for purge-mail Graph enumeration behaviour."""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from jtet_pstmigrate.config import AppConfig, AuthConfig
from jtet_pstmigrate.graph_client import GraphError
from jtet_pstmigrate.orchestrator import Orchestrator


def _minimal_orch() -> Orchestrator:
    cfg = AppConfig(
        apps=[AuthConfig(tenant_id="t", client_id="c", client_secret="secret")],
    )
    orch = Orchestrator.__new__(Orchestrator)
    orch._cfg = cfg
    orch._state = MagicMock()
    orch._pool = MagicMock()
    orch._console = MagicMock()
    return orch


def _is_top_level_mailfolder_collection(path: str) -> bool:
    """True for GET /users/{upn}/mailFolders?... but not childFolders/messages/deleteditems."""
    return (
        "mailFolders" in path
        and "deleteditems" not in path
        and "/childFolders" not in path
        and "/messages" not in path
    )


def test_purge_mail_aborts_when_mailfolders_pagination_fails() -> None:
    """Second page of /mailFolders must not be dropped silently (partial tree purge)."""
    orch = _minimal_orch()
    graph = MagicMock()
    folder_pages: list[int] = []

    def get_side_effect(path: str, **kwargs):
        if "mailFolders/deleteditems" in path:
            r = MagicMock()
            r.json.return_value = {"id": "deleted-items-folder-id"}
            return r
        if _is_top_level_mailfolder_collection(path):
            folder_pages.append(1)
            if len(folder_pages) == 1:
                r = MagicMock()
                r.json.return_value = {
                    "value": [
                        {
                            "id": "folder-on-page-1",
                            "displayName": "OnlySeenIfBuggy",
                            "totalItemCount": 0,
                            "childFolderCount": 0,
                        }
                    ],
                    "@odata.nextLink": (
                        "https://graph.microsoft.com/v1.0/users/u%40x.com/mailFolders"
                        "?$skip=100&$top=100"
                    ),
                }
                return r
            raise GraphError(503, {"detail": "simulated page-2 failure"})

        pytest.fail(f"unexpected GET path: {path}")

    graph.get.side_effect = get_side_effect

    deleted, missing, errors = orch._purge_mail_one(graph, "u@x.com")

    assert errors == 1
    assert deleted == 0
    assert missing == 0
    graph.delete.assert_not_called()
