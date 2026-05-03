"""Regression: purge-mail must not treat partial Graph pagination as complete."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from jtet_pstmigrate.config import AppConfig, AuthConfig, MigrationConfig, PathsConfig, ThrottleConfig
from jtet_pstmigrate.graph_client import GraphError
from jtet_pstmigrate.orchestrator import MappingRow, Orchestrator
from jtet_pstmigrate.state import StateStore


def _minimal_cfg(tmp_path: Path) -> AppConfig:
    auth = AuthConfig(
        tenant_id="00000000-0000-0000-0000-000000000000",
        client_id="00000000-0000-0000-0000-000000000001",
        client_secret="dummy-secret-for-tests",
    )
    return AppConfig(
        apps=[auth],
        throttle=ThrottleConfig(max_retries=0, request_timeout_seconds=5.0),
        migration=MigrationConfig(workers_per_mailbox=2, max_parallel_mailboxes=2),
        paths=PathsConfig(state_dir=tmp_path / "state", work_dir=tmp_path / "work"),
    )


def _resp_json(payload: dict) -> MagicMock:
    r = MagicMock()
    r.json.return_value = payload
    return r


@pytest.fixture
def orch_and_state(tmp_path: Path) -> tuple[Orchestrator, StateStore]:
    cfg = _minimal_cfg(tmp_path)
    state_path = tmp_path / "state.sqlite"
    # Avoid constructing MSAL clients (network + real tenant) in unit tests.
    pool = MagicMock()
    pool.names = ["testapp"]
    pool.pick.return_value = "testapp"
    token_provider = MagicMock()
    token_provider.get.return_value = "fake-token"
    pool.get.return_value = token_provider
    return Orchestrator(cfg, StateStore(state_path), pool), StateStore(state_path)


def test_purge_mail_enum_second_page_failure_counts_error(orch_and_state: tuple[Orchestrator, StateStore]) -> None:
    """If page 2 of mailFolders fails, we must not return success with a truncated tree."""
    orch, _state = orch_and_state
    row = MappingRow(pst_path=Path("/x.pst"), target_mailbox="user@contoso.com", target_root_folder="")

    page1 = {
        "value": [
            {
                "id": "folder-a",
                "displayName": "A",
                "totalItemCount": 0,
                "childFolderCount": 0,
            },
        ],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/user%40contoso.com/mailFolders?$skip=1",
    }

    def fake_get(path: str, **kwargs):
        if "mailFolders" in path and "childFolders" not in path and "messages" not in path:
            if "skip=1" in path or "$skip=1" in path:
                raise GraphError(503, {"error": "simulated"}, request_id="r1")
            return _resp_json(page1)
        if path.endswith("/mailFolders/deleteditems") or "deleteditems" in path:
            return _resp_json({"id": "del-items-id"})
        raise AssertionError(f"unexpected GET {path!r}")

    from jtet_pstmigrate.graph_client import GraphClient as GC

    with patch.object(GC, "get", side_effect=fake_get):
        deleted, missing, errors = orch.purge_mail([row])

    assert errors == 1
    assert deleted == 0
    assert missing == 0


def test_purge_mail_enum_follows_next_link(orch_and_state: tuple[Orchestrator, StateStore]) -> None:
    """Both pages of top-level mailFolders are processed (no silent truncation)."""
    orch, _state = orch_and_state
    row = MappingRow(pst_path=Path("/x.pst"), target_mailbox="u@t.com", target_root_folder="")

    page1 = {
        "value": [
            {
                "id": "custom-1",
                "displayName": "Custom",
                "totalItemCount": 1,
                "childFolderCount": 0,
            },
        ],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/u%40t.com/mailFolders?$skip=1",
    }
    page2 = {
        "value": [
            {
                "id": "custom-2",
                "displayName": "Custom2",
                "totalItemCount": 0,
                "childFolderCount": 0,
            },
        ],
    }

    def fake_get(path: str, **kwargs):
        if "deleteditems" in path:
            return _resp_json({"id": "del"})
        if "mailFolders" in path and "childFolders" not in path and "messages" not in path:
            if "skip=1" in path or "$skip=1" in path:
                return _resp_json(page2)
            return _resp_json(page1)
        raise AssertionError(f"unexpected GET {path!r}")

    delete_calls: list[str] = []

    def fake_delete(path: str, **kwargs):
        delete_calls.append(path)
        return MagicMock(status_code=204)

    from jtet_pstmigrate.graph_client import GraphClient as GC

    with patch.object(GC, "get", side_effect=fake_get), patch.object(GC, "delete", side_effect=fake_delete):
        deleted, missing, errors = orch.purge_mail([row])

    assert errors == 0
    assert deleted == 1  # only custom-1 had totalItemCount 1
    folder_delete_paths = [p for p in delete_calls if "/mailFolders/" in p and "/messages/" not in p]
    assert {p.split("/mailFolders/")[-1] for p in folder_delete_paths} == {"custom-1", "custom-2"}
    assert missing == 0
