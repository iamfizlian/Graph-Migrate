"""Regression tests for purge-mail Graph enumeration."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import pytest
from jtet_pstmigrate.config import AppConfig, AuthConfig, MigrationConfig, PathsConfig, ThrottleConfig
from jtet_pstmigrate.orchestrator import Orchestrator
from jtet_pstmigrate.state import StateStore


@pytest.fixture
def purge_mail_cfg(tmp_path: Path) -> AppConfig:
    return AppConfig(
        apps=[
            AuthConfig(
                tenant_id="00000000-0000-0000-0000-000000000001",
                client_id="00000000-0000-0000-0000-000000000002",
                client_secret="unused-in-tests",
            )
        ],
        paths=PathsConfig(state_dir=tmp_path / "st", work_dir=tmp_path / "wk"),
        migration=MigrationConfig(workers_per_mailbox=2, max_parallel_mailboxes=1),
        throttle=ThrottleConfig(),
    )


def test_purge_mail_mailfolders_enum_raises_on_second_page(
    tmp_path: Path, purge_mail_cfg: AppConfig
) -> None:
    """Mid-pagination failure must not return a partial folder list (silent incomplete purge)."""
    state = StateStore(tmp_path / "state.sqlite")
    # ``_purge_mail_one`` does not touch the pool; avoid constructing MSAL clients.
    pool = MagicMock()
    pool.names = ["test-app"]
    orch = Orchestrator(purge_mail_cfg, state, pool)

    del_resp = MagicMock()
    del_resp.json.return_value = {"id": "deleted-items-id"}
    page1 = MagicMock()
    page1.json.return_value = {
        "value": [
            {
                "id": "folder-a",
                "displayName": "Custom",
                "wellKnownName": None,
                "totalItemCount": 0,
                "childFolderCount": 0,
            }
        ],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/u/mailFolders?$skiptoken=x",
    }

    calls: dict[str, int] = {"n": 0}

    def get_side_effect(url: str, expect_status: tuple[int, ...] = (200,)) -> MagicMock:
        calls["n"] += 1
        if calls["n"] == 1:
            assert "deleteditems" in url
            return del_resp
        if calls["n"] == 2:
            assert "mailFolders" in url and "$skiptoken" not in url
            return page1
        raise RuntimeError("simulated page-2 failure")

    graph = MagicMock()
    graph.get.side_effect = get_side_effect

    with pytest.raises(RuntimeError, match="simulated page-2"):
        orch._purge_mail_one(graph, "user@contoso.com")


def test_deleted_items_skipped_by_wellknown_when_id_lookup_fails(
    tmp_path: Path, purge_mail_cfg: AppConfig
) -> None:
    """If resolving /deleteditems fails, wellKnownName on folder rows still skips that tree."""
    state = StateStore(tmp_path / "state.sqlite")
    pool = MagicMock()
    pool.names = ["test-app"]
    orch = Orchestrator(purge_mail_cfg, state, pool)

    page1 = MagicMock()
    page1.json.return_value = {
        "value": [
            {
                "id": "del-id",
                "displayName": "Deleted Items",
                "wellKnownName": "deleteditems",
                "totalItemCount": 3,
                "childFolderCount": 0,
            }
        ],
    }

    def get_side_effect(url: str, expect_status: tuple[int, ...] = (200,)) -> MagicMock:
        if "deleteditems" in url:
            raise ConnectionError("simulated lookup failure")
        if "mailFolders" in url and "childFolders" not in url:
            return page1
        raise AssertionError(f"unexpected GET {url!r}")

    graph = MagicMock()
    graph.get.side_effect = get_side_effect

    d, m, e = orch._purge_mail_one(graph, "user@contoso.com")
    assert (d, m, e) == (0, 0, 0)
    graph.delete.assert_not_called()
