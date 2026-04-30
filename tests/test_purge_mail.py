"""Tests for mail purge safety (Deleted Items must never be drained)."""

from __future__ import annotations

from unittest.mock import MagicMock

from jtet_pstmigrate.config import AppConfig, MigrationConfig, ThrottleConfig
from jtet_pstmigrate.orchestrator import Orchestrator


def _make_cfg() -> AppConfig:
    # ``AppConfig()`` requires validated apps/MSAL; model_construct avoids that for unit tests.
    return AppConfig.model_construct(
        migration=MigrationConfig(),
        throttle=ThrottleConfig(),
    )


def _orch() -> Orchestrator:
    pool = MagicMock()
    pool.names = ["test-app"]
    return Orchestrator(_make_cfg(), MagicMock(), pool)


def _resp_json(data: dict) -> MagicMock:
    m = MagicMock()
    m.json.return_value = data
    return m


def test_purge_mail_aborts_when_deleted_items_folder_unresolvable() -> None:
    """If we cannot GET deleteditems, draining could wipe that folder — abort."""
    graph = MagicMock()
    graph.get.side_effect = ConnectionError("simulated Graph failure")
    orch = _orch()
    assert orch._purge_mail_one(graph, "user@contoso.com") == (0, 0, 1)


def test_purge_mail_skips_deleted_items_by_wellknown_when_id_mismatch() -> None:
    """Defense in depth: skip folder when wellKnownName is deleteditems."""
    graph = MagicMock()
    graph.get.side_effect = [
        _resp_json({"id": "id-from-wellknown-endpoint"}),
        _resp_json(
            {
                "value": [
                    {
                        "id": "different-id-from-enum",
                        "displayName": "Deleted Items",
                        "wellKnownName": "deleteditems",
                        "totalItemCount": 50,
                        "childFolderCount": 0,
                    }
                ]
            }
        ),
    ]
    orch = _orch()
    d, m, e = orch._purge_mail_one(graph, "user@contoso.com")
    assert (d, m, e) == (0, 0, 0)
    graph.delete.assert_not_called()
