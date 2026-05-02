"""Safety checks for purge-mail (must never drain Deleted Items blindly)."""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from jtet_pstmigrate.config import AppConfig, AuthConfig, PathsConfig
from jtet_pstmigrate.graph_client import GraphError
from jtet_pstmigrate.orchestrator import Orchestrator


def _minimal_cfg() -> AppConfig:
    # Valid auth shape only; tests never construct AppPool / MSAL.
    return AppConfig(
        apps=[
            AuthConfig(
                name="testapp",
                tenant_id="00000000-0000-0000-0000-000000000001",
                client_id="00000000-0000-0000-0000-000000000002",
                client_secret="unused-in-unit-test",
            )
        ],
        paths=PathsConfig(),
    )


def test_purge_mail_aborts_if_deleteditems_folder_cannot_be_resolved() -> None:
    """If we cannot GET deleteditems, we must not fall through and drain that folder."""
    cfg = _minimal_cfg()
    orch = Orchestrator(cfg, MagicMock(), MagicMock())

    graph = MagicMock()
    graph.get.side_effect = GraphError(503, {"error": "Graph unavailable"})

    with pytest.raises(RuntimeError, match="Deleted Items"):
        orch._purge_mail_one(graph, "user@example.com")

    graph.get.assert_called_once()
    graph.delete.assert_not_called()
