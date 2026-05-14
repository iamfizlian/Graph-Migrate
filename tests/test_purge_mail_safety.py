"""Safety checks for purge-mail: never drain Deleted Items without a resolved id,
and never treat Graph enumeration failures as an empty mailbox."""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from jtet_pstmigrate.config import AppConfig, MigrationConfig, PathsConfig, ThrottleConfig
from jtet_pstmigrate.graph_client import GraphError
from jtet_pstmigrate.orchestrator import Orchestrator


def _orch() -> Orchestrator:
    cfg = AppConfig.model_construct(
        migration=MigrationConfig(workers_per_mailbox=2),
        throttle=ThrottleConfig(),
        paths=PathsConfig(),
    )
    return Orchestrator(cfg, MagicMock(), MagicMock())


def test_purge_mail_aborts_when_deleted_items_lookup_fails() -> None:
    graph = MagicMock()
    graph.get.side_effect = GraphError(503, "service unavailable")

    d, m, e = _orch()._purge_mail_one(graph, "user@example.com")

    assert (d, m, e) == (0, 0, 1)
    graph.get.assert_called_once()
    graph.delete.assert_not_called()


def test_purge_mail_aborts_when_deleted_items_response_has_no_id() -> None:
    graph = MagicMock()
    ok = MagicMock()
    ok.json.return_value = {}
    graph.get.return_value = ok

    d, m, e = _orch()._purge_mail_one(graph, "user@example.com")

    assert (d, m, e) == (0, 0, 1)
    graph.delete.assert_not_called()


def test_purge_mail_propagates_graph_error_during_mailfolder_enum() -> None:
    graph = MagicMock()
    del_resp = MagicMock()
    del_resp.json.return_value = {"id": "deleted-items-folder-id"}

    def get_side_effect(url: str, **_kwargs):
        if "deleteditems" in url:
            return del_resp
        raise GraphError(500, "internal error")

    graph.get.side_effect = get_side_effect

    with pytest.raises(GraphError):
        _orch()._purge_mail_one(graph, "user@example.com")

    graph.delete.assert_not_called()
