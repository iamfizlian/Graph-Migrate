"""Safety regressions for ``purge-mail`` (Graph bulk delete)."""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from jtet_pstmigrate.graph_client import GraphError
from jtet_pstmigrate.orchestrator import Orchestrator


@pytest.fixture
def orch() -> Orchestrator:
    cfg = MagicMock()
    cfg.migration.workers_per_mailbox = 2
    return Orchestrator(cfg, MagicMock(), MagicMock())


def _json_response(data: dict) -> MagicMock:
    r = MagicMock()
    r.json.return_value = data
    return r


def test_purge_mail_aborts_if_deleted_items_folder_cannot_be_resolved(orch: Orchestrator) -> None:
    graph = MagicMock()
    graph.get.side_effect = GraphError(503, {"error": "Service unavailable"})

    d, m, e = orch._purge_mail_one(graph, "user@contoso.com")

    assert (d, m, e) == (0, 0, 1)
    graph.get.assert_called_once()
    graph.delete.assert_not_called()


def test_purge_mail_aborts_if_deleted_items_response_has_no_id(orch: Orchestrator) -> None:
    graph = MagicMock()
    graph.get.return_value = _json_response({})

    d, m, e = orch._purge_mail_one(graph, "user@contoso.com")

    assert (d, m, e) == (0, 0, 1)
    graph.delete.assert_not_called()


def test_purge_mail_propagates_mailfolder_enum_failure(orch: Orchestrator) -> None:
    """A failed /mailFolders page must not be misread as an empty mailbox."""
    graph = MagicMock()
    di = _json_response({"id": "deleted-items-id"})

    def get_side_effect(url: str, **kwargs: object) -> MagicMock:
        if "deleteditems" in url:
            return di
        if "/mailFolders?" in url or url.endswith("/mailFolders"):
            raise GraphError(500, {"error": "Internal error"})
        raise AssertionError(f"unexpected GET {url!r}")

    graph.get.side_effect = get_side_effect

    _d, _m, e = orch._purge_mail_one(graph, "user@contoso.com")

    assert e >= 1
    graph.delete.assert_not_called()


def test_purge_mail_empty_mailbox_after_deleted_items_ok(orch: Orchestrator) -> None:
    graph = MagicMock()
    di = _json_response({"id": "deleted-items-id"})
    empty_roots = _json_response({"value": []})

    def get_side_effect(url: str, **kwargs: object) -> MagicMock:
        if "deleteditems" in url:
            return di
        if "/mailFolders?" in url or url.endswith("/mailFolders"):
            return empty_roots
        raise AssertionError(f"unexpected GET {url!r}")

    graph.get.side_effect = get_side_effect

    d, m, e = orch._purge_mail_one(graph, "user@contoso.com")

    assert (d, m, e) == (0, 0, 0)
    graph.delete.assert_not_called()
