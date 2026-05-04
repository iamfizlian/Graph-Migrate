"""Regression tests for purge-mail Deleted Items handling."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import Mock

from jtet_pstmigrate.config import AppConfig, AuthConfig
from jtet_pstmigrate.graph_client import GraphError
from jtet_pstmigrate.orchestrator import Orchestrator
from jtet_pstmigrate.state import StateStore


def _test_auth() -> AuthConfig:
    return AuthConfig(
        tenant_id="aaa00000-0000-0000-0000-000000000000",
        client_id="00000000-0000-0000-0000-000000000001",
        client_secret="test-secret",
        name="testapp",
    )


def _test_cfg() -> AppConfig:
    return AppConfig(apps=[_test_auth()])


def test_purge_mail_aborts_when_deleted_items_id_unresolvable(
    tmp_path: Path,
) -> None:
    """Must not cascade-delete without knowing Deleted Items -- data loss."""
    graph = Mock()
    graph.get.side_effect = GraphError(503, "temporarily unavailable")

    cfg = _test_cfg()
    state = StateStore(tmp_path / "state.sqlite")
    orch = Orchestrator(cfg, state, Mock())

    d, m, e = orch._purge_mail_one(graph, "user@contoso.com")

    assert (d, m, e) == (0, 0, 1)
    graph.delete.assert_not_called()


def test_purge_mail_skips_resolved_deleted_items_folder(
    tmp_path: Path,
) -> None:
    del_id = "del-folder-id"
    inbox_id = "inbox-id"

    def ok(body):
        m = Mock()
        m.json.return_value = body
        return m

    responses = [
        ok({"id": del_id}),
        ok(
            {
                "value": [
                    {
                        "id": del_id,
                        "displayName": "Deleted Items",
                        "totalItemCount": 3,
                        "childFolderCount": 0,
                    },
                    {
                        "id": inbox_id,
                        "displayName": "Inbox",
                        "totalItemCount": 0,
                        "childFolderCount": 0,
                    },
                ]
            }
        ),
        ok({"value": []}),
    ]

    graph = Mock()
    graph.get.side_effect = responses

    def delete_guard(path: str, **kwargs):
        if del_id in path:
            raise AssertionError("Deleted Items folder must never be DELETE'd")
        if inbox_id in path:
            raise GraphError(400, "cannot delete distinguished folder")
        raise AssertionError(f"unexpected DELETE {path}")

    graph.delete.side_effect = delete_guard

    cfg = _test_cfg()
    state = StateStore(tmp_path / "s2.sqlite")
    orch = Orchestrator(cfg, state, Mock())

    _d, _m, e = orch._purge_mail_one(graph, "user@contoso.com")

    assert e == 0
    assert graph.delete.call_count == 1
