"""Regression: purge-mail must drain a folder when cascade DELETE fails with throttle.

If ``DELETE /mailFolders/{id}`` exhausts retries on 429/5xx, we still recurse into
children and DELETE messages. Returning early would leave mail in the mailbox
while the command appeared to finish.
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

from jtet_pstmigrate.config import AppConfig, AuthConfig, MigrationConfig, PathsConfig
from jtet_pstmigrate.graph_client import GraphError
from jtet_pstmigrate.orchestrator import Orchestrator
from jtet_pstmigrate.state import StateStore


def _resp_json(payload: dict) -> MagicMock:
    r = MagicMock()
    r.json.return_value = payload
    return r


class _ScriptedGraph:
    """Minimal stand-in for GraphClient used by ``_purge_mail_one``."""

    def __init__(self, gets: list, deletes: list):
        self._gets = list(gets)
        self._deletes = list(deletes)
        self.get_calls: list[str] = []
        self.delete_calls: list[str] = []

    def __enter__(self) -> _ScriptedGraph:
        return self

    def __exit__(self, *exc) -> None:
        pass

    def close(self) -> None:
        pass

    def get(self, path: str, **kwargs) -> MagicMock:
        self.get_calls.append(path)
        if not self._gets:
            raise AssertionError(f"unexpected GET {path!r}")
        item = self._gets.pop(0)
        if isinstance(item, Exception):
            raise item
        return item

    def delete(self, path: str, **kwargs) -> MagicMock:
        self.delete_calls.append(path)
        if not self._deletes:
            raise AssertionError(f"unexpected DELETE {path!r}")
        item = self._deletes.pop(0)
        if isinstance(item, Exception):
            raise item
        return _resp_json({})


def test_folder_delete_throttle_still_drains_messages(tmp_path: Path) -> None:
    mailbox = "user@example.com"
    upn_q = "user%40example.com"
    folder_id = "folder-1"
    msg_id = "msg-1"

    graph = _ScriptedGraph(
        gets=[
            _resp_json({"id": "deleted-items-id"}),
            _resp_json(
                {
                    "value": [
                        {
                            "id": folder_id,
                            "displayName": "Custom",
                            "totalItemCount": 1,
                            "childFolderCount": 0,
                        }
                    ]
                }
            ),
            # No GET /childFolders when childFolderCount is 0 — next read is message page.
            _resp_json({"value": [{"id": msg_id}]}),
        ],
        deletes=[
            GraphError(429, {"error": "throttled"}),
            _resp_json({}),
        ],
    )

    cfg = AppConfig(
        apps=[
            AuthConfig(
                tenant_id="00000000-0000-0000-0000-000000000001",
                client_id="00000000-0000-0000-0000-000000000002",
                client_secret="x",
            )
        ],
        migration=MigrationConfig(workers_per_mailbox=1, max_parallel_mailboxes=1),
        paths=PathsConfig(state_dir=tmp_path),
    )
    pool = MagicMock()
    pool.names = ["app0"]
    pool.pick.return_value = "app0"

    orch = Orchestrator(cfg, StateStore(tmp_path / "state.sqlite"), pool)
    deleted, missing, errors = orch._purge_mail_one(graph, mailbox)

    assert deleted == 1
    assert missing == 0
    assert errors == 0
    assert not graph._gets, f"unused GET scripts left: {graph._gets!r}"
    assert not graph._deletes, f"unused DELETE scripts left: {graph._deletes!r}"

    assert any(f"/users/{upn_q}/mailFolders/{folder_id}/messages" in c for c in graph.get_calls)
    assert any(f"/users/{upn_q}/messages/{msg_id}" in c for c in graph.delete_calls)


def test_folder_delete_404_skips_drain(tmp_path: Path) -> None:
    mailbox = "user@example.com"
    folder_id = "folder-gone"

    graph = _ScriptedGraph(
        gets=[
            _resp_json({"id": "deleted-items-id"}),
            _resp_json(
                {
                    "value": [
                        {
                            "id": folder_id,
                            "displayName": "Gone",
                            "totalItemCount": 5,
                            "childFolderCount": 0,
                        }
                    ]
                }
            ),
        ],
        deletes=[GraphError(404, {})],
    )

    cfg = AppConfig(
        apps=[
            AuthConfig(
                tenant_id="00000000-0000-0000-0000-000000000001",
                client_id="00000000-0000-0000-0000-000000000002",
                client_secret="x",
            )
        ],
        migration=MigrationConfig(workers_per_mailbox=1, max_parallel_mailboxes=1),
        paths=PathsConfig(state_dir=tmp_path),
    )
    pool = MagicMock()
    pool.names = ["app0"]
    pool.pick.return_value = "app0"

    orch = Orchestrator(cfg, StateStore(tmp_path / "state.sqlite"), pool)
    deleted, missing, errors = orch._purge_mail_one(graph, mailbox)

    assert deleted == 0
    assert missing == 0
    assert errors == 0
    assert not any("/messages" in c for c in graph.get_calls), "should not enumerate messages"
    assert not any("/messages/" in c for c in graph.delete_calls), (
        "404 on folder DELETE should not enqueue per-message DELETE"
    )
    assert not graph._gets
    assert not graph._deletes
