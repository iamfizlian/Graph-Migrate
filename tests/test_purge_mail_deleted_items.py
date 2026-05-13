"""Regression tests for purge-mail Deleted Items handling."""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from jtet_pstmigrate.graph_client import GraphError
from jtet_pstmigrate.orchestrator import _deleted_items_folder_id


def test_deleted_items_id_success() -> None:
    graph = MagicMock()
    graph.get.return_value.json.return_value = {"id": "folder-guid-1"}
    assert _deleted_items_folder_id(graph, "user%40contoso.com") == "folder-guid-1"
    graph.get.assert_called_once()


def test_deleted_items_id_missing_raises() -> None:
    graph = MagicMock()
    graph.get.return_value.json.return_value = {}
    with pytest.raises(ValueError, match="no id"):
        _deleted_items_folder_id(graph, "u")


def test_deleted_items_graph_error_propagates() -> None:
    graph = MagicMock()
    graph.get.side_effect = GraphError(503, "unavailable", None)
    with pytest.raises(GraphError):
        _deleted_items_folder_id(graph, "u")
