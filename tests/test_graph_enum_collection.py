"""Tests for Graph collection paging used by purge-mail."""

from unittest.mock import MagicMock

from jtet_pstmigrate.graph_client import GraphError
from jtet_pstmigrate.orchestrator import graph_enum_collection


class _JsonResp:
    __slots__ = ("_data",)

    def __init__(self, data: dict) -> None:
        self._data = data

    def json(self) -> dict:
        return self._data


def test_graph_enum_collection_single_page_ok() -> None:
    graph = MagicMock()
    graph.get.return_value = _JsonResp({"value": [{"id": "a"}, {"id": "b"}]})
    rows, ok = graph_enum_collection(graph, "/users/x/mailFolders?$top=100")
    assert ok is True
    assert [r["id"] for r in rows] == ["a", "b"]
    graph.get.assert_called_once()


def test_graph_enum_collection_follows_next_link() -> None:
    graph = MagicMock()
    graph.get.side_effect = [
        _JsonResp(
            {
                "value": [{"id": "1"}],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/x/m?$skip=1",
            }
        ),
        _JsonResp({"value": [{"id": "2"}]}),
    ]
    rows, ok = graph_enum_collection(graph, "/users/x/m")
    assert ok is True
    assert [r["id"] for r in rows] == ["1", "2"]
    assert graph.get.call_count == 2


def test_graph_enum_collection_first_page_failure_not_ok() -> None:
    graph = MagicMock()
    graph.get.side_effect = GraphError(503, "busy")
    rows, ok = graph_enum_collection(graph, "/users/x/mailFolders")
    assert ok is False
    assert rows == []


def test_graph_enum_collection_mid_sequence_failure_partial_not_ok() -> None:
    graph = MagicMock()
    graph.get.side_effect = [
        _JsonResp(
            {
                "value": [{"id": "1"}],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/next",
            }
        ),
        GraphError(429, "throttled"),
    ]
    rows, ok = graph_enum_collection(graph, "/start")
    assert ok is False
    assert [r["id"] for r in rows] == ["1"]
