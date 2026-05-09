"""Tests for Graph collection paging used by purge-mail."""

from unittest.mock import MagicMock

import pytest

from jtet_pstmigrate.graph_client import GRAPH_BASE
from jtet_pstmigrate.orchestrator import _graph_collect_paged


def test_graph_collect_paged_follows_next_link() -> None:
    g = MagicMock()
    first = MagicMock()
    first.json.return_value = {
        "value": [{"id": "a"}],
        "@odata.nextLink": f"{GRAPH_BASE}/users/x/mailFolders?$skip=1",
    }
    second = MagicMock()
    second.json.return_value = {"value": [{"id": "b"}]}
    g.get.side_effect = [first, second]

    rows = _graph_collect_paged(g, "/users/x/mailFolders")

    assert [r["id"] for r in rows] == ["a", "b"]
    assert g.get.call_count == 2


def test_graph_collect_paged_propagates_errors() -> None:
    g = MagicMock()
    g.get.side_effect = OSError("network down")

    with pytest.raises(OSError, match="network down"):
        _graph_collect_paged(g, "/users/x/mailFolders")
