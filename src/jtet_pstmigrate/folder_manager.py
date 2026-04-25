"""Resolve and create destination mailFolders in a Graph mailbox.

Folder paths are joined with '/' for cache keys (consistent across runs),
but Graph API uses parent/child IDs, so we walk the tree one segment at a
time, creating missing folders and caching IDs in SQLite.
"""

from __future__ import annotations

import threading
from collections.abc import Sequence
from urllib.parse import quote

from loguru import logger

from jtet_pstmigrate.graph_client import GraphClient, GraphError
from jtet_pstmigrate.state import StateStore


class FolderManager:
    """One per (mailbox, run). Internally caches resolved folder IDs."""

    def __init__(self, graph: GraphClient, state: StateStore, mailbox: str):
        self._graph = graph
        self._state = state
        self._mailbox = mailbox
        self._mem_cache: dict[str, str] = {}
        self._lock = threading.Lock()
        self._log = logger.bind(ctx=f"folders[{mailbox}]")

    def ensure_path(self, segments: Sequence[str], *, root_folder: str | None = None) -> str:
        """Return the Graph folder ID for the deepest segment.

        Optional `root_folder` is created at the mailbox top level and prepended.
        """
        full_path: list[str] = []
        if root_folder:
            full_path.append(root_folder)
        full_path.extend(segments)

        if not full_path:
            return self._get_well_known("inbox")  # fall back, shouldn't happen

        cache_key = "/".join(full_path)
        with self._lock:
            cached = self._mem_cache.get(cache_key) or self._state.get_folder_id(self._mailbox, cache_key)
            if cached:
                self._mem_cache[cache_key] = cached
                return cached

        # Walk and create level by level
        parent_id: str | None = None
        accumulated: list[str] = []
        for segment in full_path:
            accumulated.append(segment)
            partial_key = "/".join(accumulated)
            with self._lock:
                level_id = self._mem_cache.get(partial_key) or self._state.get_folder_id(self._mailbox, partial_key)
            if level_id:
                parent_id = level_id
                continue
            level_id = self._create_or_find(parent_id, segment)
            with self._lock:
                self._mem_cache[partial_key] = level_id
                self._state.put_folder_id(self._mailbox, partial_key, level_id)
            parent_id = level_id

        return parent_id  # type: ignore[return-value]

    def _create_or_find(self, parent_id: str | None, name: str) -> str:
        """Create folder under parent (or at root); reuse if it already exists.

        Two workers can race on the same target folder: both find nothing,
        both POST, second one gets 409 ErrorFolderExists. On 409 we re-query
        and return the existing folder's id rather than letting the failure
        propagate up and mark every queued message as failed.
        """
        if parent_id:
            list_path = f"/users/{quote(self._mailbox)}/mailFolders/{parent_id}/childFolders"
        else:
            list_path = f"/users/{quote(self._mailbox)}/mailFolders"

        existing = self._find_child_by_name(list_path, name)
        if existing:
            self._log.debug("Reusing folder '{}' = {}", name, existing)
            return existing

        try:
            resp = self._graph.post(list_path, json={"displayName": name})
        except GraphError as e:
            if e.status == 409:
                refound = self._find_child_by_name(list_path, name)
                if refound:
                    self._log.debug("Folder '{}' created concurrently; reusing {}", name, refound)
                    return refound
            raise
        new_id = resp.json()["id"]
        self._log.info("Created folder '{}' = {}", name, new_id)
        return new_id

    def _find_child_by_name(self, list_path: str, name: str) -> str | None:
        """Search for a folder by displayName (case-insensitive)."""
        params = {
            "$filter": f"displayName eq '{name.replace(chr(39), chr(39) * 2)}'",
            "$select": "id,displayName",
            "$top": "10",
        }
        resp = self._graph.get(list_path, params=params)
        for item in resp.json().get("value", []):
            if item.get("displayName", "").lower() == name.lower():
                return item["id"]
        return None

    def _get_well_known(self, name: str) -> str:
        resp = self._graph.get(f"/users/{quote(self._mailbox)}/mailFolders/{name}", params={"$select": "id"})
        return resp.json()["id"]
