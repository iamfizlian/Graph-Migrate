"""Resolve and create destination mailFolders in a Graph mailbox.

Routing (Outlook folder structure, NOT 'Imported PST' parent):

1. Strip readpst's pseudo-roots ('Top of Personal Folders' etc.) from the
   incoming segment list -- those are MAPI tree placeholders and don't
   correspond to a user-visible Outlook folder.

2. If the first remaining segment matches a well-known Outlook folder
   ('Inbox', 'Sent Items', 'Drafts', 'Deleted Items', 'Junk Email',
   'Outbox', 'Archive', case-insensitive) we resolve it via Graph's
   well-known endpoint (``/mailFolders/{name}``) and create the rest of
   the path as children of that folder. This makes PST 'Inbox/Customers'
   land in the mailbox's real Inbox under 'Customers', not under a
   parallel 'Inbox' folder we created ourselves.

3. Anything else is created at the mailbox root, preserving the PST's
   subfolder hierarchy as-is.

Folder IDs are cached in-memory and in SQLite (``folder_map``). Cache keys
are scoped to avoid collisions between runs that used the old
'Imported PST'-prefix layout and runs using the new flat layout.
"""

from __future__ import annotations

import threading
from collections.abc import Sequence
from urllib.parse import quote

from loguru import logger

from jtet_pstmigrate.graph_client import GraphClient, GraphError
from jtet_pstmigrate.state import StateStore

# Top-of-store names readpst emits as the outermost folder segment. They
# don't represent a real Outlook folder, just the MAPI tree root, so we
# strip them before mapping anything else.
_PST_ROOT_NAMES: frozenset[str] = frozenset({
    "top of personal folders",
    "personal folders",
    "top of outlook data file",
    "outlook data file",
    # libpst sometimes uses these too:
    "top of information store",
    "imap",
})

# Display name (case-insensitive, trimmed) -> Graph well-known folder name
# identifier. Setting an entry here means messages whose PST first segment
# matches will land in the mailbox's real folder of that kind.
# https://learn.microsoft.com/en-us/graph/api/resources/mailfolder
_WELL_KNOWN_FOLDERS: dict[str, str] = {
    "inbox":          "inbox",
    "sent items":     "sentitems",
    "sent":           "sentitems",
    "drafts":         "drafts",
    "deleted items":  "deleteditems",
    "junk email":     "junkemail",
    "junk e-mail":    "junkemail",
    "junk":           "junkemail",
    "outbox":         "outbox",
    "archive":        "archive",
}


class FolderManager:
    """One per (mailbox, run). Internally caches resolved folder IDs."""

    def __init__(self, graph: GraphClient, state: StateStore, mailbox: str):
        self._graph = graph
        self._state = state
        self._mailbox = mailbox
        self._mem_cache: dict[str, str] = {}
        self._lock = threading.Lock()
        self._log = logger.bind(ctx=f"folders[{mailbox}]")

    def ensure_path(self, segments: Sequence[str]) -> str:
        """Return the Graph folder ID for the deepest segment.

        Resolves PST segments to the mailbox's real Outlook folder
        structure (Inbox, Sent Items, etc.) where possible. Anything that
        doesn't map is created at the mailbox root.
        """
        cleaned = self._strip_pst_root(segments)
        if not cleaned:
            # PST had only a pseudo-root with messages directly under it.
            # Stash these in Inbox; they're loose top-level mail.
            return self._resolve_well_known("inbox")

        first_norm = cleaned[0].strip().lower()
        well_known_id = _WELL_KNOWN_FOLDERS.get(first_norm)
        if well_known_id is not None:
            anchor_id = self._resolve_well_known(well_known_id)
            rest = cleaned[1:]
            if not rest:
                return anchor_id
            cache_prefix = f"@{well_known_id}"
            return self._walk_children(
                parent_id=anchor_id, segments=rest, cache_prefix=cache_prefix,
            )

        return self._walk_children(
            parent_id=None, segments=cleaned, cache_prefix="",
        )

    @staticmethod
    def _strip_pst_root(segments: Sequence[str]) -> list[str]:
        """Drop a single leading pseudo-root segment if present.

        Only the *first* segment is checked. Subfolders that happen to
        share the same name keep their position (e.g. a user folder
        literally called 'Personal Folders' under Inbox is preserved).
        """
        if not segments:
            return []
        if segments[0].strip().lower() in _PST_ROOT_NAMES:
            return list(segments[1:])
        return list(segments)

    def _resolve_well_known(self, well_known_id: str) -> str:
        """Cached lookup of a Graph well-known mailFolder."""
        cache_key = f"@{well_known_id}"
        with self._lock:
            cached = self._mem_cache.get(cache_key)
            if cached:
                return cached
        actual_id = self._get_well_known(well_known_id)
        with self._lock:
            self._mem_cache[cache_key] = actual_id
        return actual_id

    def _walk_children(
        self,
        *,
        parent_id: str | None,
        segments: Sequence[str],
        cache_prefix: str,
    ) -> str:
        """Walk segments below ``parent_id``, creating folders as needed.

        ``cache_prefix`` makes cache keys unique across roots, so
        ``Inbox/Customers/2019`` and ``Customers/2019`` (at the mailbox
        root) don't collide.
        """
        if not segments:
            if parent_id is None:
                # Shouldn't happen at this call-site, but be defensive.
                return self._resolve_well_known("inbox")
            return parent_id

        full_key = f"{cache_prefix}/{'/'.join(segments)}" if cache_prefix else "/".join(segments)
        with self._lock:
            cached = (
                self._mem_cache.get(full_key)
                or self._state.get_folder_id(self._mailbox, full_key)
            )
            if cached:
                self._mem_cache[full_key] = cached
                return cached

        accumulated: list[str] = []
        current_parent = parent_id
        for segment in segments:
            accumulated.append(segment)
            joined = "/".join(accumulated)
            partial_key = f"{cache_prefix}/{joined}" if cache_prefix else joined
            with self._lock:
                level_id = (
                    self._mem_cache.get(partial_key)
                    or self._state.get_folder_id(self._mailbox, partial_key)
                )
            if level_id:
                current_parent = level_id
                continue
            level_id = self._create_or_find(current_parent, segment)
            with self._lock:
                self._mem_cache[partial_key] = level_id
                self._state.put_folder_id(self._mailbox, partial_key, level_id)
            current_parent = level_id

        return current_parent  # type: ignore[return-value]

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
