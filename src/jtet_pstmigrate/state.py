"""SQLite-backed state store for resume + dedupe.

Schema is intentionally tiny: one row per source message, one row per folder
mapping. WAL mode + per-connection short transactions keep concurrent workers
from blocking each other.

Why SQLite over flat files?
  - Atomic upserts across many concurrent workers (the prior PowerShell script's
    StreamWriter approach was fine for one writer; not for parallel mailboxes).
  - Trivial to query post-run: 'how many failed?', 'list failures by folder', etc.
  - Single-file artefact, easy to ship as part of the run output.
"""

from __future__ import annotations

import sqlite3
import threading
import time
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

SCHEMA = """
PRAGMA journal_mode = WAL;
PRAGMA synchronous = NORMAL;
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS pst_runs (
    pst_path        TEXT NOT NULL,
    target_mailbox  TEXT NOT NULL,
    status          TEXT NOT NULL,        -- pending|extracting|uploading|done|failed
    started_at      REAL,
    finished_at     REAL,
    items_total     INTEGER DEFAULT 0,
    items_uploaded  INTEGER DEFAULT 0,
    items_skipped   INTEGER DEFAULT 0,
    items_failed    INTEGER DEFAULT 0,
    last_error      TEXT,
    PRIMARY KEY (pst_path, target_mailbox)
);

CREATE TABLE IF NOT EXISTS folder_map (
    target_mailbox    TEXT NOT NULL,
    folder_path       TEXT NOT NULL,        -- e.g. 'Imported PST/Inbox/Subfolder'
    graph_folder_id   TEXT NOT NULL,
    PRIMARY KEY (target_mailbox, folder_path)
);

CREATE TABLE IF NOT EXISTS messages (
    target_mailbox     TEXT NOT NULL,
    pst_path           TEXT NOT NULL,
    source_path        TEXT NOT NULL,        -- absolute path of MIME file
    dedupe_key         TEXT NOT NULL,        -- internet message-id or content hash
    graph_message_id   TEXT,
    graph_folder_id    TEXT,
    app_id             TEXT,                 -- which app pool member did the upload
    status             TEXT NOT NULL,        -- queued|done|failed|skipped
    bytes              INTEGER DEFAULT 0,
    last_error         TEXT,
    updated_at         REAL,
    PRIMARY KEY (target_mailbox, pst_path, source_path)
);

CREATE INDEX IF NOT EXISTS idx_messages_dedupe
    ON messages(target_mailbox, dedupe_key);
CREATE INDEX IF NOT EXISTS idx_messages_status
    ON messages(target_mailbox, status);
CREATE INDEX IF NOT EXISTS idx_messages_app
    ON messages(app_id);

-- Non-mail PST items: appointments/events, contacts, tasks, notes, journal.
-- Kept in a dedicated table so the mail dedup keys (RFC822 Message-ID) and
-- the calendar/contact dedup keys (UID, vCard FN+EMAIL) cannot collide.
CREATE TABLE IF NOT EXISTS non_mail_items (
    target_mailbox    TEXT NOT NULL,
    pst_path          TEXT NOT NULL,
    source_path       TEXT NOT NULL,
    item_type         TEXT NOT NULL,        -- 'event'|'contact'|'task'|'note'
    dedupe_key        TEXT NOT NULL,
    graph_id          TEXT,                 -- /events/{id}, /contacts/{id}, ...
    app_id            TEXT,
    status            TEXT NOT NULL,        -- queued|done|failed|skipped
    bytes             INTEGER DEFAULT 0,
    last_error        TEXT,
    updated_at        REAL,
    PRIMARY KEY (target_mailbox, pst_path, source_path, item_type)
);

CREATE INDEX IF NOT EXISTS idx_non_mail_items_dedupe
    ON non_mail_items(target_mailbox, item_type, dedupe_key);
CREATE INDEX IF NOT EXISTS idx_non_mail_items_status
    ON non_mail_items(target_mailbox, item_type, status);
"""

# Forward-compatible additions for existing DBs created by an older version.
MIGRATIONS = [
    ("messages.app_id", "ALTER TABLE messages ADD COLUMN app_id TEXT"),
]


class StateStore:
    """Thread-safe wrapper. One instance per run, shared across workers."""

    def __init__(self, db_path: Path):
        db_path.parent.mkdir(parents=True, exist_ok=True)
        self._path = db_path
        self._local = threading.local()
        with self._connect() as conn:
            conn.executescript(SCHEMA)
            self._apply_migrations(conn)

    def _apply_migrations(self, conn: sqlite3.Connection) -> None:
        import contextlib

        existing_cols = {row["name"] for row in conn.execute("PRAGMA table_info(messages)")}
        for marker, sql in MIGRATIONS:
            _, col = marker.split(".", 1)
            if col not in existing_cols:
                # Concurrent column add is safe to ignore (another process raced us)
                with contextlib.suppress(sqlite3.OperationalError):
                    conn.execute(sql)

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        conn: sqlite3.Connection | None = getattr(self._local, "conn", None)
        if conn is None:
            conn = sqlite3.connect(
                self._path,
                timeout=30.0,
                isolation_level=None,  # autocommit; we manage transactions explicitly
                check_same_thread=False,
            )
            conn.row_factory = sqlite3.Row
            self._local.conn = conn
        yield conn

    # PST run lifecycle --------------------------------------------------

    def start_pst_run(self, pst_path: str, target_mailbox: str) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO pst_runs(pst_path, target_mailbox, status, started_at)
                VALUES (?, ?, 'extracting', ?)
                ON CONFLICT(pst_path, target_mailbox) DO UPDATE SET
                    status='extracting',
                    started_at=excluded.started_at,
                    last_error=NULL
                """,
                (pst_path, target_mailbox, time.time()),
            )

    def update_pst_run(
        self,
        pst_path: str,
        target_mailbox: str,
        *,
        status: str | None = None,
        items_total: int | None = None,
        last_error: str | None = None,
    ) -> None:
        sets, vals = [], []
        if status is not None:
            sets.append("status=?")
            vals.append(status)
            if status in ("done", "failed"):
                sets.append("finished_at=?")
                vals.append(time.time())
        if items_total is not None:
            sets.append("items_total=?")
            vals.append(items_total)
        if last_error is not None:
            sets.append("last_error=?")
            vals.append(last_error)
        if not sets:
            return
        vals.extend([pst_path, target_mailbox])
        with self._connect() as conn:
            conn.execute(
                f"UPDATE pst_runs SET {', '.join(sets)} WHERE pst_path=? AND target_mailbox=?",
                vals,
            )

    def get_pst_run(self, pst_path: str, target_mailbox: str) -> sqlite3.Row | None:
        with self._connect() as conn:
            return conn.execute(
                "SELECT * FROM pst_runs WHERE pst_path=? AND target_mailbox=?",
                (pst_path, target_mailbox),
            ).fetchone()

    # Folder cache -------------------------------------------------------

    def get_folder_id(self, mailbox: str, folder_path: str) -> str | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT graph_folder_id FROM folder_map WHERE target_mailbox=? AND folder_path=?",
                (mailbox, folder_path),
            ).fetchone()
            return row["graph_folder_id"] if row else None

    def put_folder_id(self, mailbox: str, folder_path: str, graph_id: str) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO folder_map(target_mailbox, folder_path, graph_folder_id)
                VALUES (?, ?, ?)
                ON CONFLICT(target_mailbox, folder_path) DO UPDATE SET
                    graph_folder_id = excluded.graph_folder_id
                """,
                (mailbox, folder_path, graph_id),
            )

    # Message lifecycle --------------------------------------------------

    def is_message_done(self, mailbox: str, dedupe_key: str) -> bool:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT 1 FROM messages WHERE target_mailbox=? AND dedupe_key=? AND status='done' LIMIT 1",
                (mailbox, dedupe_key),
            ).fetchone()
            return row is not None

    def is_row_done(self, mailbox: str, pst_path: str, source_path: str) -> bool:
        """Whether the *exact* (mailbox, pst, source) triple is already done.

        Used to avoid downgrading a 'done' row to 'skipped' when a re-run
        encounters the same .eml file. Without this check, the dedupe hit on
        the message's own done row caused the upsert to overwrite status to
        'skipped', which broke idempotency on the next run.
        """
        with self._connect() as conn:
            row = conn.execute(
                "SELECT 1 FROM messages "
                "WHERE target_mailbox=? AND pst_path=? AND source_path=? AND status='done' "
                "LIMIT 1",
                (mailbox, pst_path, source_path),
            ).fetchone()
            return row is not None

    def upsert_message(
        self,
        *,
        mailbox: str,
        pst_path: str,
        source_path: str,
        dedupe_key: str,
        status: str,
        graph_message_id: str | None = None,
        graph_folder_id: str | None = None,
        app_id: str | None = None,
        bytes_: int = 0,
        last_error: str | None = None,
    ) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO messages(
                    target_mailbox, pst_path, source_path, dedupe_key,
                    graph_message_id, graph_folder_id, app_id,
                    status, bytes, last_error, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(target_mailbox, pst_path, source_path) DO UPDATE SET
                    dedupe_key       = excluded.dedupe_key,
                    graph_message_id = excluded.graph_message_id,
                    graph_folder_id  = excluded.graph_folder_id,
                    app_id           = excluded.app_id,
                    status           = excluded.status,
                    bytes            = excluded.bytes,
                    last_error       = excluded.last_error,
                    updated_at       = excluded.updated_at
                """,
                (
                    mailbox, pst_path, source_path, dedupe_key,
                    graph_message_id, graph_folder_id, app_id,
                    status, bytes_, last_error,
                    time.time(),
                ),
            )

    def app_breakdown(self) -> dict[str, dict[str, int]]:
        """Return {app_id: {status: count}} across all messages."""
        out: dict[str, dict[str, int]] = {}
        with self._connect() as conn:
            for row in conn.execute(
                """
                SELECT COALESCE(app_id, '(unset)') AS app, status, COUNT(*) AS c
                FROM messages
                GROUP BY app, status
                """
            ):
                out.setdefault(row["app"], {})[row["status"]] = row["c"]
        return out

    def counts_for_run(self, mailbox: str, pst_path: str) -> dict[str, int]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT status, COUNT(*) AS c
                FROM messages
                WHERE target_mailbox=? AND pst_path=?
                GROUP BY status
                """,
                (mailbox, pst_path),
            ).fetchall()
            out = {r["status"]: r["c"] for r in rows}
            return {
                "queued": out.get("queued", 0),
                "done": out.get("done", 0),
                "failed": out.get("failed", 0),
                "skipped": out.get("skipped", 0),
            }

    def all_runs(self) -> list[sqlite3.Row]:
        with self._connect() as conn:
            return list(conn.execute("SELECT * FROM pst_runs ORDER BY started_at"))

    # Non-mail item lifecycle (calendar/contacts/tasks/notes) -----------

    def is_item_done(self, mailbox: str, item_type: str, dedupe_key: str) -> bool:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT 1 FROM non_mail_items "
                "WHERE target_mailbox=? AND item_type=? AND dedupe_key=? AND status='done' "
                "LIMIT 1",
                (mailbox, item_type, dedupe_key),
            ).fetchone()
            return row is not None

    def is_item_row_done(
        self, mailbox: str, pst_path: str, source_path: str, item_type: str
    ) -> bool:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT 1 FROM non_mail_items "
                "WHERE target_mailbox=? AND pst_path=? AND source_path=? "
                "AND item_type=? AND status='done' LIMIT 1",
                (mailbox, pst_path, source_path, item_type),
            ).fetchone()
            return row is not None

    def upsert_item(
        self,
        *,
        mailbox: str,
        pst_path: str,
        source_path: str,
        item_type: str,
        dedupe_key: str,
        status: str,
        graph_id: str | None = None,
        app_id: str | None = None,
        bytes_: int = 0,
        last_error: str | None = None,
    ) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO non_mail_items(
                    target_mailbox, pst_path, source_path, item_type, dedupe_key,
                    graph_id, app_id, status, bytes, last_error, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(target_mailbox, pst_path, source_path, item_type) DO UPDATE SET
                    dedupe_key  = excluded.dedupe_key,
                    graph_id    = excluded.graph_id,
                    app_id      = excluded.app_id,
                    status      = excluded.status,
                    bytes       = excluded.bytes,
                    last_error  = excluded.last_error,
                    updated_at  = excluded.updated_at
                """,
                (
                    mailbox, pst_path, source_path, item_type, dedupe_key,
                    graph_id, app_id, status, bytes_, last_error,
                    time.time(),
                ),
            )

    def list_done_items(
        self, mailbox: str, pst_path: str, item_type: str
    ) -> list[sqlite3.Row]:
        """Items uploaded successfully, with the Graph id needed for deletion."""
        with self._connect() as conn:
            return list(conn.execute(
                """
                SELECT source_path, graph_id, app_id
                FROM non_mail_items
                WHERE target_mailbox=? AND pst_path=? AND item_type=?
                  AND status='done' AND graph_id IS NOT NULL
                """,
                (mailbox, pst_path, item_type),
            ))

    def delete_item(
        self, mailbox: str, pst_path: str, source_path: str, item_type: str
    ) -> None:
        with self._connect() as conn:
            conn.execute(
                "DELETE FROM non_mail_items "
                "WHERE target_mailbox=? AND pst_path=? AND source_path=? AND item_type=?",
                (mailbox, pst_path, source_path, item_type),
            )

    def counts_for_items(
        self, mailbox: str, pst_path: str, item_type: str
    ) -> dict[str, int]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT status, COUNT(*) AS c
                FROM non_mail_items
                WHERE target_mailbox=? AND pst_path=? AND item_type=?
                GROUP BY status
                """,
                (mailbox, pst_path, item_type),
            ).fetchall()
            out = {r["status"]: r["c"] for r in rows}
            return {
                "queued":  out.get("queued", 0),
                "done":    out.get("done", 0),
                "failed":  out.get("failed", 0),
                "skipped": out.get("skipped", 0),
            }

    # Bulk reset --------------------------------------------------------
    #
    # Used by ``pstmigrate reset-state`` to wipe local dedup/run rows
    # for a given (mailbox, pst) scope. Caller is responsible for any
    # destination-side cleanup -- this only touches our SQLite.

    def count_scope(self, mailbox: str, pst_path: str) -> dict[str, int]:
        """Return row counts in each scoped table for the given mailbox+PST.

        Used to show the user what they're about to wipe before
        confirming. Folder-map count is per-mailbox only since that table
        isn't keyed by PST.
        """
        with self._connect() as conn:
            mail = conn.execute(
                "SELECT COUNT(*) FROM messages WHERE target_mailbox=? AND pst_path=?",
                (mailbox, pst_path),
            ).fetchone()[0]
            non_mail = conn.execute(
                "SELECT COUNT(*) FROM non_mail_items "
                "WHERE target_mailbox=? AND pst_path=?",
                (mailbox, pst_path),
            ).fetchone()[0]
            runs = conn.execute(
                "SELECT COUNT(*) FROM pst_runs WHERE target_mailbox=? AND pst_path=?",
                (mailbox, pst_path),
            ).fetchone()[0]
        return {"messages": mail, "non_mail_items": non_mail, "pst_runs": runs}

    def count_folder_map(self, mailbox: str) -> int:
        with self._connect() as conn:
            return conn.execute(
                "SELECT COUNT(*) FROM folder_map WHERE target_mailbox=?",
                (mailbox,),
            ).fetchone()[0]

    def clear_scope(
        self, mailbox: str, pst_path: str
    ) -> dict[str, int]:
        """Delete messages/non_mail_items/pst_runs rows for the scope.

        Returns the number of rows deleted from each table so the caller
        can report it. Does NOT touch folder_map (that table isn't
        keyed by PST; clear it separately via :meth:`clear_folder_map`).
        """
        with self._connect() as conn:
            mail = conn.execute(
                "DELETE FROM messages WHERE target_mailbox=? AND pst_path=?",
                (mailbox, pst_path),
            ).rowcount
            non_mail = conn.execute(
                "DELETE FROM non_mail_items "
                "WHERE target_mailbox=? AND pst_path=?",
                (mailbox, pst_path),
            ).rowcount
            runs = conn.execute(
                "DELETE FROM pst_runs WHERE target_mailbox=? AND pst_path=?",
                (mailbox, pst_path),
            ).rowcount
        return {"messages": mail, "non_mail_items": non_mail, "pst_runs": runs}

    def clear_folder_map(self, mailbox: str) -> int:
        """Delete cached folder-id rows for a mailbox.

        Safe to clear without touching Graph: the folder manager always
        re-resolves cache misses via ``_create_or_find``, which queries
        Graph for an existing folder by name before creating a new one,
        so the worst case is a few extra GETs on the next import.
        """
        with self._connect() as conn:
            return conn.execute(
                "DELETE FROM folder_map WHERE target_mailbox=?",
                (mailbox,),
            ).rowcount
