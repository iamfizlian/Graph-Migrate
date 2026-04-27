"""Top-level orchestration: extract, queue, upload, summarize.

Concurrency model:
  - We use a ThreadPoolExecutor sized by max_parallel_mailboxes for whole-PST
    workers; each worker processes one (PST, mailbox) row.
  - Within a single mailbox/PST job we use another small pool sized by
    workers_per_mailbox to upload messages in parallel. Graph throttles per
    mailbox so going wider than ~4 here yields diminishing returns.
  - SQLite WAL mode + per-thread connections keeps state writes lock-free.
"""

from __future__ import annotations

import csv
import dataclasses
import threading
import time
from collections.abc import Iterable
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

from loguru import logger
from rich.console import Console
from rich.progress import (
    BarColumn,
    MofNCompleteColumn,
    Progress,
    SpinnerColumn,
    TextColumn,
    TimeElapsedColumn,
    TimeRemainingColumn,
)
from rich.table import Table

from jtet_pstmigrate.auth import AppPool
from jtet_pstmigrate.calendar_uploader import (
    CalendarUploader,
    CalendarUploadError,
)
from jtet_pstmigrate.config import AppConfig, MappingRow
from jtet_pstmigrate.contact_uploader import (
    ContactUploader,
    ContactUploadError,
)
from jtet_pstmigrate.folder_manager import FolderManager
from jtet_pstmigrate.graph_client import GraphClient, GraphError
from jtet_pstmigrate.pst_reader import (
    ExtractedAppointment,
    ExtractedContact,
    ExtractedMessage,
    ReadpstError,
    extract_pst,
    extract_pst_calendar,
    extract_pst_contacts,
    iter_appointments,
    iter_contacts,
    iter_messages,
)
from jtet_pstmigrate.state import StateStore
from jtet_pstmigrate.uploader import MessageUploader


@dataclasses.dataclass
class RunReport:
    pst_path: Path
    mailbox: str
    items_total: int = 0
    items_uploaded: int = 0
    items_skipped: int = 0
    items_failed: int = 0
    elapsed_seconds: float = 0.0
    status: str = "pending"
    last_error: str | None = None


def load_mapping(csv_path: Path) -> list[MappingRow]:
    rows: list[MappingRow] = []
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        required = {"PSTPath", "TargetMailbox"}
        if not required.issubset({c.strip() for c in (reader.fieldnames or [])}):
            raise ValueError(f"CSV missing required columns: {required}. Got: {reader.fieldnames}")
        for raw in reader:
            row = MappingRow(
                pst_path=Path(raw["PSTPath"].strip()),
                target_mailbox=raw["TargetMailbox"].strip(),
                target_root_folder=(raw.get("TargetRootFolder") or "").strip(),
            )
            rows.append(row)
    return rows


class Orchestrator:
    def __init__(self, cfg: AppConfig, state: StateStore, pool: AppPool):
        self._cfg = cfg
        self._state = state
        self._pool = pool
        self._console = Console()

    def run(self, mapping: list[MappingRow]) -> list[RunReport]:
        if not mapping:
            return []

        # One shared GraphClient — httpx is thread-safe for sync requests
        graph = GraphClient(self._pool, self._cfg.throttle)
        reports: list[RunReport] = []

        with graph, ThreadPoolExecutor(
            max_workers=self._cfg.migration.max_parallel_mailboxes,
            thread_name_prefix="pst-mbx",
        ) as pool:
            futures: dict[Future, MappingRow] = {
                pool.submit(self._run_one, graph, row): row for row in mapping
            }

            with self._make_progress() as progress:
                task = progress.add_task("PST jobs", total=len(futures))
                for fut in as_completed(futures):
                    row = futures[fut]
                    try:
                        rep = fut.result()
                    except Exception as e:
                        logger.bind(ctx=f"{row.target_mailbox}").exception("Worker crashed")
                        rep = RunReport(
                            pst_path=row.pst_path,
                            mailbox=row.target_mailbox,
                            status="failed",
                            last_error=str(e),
                        )
                    reports.append(rep)
                    progress.advance(task)

        self._render_summary(reports, graph)
        return reports

    def _run_one(self, graph: GraphClient, row: MappingRow) -> RunReport:
        log = logger.bind(ctx=f"job[{row.target_mailbox}|{row.pst_path.name}]")
        report = RunReport(pst_path=row.pst_path, mailbox=row.target_mailbox)
        started = time.time()
        pst_str = str(row.pst_path)

        self._state.start_pst_run(pst_str, row.target_mailbox)

        try:
            extracted_dir = extract_pst(
                row.pst_path,
                self._cfg.paths.work_dir,
                binary=self._cfg.paths.readpst_binary,
            )
        except ReadpstError as e:
            log.error("Extraction failed: {}", e)
            self._state.update_pst_run(pst_str, row.target_mailbox, status="failed", last_error=str(e))
            report.status = "failed"
            report.last_error = str(e)
            report.elapsed_seconds = time.time() - started
            return report

        messages = list(iter_messages(extracted_dir))
        report.items_total = len(messages)
        self._state.update_pst_run(pst_str, row.target_mailbox, status="uploading", items_total=len(messages))
        log.info("{} messages to consider", len(messages))

        if not messages:
            self._state.update_pst_run(pst_str, row.target_mailbox, status="done")
            report.status = "done"
            report.elapsed_seconds = time.time() - started
            return report

        folders = FolderManager(graph, self._state, row.target_mailbox)
        uploader = MessageUploader(graph, row.target_mailbox, self._cfg.migration.large_attachment_threshold_bytes)
        # `target_root_folder` is intentionally ignored. Mail is routed into
        # the mailbox's real Outlook folder structure (Inbox / Sent Items /
        # etc.) by FolderManager; there is no 'Imported PST' wrapper.

        # Upload in parallel within this mailbox
        with ThreadPoolExecutor(
            max_workers=self._cfg.migration.workers_per_mailbox,
            thread_name_prefix=f"up-{row.target_mailbox.split('@')[0][:6]}",
        ) as up_pool:
            futures = {
                up_pool.submit(self._upload_one, msg, row, folders, uploader): msg
                for msg in messages
            }
            for fut in as_completed(futures):
                outcome = fut.result()
                if outcome == "uploaded":
                    report.items_uploaded += 1
                elif outcome == "skipped":
                    report.items_skipped += 1
                else:
                    report.items_failed += 1

        if report.items_failed and self._cfg.migration.fail_fast:
            self._state.update_pst_run(pst_str, row.target_mailbox, status="failed", last_error=f"{report.items_failed} item failures")
            report.status = "failed"
        else:
            self._state.update_pst_run(pst_str, row.target_mailbox, status="done")
            report.status = "done"
        report.elapsed_seconds = time.time() - started
        return report

    def _upload_one(
        self,
        msg: ExtractedMessage,
        row: MappingRow,
        folders: FolderManager,
        uploader: MessageUploader,
    ) -> str:
        pst_str = str(row.pst_path)
        src = str(msg.file_path)
        # If THIS exact .eml file is already marked done, don't touch the row
        # (touching it as 'skipped' would downgrade it and break dedupe on the
        # next run, causing the message to be re-uploaded as a duplicate).
        if self._state.is_row_done(row.target_mailbox, pst_str, src):
            return "skipped"
        if self._state.is_message_done(row.target_mailbox, msg.dedupe_key):
            # Different file with the same internet message-id (e.g. a copy in
            # a different folder). Record it as a skipped *new* row so the
            # audit shows which copies were de-duplicated.
            self._state.upsert_message(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=src,
                dedupe_key=msg.dedupe_key,
                status="skipped",
                bytes_=msg.bytes_,
            )
            return "skipped"
        # Pin one app for this message so all of its requests (folder ensure,
        # message create, attachment upload) share the same throttle bucket
        # and stats attribution.
        chosen_app = self._pool.pick()
        try:
            folder_id = folders.ensure_path(msg.folder_path)
            result = uploader.upload(msg, folder_id, app_id=chosen_app)
            self._state.upsert_message(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=str(msg.file_path),
                dedupe_key=msg.dedupe_key,
                graph_message_id=result.graph_message_id,
                graph_folder_id=folder_id,
                app_id=chosen_app,
                status="done",
                bytes_=result.bytes_uploaded,
            )
            return "uploaded"
        except GraphError as e:
            # Exchange caps individual messages at 150 MB; nothing we can do
            # about a true ErrorMessageSizeExceeded — record as skipped so
            # we don't keep retrying it forever.
            body_str = str(e.body) if e.body is not None else ""
            terminal = "ErrorMessageSizeExceeded" in body_str
            status = "skipped" if terminal else "failed"
            self._state.upsert_message(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=str(msg.file_path),
                dedupe_key=msg.dedupe_key,
                app_id=chosen_app,
                status=status,
                bytes_=msg.bytes_,
                last_error=f"graph {e.status}: {body_str[:300]}",
            )
            return "skipped" if terminal else "failed"
        except Exception as e:
            self._state.upsert_message(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=str(msg.file_path),
                dedupe_key=msg.dedupe_key,
                app_id=chosen_app,
                status="failed",
                bytes_=msg.bytes_,
                last_error=f"{type(e).__name__}: {str(e)[:300]}",
            )
            return "failed"

    def _make_progress(self) -> Progress:
        return Progress(
            SpinnerColumn(),
            TextColumn("[bold]{task.description}"),
            BarColumn(bar_width=None),
            MofNCompleteColumn(),
            TimeElapsedColumn(),
            TimeRemainingColumn(),
            console=self._console,
            transient=False,
        )

    # ------------------------------------------------------------------
    # Calendar import
    # ------------------------------------------------------------------
    #
    # The calendar path mirrors the mail path but with three differences:
    #   - readpst is run with -t a (appointments) into a separate work dir
    #   - items are deduped via the ``non_mail_items`` table, not ``messages``
    #   - they're POSTed as Graph events to the user's default calendar,
    #     not as messages into a folder hierarchy
    #
    # Folder routing (Inbox/Sent/etc.) doesn't apply: a mailbox has exactly
    # one default calendar, and we don't recreate sub-calendar structure
    # in v1. Sub-folders inside the PST's Calendar tree (custom calendars
    # the user kept) are flattened into the default calendar; the original
    # source folder is recoverable from ``non_mail_items.source_path`` if
    # we ever need to revisit that.

    def run_calendar(self, mapping: list[MappingRow]) -> list[RunReport]:
        if not mapping:
            return []

        graph = GraphClient(self._pool, self._cfg.throttle)
        reports: list[RunReport] = []

        with graph, ThreadPoolExecutor(
            max_workers=self._cfg.migration.max_parallel_mailboxes,
            thread_name_prefix="pst-cal",
        ) as pool:
            futures: dict[Future, MappingRow] = {
                pool.submit(self._run_calendar_one, graph, row): row for row in mapping
            }

            with self._make_progress() as progress:
                task = progress.add_task("PST calendar jobs", total=len(futures))
                for fut in as_completed(futures):
                    row = futures[fut]
                    try:
                        rep = fut.result()
                    except Exception as e:
                        logger.bind(ctx=f"{row.target_mailbox}").exception("Calendar worker crashed")
                        rep = RunReport(
                            pst_path=row.pst_path,
                            mailbox=row.target_mailbox,
                            status="failed",
                            last_error=str(e),
                        )
                    reports.append(rep)
                    progress.advance(task)

        self._render_summary(reports, graph, title_suffix=" (calendar)")
        return reports

    def _run_calendar_one(self, graph: GraphClient, row: MappingRow) -> RunReport:
        log = logger.bind(ctx=f"caljob[{row.target_mailbox}|{row.pst_path.name}]")
        report = RunReport(pst_path=row.pst_path, mailbox=row.target_mailbox)
        started = time.time()

        try:
            extracted_dir = extract_pst_calendar(
                row.pst_path,
                self._cfg.paths.work_dir,
                binary=self._cfg.paths.readpst_binary,
            )
        except ReadpstError as e:
            log.error("Calendar extraction failed: {}", e)
            report.status = "failed"
            report.last_error = str(e)
            report.elapsed_seconds = time.time() - started
            return report

        appointments = list(iter_appointments(extracted_dir))
        report.items_total = len(appointments)
        log.info("{} appointments to consider", len(appointments))

        if not appointments:
            report.status = "done"
            report.elapsed_seconds = time.time() - started
            return report

        uploader = CalendarUploader(graph, row.target_mailbox)

        with ThreadPoolExecutor(
            max_workers=self._cfg.migration.workers_per_mailbox,
            thread_name_prefix=f"cal-{row.target_mailbox.split('@')[0][:6]}",
        ) as up_pool:
            futures = {
                up_pool.submit(self._upload_appointment, appt, row, uploader): appt
                for appt in appointments
            }
            for fut in as_completed(futures):
                outcome = fut.result()
                if outcome == "uploaded":
                    report.items_uploaded += 1
                elif outcome == "skipped":
                    report.items_skipped += 1
                else:
                    report.items_failed += 1

        if report.items_failed and self._cfg.migration.fail_fast:
            report.status = "failed"
            report.last_error = f"{report.items_failed} item failures"
        else:
            report.status = "done"
        report.elapsed_seconds = time.time() - started
        return report

    def _upload_appointment(
        self,
        appt: ExtractedAppointment,
        row: MappingRow,
        uploader: CalendarUploader,
    ) -> str:
        pst_str = str(row.pst_path)
        src = str(appt.file_path)
        # Same dedup discipline as messages: per-row done check first (don't
        # downgrade an existing 'done' to 'skipped'), then per-mailbox UID
        # check (handles the same UID showing up under two folders inside
        # the same PST or across PSTs assigned to one mailbox).
        if self._state.is_item_row_done(row.target_mailbox, pst_str, src, "event"):
            return "skipped"
        if self._state.is_item_done(row.target_mailbox, "event", appt.dedupe_key):
            self._state.upsert_item(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=src,
                item_type="event",
                dedupe_key=appt.dedupe_key,
                status="skipped",
                bytes_=appt.bytes_,
            )
            return "skipped"

        chosen_app = self._pool.pick()
        try:
            result = uploader.upload(appt, app_id=chosen_app)
            self._state.upsert_item(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=src,
                item_type="event",
                dedupe_key=appt.dedupe_key,
                graph_id=result.graph_event_id,
                app_id=chosen_app,
                status="done",
                bytes_=result.bytes_uploaded,
            )
            return "uploaded"
        except CalendarUploadError as e:
            # Unparseable .ics or no VEVENT -- terminal, mark skipped so we
            # don't keep retrying it on the next run.
            self._state.upsert_item(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=src,
                item_type="event",
                dedupe_key=appt.dedupe_key,
                app_id=chosen_app,
                status="skipped",
                bytes_=appt.bytes_,
                last_error=str(e)[:300],
            )
            return "skipped"
        except GraphError as e:
            body_str = str(e.body) if e.body is not None else ""
            self._state.upsert_item(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=src,
                item_type="event",
                dedupe_key=appt.dedupe_key,
                app_id=chosen_app,
                status="failed",
                bytes_=appt.bytes_,
                last_error=f"graph {e.status}: {body_str[:300]}",
            )
            return "failed"
        except Exception as e:
            self._state.upsert_item(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=src,
                item_type="event",
                dedupe_key=appt.dedupe_key,
                app_id=chosen_app,
                status="failed",
                bytes_=appt.bytes_,
                last_error=f"{type(e).__name__}: {str(e)[:300]}",
            )
            return "failed"

    # ------------------------------------------------------------------
    # Calendar purge
    # ------------------------------------------------------------------
    #
    # Used to undo a calendar import after a code-side bug fix. Walks the
    # ``non_mail_items`` table for the selected (mailbox, PST) pairs,
    # DELETEs each event from Graph by stored ``graph_id``, and removes
    # the row so the next ``import-calendar`` re-uploads from scratch.
    #
    # 404s from Graph (already deleted in OWA) are treated as success --
    # the purge's job is to reach a "no row, no event" state, however we
    # got there. Other Graph errors are logged but don't stop the run; if
    # half the purge succeeds and half fails, the user can re-run the
    # purge to retry the leftover rows.

    def purge_calendar(self, mapping: list[MappingRow]) -> tuple[int, int, int]:
        """DELETE imported calendar events and clear their state rows.

        Returns ``(deleted_in_graph, missing_in_graph, errors)``.
        """
        if not mapping:
            return (0, 0, 0)

        graph = GraphClient(self._pool, self._cfg.throttle)
        deleted = 0
        missing = 0
        errors = 0

        with graph:
            for row in mapping:
                items = self._state.list_done_items(
                    row.target_mailbox, str(row.pst_path), "event"
                )
                if not items:
                    continue
                log = logger.bind(
                    ctx=f"purge[{row.target_mailbox}|{row.pst_path.name}]"
                )
                log.info("Purging {} event(s) from Graph", len(items))
                for item in items:
                    graph_id = item["graph_id"]
                    app_id = item["app_id"] or self._pool.names[0]
                    try:
                        graph.delete(
                            f"/users/{quote(row.target_mailbox)}/events/{quote(graph_id)}",
                            expect_status=(204, 200),
                            app_id=app_id,
                        )
                        deleted += 1
                    except GraphError as e:
                        if e.status == 404:
                            # Already gone -- still a success from our POV.
                            missing += 1
                        else:
                            log.warning(
                                "DELETE failed for {}: {} {}", graph_id, e.status, e.body
                            )
                            errors += 1
                            continue
                    except Exception as e:
                        log.warning("DELETE failed for {}: {}", graph_id, e)
                        errors += 1
                        continue
                    self._state.delete_item(
                        row.target_mailbox, str(row.pst_path), item["source_path"], "event"
                    )
        return (deleted, missing, errors)

    # ------------------------------------------------------------------
    # Mail purge
    # ------------------------------------------------------------------
    #
    # Two-phase delete that matches Search-Mailbox semantics on a Graph-only
    # codepath (Search-Mailbox itself is unreachable from the EXO V3
    # PowerShell module, which dropped Remote PowerShell -- so we do the
    # equivalent here):
    #
    #   Phase 1: DELETE /users/{upn}/messages/{graph_message_id}
    #            for every row we recorded as 'done'. This soft-deletes
    #            the message: Graph moves it from its current folder
    #            (Inbox / sub-folder / etc.) into the user's Deleted
    #            Items folder. The id assigned in Deleted Items differs
    #            from the original, so phase 1 is one-shot per id.
    #
    #   Phase 2: page through /users/{upn}/mailFolders/deletedItems/messages
    #            and DELETE every entry whose internetMessageId is in the
    #            set we just soft-deleted. A second DELETE on a message
    #            that is already in Deleted Items moves it to Recoverable
    #            Items (the hidden 14-day retention area), which is what
    #            users mean by "permanent" -- the message disappears from
    #            every visible folder. Targeting by internetMessageId
    #            keeps us from touching the user's pre-existing deleted
    #            items.
    #
    # Why not call Graph's `permanentDelete` action and skip phase 2?
    # That endpoint only exists on /beta and isn't a contractually stable
    # surface. The DELETE-twice approach uses /v1.0 for both phases,
    # which is what we already trust for every other request.
    #
    # Throughput: phase 1 parallelises per-mailbox via ThreadPoolExecutor
    # (workers_per_mailbox) and per-batch via max_parallel_mailboxes.
    # Phase 2 enumerates Deleted Items in 999-item pages (Graph's max);
    # the GET pages serialise per mailbox but the resulting DELETEs are
    # parallel again. Realistic end-to-end throughput sits around
    # 200-400 messages/sec across the app pool.

    def purge_mail(self, mapping: list[MappingRow]) -> tuple[int, int, int]:
        """DELETE imported mail messages and clear their state rows.

        Returns ``(deleted_in_graph, missing_in_graph, errors)``. ``deleted``
        counts phase-1 successes (message removed from its origin folder);
        ``missing`` counts phase-1 404s (message already gone); ``errors``
        counts everything else, including phase-2 failures (items that were
        soft-deleted but not permanently purged from Deleted Items).
        """
        if not mapping:
            return (0, 0, 0)

        graph = GraphClient(self._pool, self._cfg.throttle)
        totals = {"deleted": 0, "missing": 0, "errors": 0}
        totals_lock = threading.Lock()

        with graph, ThreadPoolExecutor(
            max_workers=self._cfg.migration.max_parallel_mailboxes,
            thread_name_prefix="pst-purge-mail",
        ) as pool:
            futures: dict[Future, MappingRow] = {
                pool.submit(self._purge_mail_one, graph, row): row for row in mapping
            }
            with self._make_progress() as progress:
                task = progress.add_task("Purge mail jobs", total=len(futures))
                for fut in as_completed(futures):
                    row = futures[fut]
                    try:
                        d, m, e = fut.result()
                    except Exception as exc:
                        logger.bind(ctx=row.target_mailbox).exception(
                            "Purge worker crashed"
                        )
                        d, m, e = (0, 0, 1)
                        del exc
                    with totals_lock:
                        totals["deleted"] += d
                        totals["missing"] += m
                        totals["errors"] += e
                    progress.advance(task)

        return (totals["deleted"], totals["missing"], totals["errors"])

    def _purge_mail_one(
        self, graph: GraphClient, row: MappingRow
    ) -> tuple[int, int, int]:
        items = self._state.list_done_messages(
            row.target_mailbox, str(row.pst_path)
        )
        if not items:
            return (0, 0, 0)

        log = logger.bind(
            ctx=f"purge-mail[{row.target_mailbox}|{row.pst_path.name}]"
        )
        log.info("Phase 1 (soft delete): {} message(s)", len(items))

        # imid:<lower-cased-Message-ID> -> phase-2 can find this in Deleted
        # Items by matching internetMessageId. sha256:<hex> -> source had no
        # Message-ID header, so phase-2 can't address it; we still soft-delete
        # in phase 1 and the user can manually empty Deleted Items if needed.
        imids_to_purge: set[str] = set()
        deleted = 0
        missing = 0
        errors = 0

        workers = max(1, self._cfg.migration.workers_per_mailbox)
        upn = quote(row.target_mailbox)

        def _phase1_one(item) -> tuple[str, str | None]:
            graph_id = item["graph_message_id"]
            app_id = item["app_id"] or self._pool.names[0]
            dedupe_key = item["dedupe_key"] or ""
            try:
                graph.delete(
                    f"/users/{upn}/messages/{quote(graph_id)}",
                    expect_status=(204, 200),
                    app_id=app_id,
                )
                outcome = "deleted"
            except GraphError as e:
                if e.status == 404:
                    outcome = "missing"
                else:
                    log.warning(
                        "phase-1 DELETE failed for {}: {} {}",
                        graph_id, e.status, e.body,
                    )
                    return ("error", None)
            except Exception as e:
                log.warning("phase-1 DELETE failed for {}: {}", graph_id, e)
                return ("error", None)

            # Drop the state row whether it was soft-deleted or 'already
            # gone'; either way a follow-up `import` should re-upload.
            self._state.delete_message_row(
                row.target_mailbox, str(row.pst_path), item["source_path"]
            )
            imid = (
                dedupe_key[len("imid:"):]
                if dedupe_key.startswith("imid:")
                else None
            )
            return (outcome, imid)

        with ThreadPoolExecutor(
            max_workers=workers, thread_name_prefix="purge-p1"
        ) as p1:
            for outcome, imid in p1.map(_phase1_one, items):
                if outcome == "deleted":
                    deleted += 1
                elif outcome == "missing":
                    missing += 1
                else:
                    errors += 1
                if imid:
                    imids_to_purge.add(imid.lower())

        if not imids_to_purge:
            log.info(
                "Phase 2 skipped: no imid-keyed messages to permanently delete"
            )
            return (deleted, missing, errors)

        # Phase 2 -- enumerate Deleted Items, match by internetMessageId.
        log.info(
            "Phase 2 (permanent delete): scanning Deleted Items for "
            "{} target message(s)", len(imids_to_purge),
        )

        next_url: str | None = (
            f"/users/{upn}/mailFolders/deletedItems/messages"
            f"?$select=id,internetMessageId&$top=999"
        )
        delete_ids: list[str] = []
        pages = 0
        while next_url:
            try:
                resp = graph.get(next_url, expect_status=(200,))
            except Exception as e:
                log.warning("phase-2 page GET failed: {}", e)
                errors += 1
                break
            body = resp.json()
            for msg in body.get("value", []):
                raw = (msg.get("internetMessageId") or "").strip("<> ").lower()
                if raw and raw in imids_to_purge:
                    delete_ids.append(msg["id"])
            next_url = body.get("@odata.nextLink")
            # @odata.nextLink is an absolute URL; httpx with a base_url will
            # ignore the base when the path itself is absolute, so passing it
            # through is fine.
            pages += 1
            if pages > 1000:
                log.warning(
                    "phase-2 paging gave up after 1000 pages -- runaway "
                    "Deleted Items folder?"
                )
                errors += 1
                break

        if not delete_ids:
            log.warning(
                "Phase 2 matched 0 messages in Deleted Items. The phase-1 "
                "soft-deletes succeeded but the items are no longer "
                "addressable -- either an inbox rule moved them, or a prior "
                "purge already permanently deleted them."
            )
            return (deleted, missing, errors)

        log.info(
            "Phase 2: deleting {} matched message(s) from Deleted Items",
            len(delete_ids),
        )

        def _phase2_one(graph_id: str) -> str:
            try:
                graph.delete(
                    f"/users/{upn}/messages/{quote(graph_id)}",
                    expect_status=(204, 200),
                )
                return "deleted"
            except GraphError as e:
                if e.status == 404:
                    return "missing"
                log.warning(
                    "phase-2 DELETE failed for {}: {} {}",
                    graph_id, e.status, e.body,
                )
                return "error"
            except Exception as e:
                log.warning("phase-2 DELETE failed for {}: {}", graph_id, e)
                return "error"

        with ThreadPoolExecutor(
            max_workers=workers, thread_name_prefix="purge-p2"
        ) as p2:
            for outcome in p2.map(_phase2_one, delete_ids):
                # Phase-2 success is a continuation of phase-1 success; we
                # already counted those messages in `deleted`. Phase-2
                # 'missing' is a benign race. Phase-2 errors leave the
                # message visible in Deleted Items, which the user
                # cares about, so we surface them as errors.
                if outcome == "error":
                    errors += 1

        return (deleted, missing, errors)

    # ------------------------------------------------------------------
    # Contacts (Stage C, second half)
    # ------------------------------------------------------------------
    #
    # Same pattern as calendar: extract .vcf with ``readpst -t c`` into a
    # sibling work dir, queue per-vcard, upload via Graph
    # /users/{upn}/contacts. Contacts state lives in ``non_mail_items``
    # with ``item_type='contact'`` and a UID/FN+email dedup key.

    def run_contacts(self, mapping: list[MappingRow]) -> list[RunReport]:
        if not mapping:
            return []

        graph = GraphClient(self._pool, self._cfg.throttle)
        reports: list[RunReport] = []

        with graph, ThreadPoolExecutor(
            max_workers=self._cfg.migration.max_parallel_mailboxes,
            thread_name_prefix="pst-con",
        ) as pool:
            futures: dict[Future, MappingRow] = {
                pool.submit(self._run_contacts_one, graph, row): row for row in mapping
            }

            with self._make_progress() as progress:
                task = progress.add_task("PST contacts jobs", total=len(futures))
                for fut in as_completed(futures):
                    row = futures[fut]
                    try:
                        rep = fut.result()
                    except Exception as e:
                        logger.bind(ctx=f"{row.target_mailbox}").exception("Contacts worker crashed")
                        rep = RunReport(
                            pst_path=row.pst_path,
                            mailbox=row.target_mailbox,
                            status="failed",
                            last_error=str(e),
                        )
                    reports.append(rep)
                    progress.advance(task)

        self._render_summary(reports, graph, title_suffix=" (contacts)")
        return reports

    def _run_contacts_one(self, graph: GraphClient, row: MappingRow) -> RunReport:
        log = logger.bind(ctx=f"conjob[{row.target_mailbox}|{row.pst_path.name}]")
        report = RunReport(pst_path=row.pst_path, mailbox=row.target_mailbox)
        started = time.time()

        try:
            extracted_dir = extract_pst_contacts(
                row.pst_path,
                self._cfg.paths.work_dir,
                binary=self._cfg.paths.readpst_binary,
            )
        except ReadpstError as e:
            log.error("Contacts extraction failed: {}", e)
            report.status = "failed"
            report.last_error = str(e)
            report.elapsed_seconds = time.time() - started
            return report

        contacts = list(iter_contacts(extracted_dir))
        report.items_total = len(contacts)
        log.info("{} contacts to consider", len(contacts))

        if not contacts:
            report.status = "done"
            report.elapsed_seconds = time.time() - started
            return report

        uploader = ContactUploader(graph, row.target_mailbox)

        with ThreadPoolExecutor(
            max_workers=self._cfg.migration.workers_per_mailbox,
            thread_name_prefix=f"con-{row.target_mailbox.split('@')[0][:6]}",
        ) as up_pool:
            futures = {
                up_pool.submit(self._upload_contact, c, row, uploader): c
                for c in contacts
            }
            for fut in as_completed(futures):
                outcome = fut.result()
                if outcome == "uploaded":
                    report.items_uploaded += 1
                elif outcome == "skipped":
                    report.items_skipped += 1
                else:
                    report.items_failed += 1

        if report.items_failed and self._cfg.migration.fail_fast:
            report.status = "failed"
            report.last_error = f"{report.items_failed} item failures"
        else:
            report.status = "done"
        report.elapsed_seconds = time.time() - started
        return report

    def _upload_contact(
        self,
        contact: ExtractedContact,
        row: MappingRow,
        uploader: ContactUploader,
    ) -> str:
        pst_str = str(row.pst_path)
        src = str(contact.file_path)
        # Same dedup discipline as messages/events: per-row done check
        # first, then per-mailbox dedup on UID or FN+email.
        if self._state.is_item_row_done(row.target_mailbox, pst_str, src, "contact"):
            return "skipped"
        if self._state.is_item_done(row.target_mailbox, "contact", contact.dedupe_key):
            self._state.upsert_item(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=src,
                item_type="contact",
                dedupe_key=contact.dedupe_key,
                status="skipped",
                bytes_=contact.bytes_,
            )
            return "skipped"

        chosen_app = self._pool.pick()
        try:
            result = uploader.upload(contact, app_id=chosen_app)
            self._state.upsert_item(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=src,
                item_type="contact",
                dedupe_key=contact.dedupe_key,
                graph_id=result.graph_contact_id,
                app_id=chosen_app,
                status="done",
                bytes_=result.bytes_uploaded,
            )
            return "uploaded"
        except ContactUploadError as e:
            # Empty/garbage vCard -- terminal, mark skipped so we don't
            # retry on the next run.
            self._state.upsert_item(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=src,
                item_type="contact",
                dedupe_key=contact.dedupe_key,
                app_id=chosen_app,
                status="skipped",
                bytes_=contact.bytes_,
                last_error=str(e)[:300],
            )
            return "skipped"
        except GraphError as e:
            body_str = str(e.body) if e.body is not None else ""
            self._state.upsert_item(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=src,
                item_type="contact",
                dedupe_key=contact.dedupe_key,
                app_id=chosen_app,
                status="failed",
                bytes_=contact.bytes_,
                last_error=f"graph {e.status}: {body_str[:300]}",
            )
            return "failed"
        except Exception as e:
            self._state.upsert_item(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=src,
                item_type="contact",
                dedupe_key=contact.dedupe_key,
                app_id=chosen_app,
                status="failed",
                bytes_=contact.bytes_,
                last_error=f"{type(e).__name__}: {str(e)[:300]}",
            )
            return "failed"

    def purge_contacts(self, mapping: list[MappingRow]) -> tuple[int, int, int]:
        """DELETE imported contacts and clear their state rows.

        Mirrors :meth:`purge_calendar`. Returns
        ``(deleted_in_graph, missing_in_graph, errors)``. 404s are
        treated as success ("already gone"); any other Graph error is
        logged and counted but doesn't abort the rest of the purge.
        """
        if not mapping:
            return (0, 0, 0)

        graph = GraphClient(self._pool, self._cfg.throttle)
        deleted = 0
        missing = 0
        errors = 0

        with graph:
            for row in mapping:
                items = self._state.list_done_items(
                    row.target_mailbox, str(row.pst_path), "contact"
                )
                if not items:
                    continue
                log = logger.bind(
                    ctx=f"purge[{row.target_mailbox}|{row.pst_path.name}]"
                )
                log.info("Purging {} contact(s) from Graph", len(items))
                for item in items:
                    graph_id = item["graph_id"]
                    app_id = item["app_id"] or self._pool.names[0]
                    try:
                        graph.delete(
                            f"/users/{quote(row.target_mailbox)}/contacts/{quote(graph_id)}",
                            expect_status=(204, 200),
                            app_id=app_id,
                        )
                        deleted += 1
                    except GraphError as e:
                        if e.status == 404:
                            missing += 1
                        else:
                            log.warning(
                                "DELETE failed for {}: {} {}", graph_id, e.status, e.body
                            )
                            errors += 1
                            continue
                    except Exception as e:
                        log.warning("DELETE failed for {}: {}", graph_id, e)
                        errors += 1
                        continue
                    self._state.delete_item(
                        row.target_mailbox, str(row.pst_path), item["source_path"], "contact"
                    )
        return (deleted, missing, errors)

    def _render_summary(
        self,
        reports: Iterable[RunReport],
        graph: GraphClient,
        *,
        title_suffix: str = "",
    ) -> None:
        table = Table(
            title=f"Migration Summary{title_suffix}  ({datetime.now().isoformat(timespec='seconds')})"
        )
        table.add_column("Mailbox", overflow="fold")
        table.add_column("PST", overflow="fold")
        table.add_column("Total", justify="right")
        table.add_column("Uploaded", justify="right", style="green")
        table.add_column("Skipped", justify="right", style="yellow")
        table.add_column("Failed", justify="right", style="red")
        table.add_column("Elapsed", justify="right")
        table.add_column("Status")

        for r in reports:
            table.add_row(
                r.mailbox,
                r.pst_path.name,
                str(r.items_total),
                str(r.items_uploaded),
                str(r.items_skipped),
                str(r.items_failed),
                f"{r.elapsed_seconds:0.0f}s",
                r.status,
            )

        self._console.print(table)

        # Per-app breakdown — useful for spotting unbalanced load or one app
        # taking the brunt of the throttling.
        per_app = Table(title="Graph throughput per app")
        per_app.add_column("App")
        per_app.add_column("Requests", justify="right")
        per_app.add_column("429/503", justify="right", style="yellow")
        per_app.add_column("5xx retries", justify="right", style="yellow")
        per_app.add_column("Backoff (s)", justify="right", style="yellow")
        for name, s in graph.stats.items():
            per_app.add_row(
                name,
                str(s.requests),
                str(s.retries_429),
                str(s.retries_5xx),
                f"{s.total_backoff_seconds:0.1f}",
            )
        self._console.print(per_app)

        total_req = sum(s.requests for s in graph.stats.values())
        total_429 = sum(s.retries_429 for s in graph.stats.values())
        total_5xx = sum(s.retries_5xx for s in graph.stats.values())
        total_backoff = sum(s.total_backoff_seconds for s in graph.stats.values())
        self._console.print(
            f"\nTotal Graph requests: [bold]{total_req}[/]  "
            f"throttled-retries: [yellow]{total_429}[/]  "
            f"server-error-retries: [yellow]{total_5xx}[/]  "
            f"total backoff: [yellow]{total_backoff:0.1f}s[/]"
        )
