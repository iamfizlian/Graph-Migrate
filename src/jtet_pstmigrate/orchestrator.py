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
from typing import Protocol, runtime_checkable

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
from jtet_pstmigrate.config import AppConfig, MappingRow
from jtet_pstmigrate.folder_manager import FolderManager
from jtet_pstmigrate.graph_client import GraphClient, GraphError, ThrottleStats
from jtet_pstmigrate.pst_reader import (
    ExtractedMessage,
    ReadpstError,
    check_readpst,
    extract_pst,
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
    items_cancelled: int = 0
    elapsed_seconds: float = 0.0
    status: str = "pending"  # pending|done|failed|cancelled
    last_error: str | None = None


@dataclasses.dataclass
class PreflightCheck:
    """One row of the pre-flight report.

    Kept intentionally tiny so it serializes cleanly to JSON for a web UI
    and renders to one row of a Rich table for the CLI without translation.
    """

    name: str
    ok: bool
    detail: str


@dataclasses.dataclass
class PreflightReport:
    checks: list[PreflightCheck]

    @property
    def all_ok(self) -> bool:
        return all(c.ok for c in self.checks)


def preflight(cfg: AppConfig, rows: list[MappingRow]) -> PreflightReport:
    """Run all environment checks and return a structured report.

    Pure function: doesn't print, doesn't exit, doesn't mutate state. The
    CLI renders the result as a Rich table; a web route can dump the same
    dataclass straight to JSON. Mapping CSV parse failures are NOT covered
    here — callers load the mapping first (so they can also apply filter
    flags) and pass the resulting rows in.
    """
    checks: list[PreflightCheck] = []

    # readpst on PATH
    try:
        ver = check_readpst(cfg.paths.readpst_binary)
        checks.append(PreflightCheck("readpst available", True, ver))
    except Exception as e:
        checks.append(PreflightCheck("readpst available", False, str(e)))

    # PST files exist on disk
    missing = [r.pst_path for r in rows if not r.pst_path.exists()]
    if missing:
        sample = ", ".join(str(p) for p in missing[:5])
        checks.append(PreflightCheck("PST files exist", False, f"missing: {sample}"))
    elif rows:
        total_gb = sum(r.pst_path.stat().st_size for r in rows) / 1024**3
        checks.append(PreflightCheck("PST files exist", True, f"{len(rows)} files, {total_gb:0.2f} GB total"))

    # Graph token + mailbox resolution — verify each app independently
    try:
        pool = AppPool(cfg.apps)
        per_app_ok: list[str] = []
        per_app_fail: list[str] = []
        with GraphClient(pool, cfg.throttle) as graph:
            for name in pool.names:
                try:
                    graph.get("/$metadata", expect_status=(200,), app_id=name)
                    per_app_ok.append(name)
                except Exception as e:
                    per_app_fail.append(f"{name}: {e}")
        if per_app_fail:
            checks.append(
                PreflightCheck("Graph token (per app)", False, "; ".join(per_app_fail[:3]))
            )
        else:
            checks.append(
                PreflightCheck(
                    "Graph token (per app)",
                    True,
                    f"{len(per_app_ok)} app(s): {', '.join(per_app_ok)}",
                )
            )

        # Mailbox accessibility — uses the same Mail.ReadWrite scope the
        # actual import will use, so failures here are the failures the
        # import would hit (missing mailbox, blocked by ApplicationAccessPolicy,
        # etc.) instead of a User.Read.All-shaped probe.
        unique = sorted({r.target_mailbox for r in rows})
        unresolved: list[str] = []
        with GraphClient(pool, cfg.throttle) as graph:
            for upn in unique:
                try:
                    graph.get(
                        f"/users/{upn}/mailFolders/inbox",
                        params={"$select": "id,displayName"},
                        app_id=pool.names[0],
                    )
                except Exception as e:
                    unresolved.append(f"{upn} ({e})")
        if unresolved:
            checks.append(
                PreflightCheck(
                    "Mailboxes accessible",
                    False,
                    f"{len(unresolved)}/{len(unique)}: {unresolved[0]}",
                )
            )
        elif unique:
            checks.append(PreflightCheck("Mailboxes accessible", True, f"{len(unique)} mailboxes"))
    except Exception as e:
        checks.append(PreflightCheck("Graph token + connectivity", False, str(e)))

    return PreflightReport(checks=checks)


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


@runtime_checkable
class ProgressReporter(Protocol):
    """Pluggable observer for orchestrator lifecycle events.

    The default implementation (RichProgressReporter) drives the CLI's live
    progress bar and end-of-run summary tables. The web layer subscribes by
    passing a queue-backed reporter that pushes structured events to clients.
    Both reporters see the same events; only the rendering differs.

    Lifecycle: `start_run` is called once before the first job is dispatched,
    `job_completed` once per finished PST/mailbox pair (in completion order,
    not submission order), and `finish_run` exactly once at the end — even
    if the run errors out, so the reporter can tear down its display.
    """

    def start_run(self, job_count: int) -> None: ...
    def job_completed(self, report: RunReport) -> None: ...
    def finish_run(
        self,
        reports: list[RunReport],
        graph_stats: dict[str, ThrottleStats],
    ) -> None: ...


class NullReporter:
    """Silent reporter — useful for tests, web mode, or `--quiet` runs."""

    def start_run(self, job_count: int) -> None:
        return None

    def job_completed(self, report: RunReport) -> None:
        return None

    def finish_run(
        self,
        reports: list[RunReport],
        graph_stats: dict[str, ThrottleStats],
    ) -> None:
        return None


class RichProgressReporter:
    """Default CLI reporter: live PST-jobs bar + end-of-run summary tables.

    This is the behaviour that shipped before the Protocol was extracted —
    `pstmigrate import` looks identical from a user's perspective.
    """

    def __init__(self, console: Console | None = None):
        self._console = console or Console()
        self._progress: Progress | None = None
        self._task_id: int | None = None

    def start_run(self, job_count: int) -> None:
        self._progress = Progress(
            SpinnerColumn(),
            TextColumn("[bold]{task.description}"),
            BarColumn(bar_width=None),
            MofNCompleteColumn(),
            TimeElapsedColumn(),
            TimeRemainingColumn(),
            console=self._console,
            transient=False,
        )
        self._progress.__enter__()
        self._task_id = self._progress.add_task("PST jobs", total=job_count)

    def job_completed(self, report: RunReport) -> None:
        if self._progress is not None and self._task_id is not None:
            self._progress.advance(self._task_id)

    def finish_run(
        self,
        reports: list[RunReport],
        graph_stats: dict[str, ThrottleStats],
    ) -> None:
        # Tear down the live bar before printing tables — otherwise the
        # tables get interleaved with Rich's transient frame buffer.
        if self._progress is not None:
            self._progress.__exit__(None, None, None)
            self._progress = None
            self._task_id = None
        self._render_summary(reports, graph_stats)

    def _render_summary(
        self,
        reports: Iterable[RunReport],
        graph_stats: dict[str, ThrottleStats],
    ) -> None:
        table = Table(title=f"Migration Summary  ({datetime.now().isoformat(timespec='seconds')})")
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
        for name, s in graph_stats.items():
            per_app.add_row(
                name,
                str(s.requests),
                str(s.retries_429),
                str(s.retries_5xx),
                f"{s.total_backoff_seconds:0.1f}",
            )
        self._console.print(per_app)

        total_req = sum(s.requests for s in graph_stats.values())
        total_429 = sum(s.retries_429 for s in graph_stats.values())
        total_5xx = sum(s.retries_5xx for s in graph_stats.values())
        total_backoff = sum(s.total_backoff_seconds for s in graph_stats.values())
        self._console.print(
            f"\nTotal Graph requests: [bold]{total_req}[/]  "
            f"throttled-retries: [yellow]{total_429}[/]  "
            f"server-error-retries: [yellow]{total_5xx}[/]  "
            f"total backoff: [yellow]{total_backoff:0.1f}s[/]"
        )


class Orchestrator:
    def __init__(
        self,
        cfg: AppConfig,
        state: StateStore,
        pool: AppPool,
        reporter: ProgressReporter | None = None,
    ):
        self._cfg = cfg
        self._state = state
        self._pool = pool
        # Default to the CLI-flavoured reporter so existing call sites keep
        # working unchanged. Pass NullReporter() (or your own) for web/tests.
        self._reporter: ProgressReporter = reporter or RichProgressReporter()
        # Set per `run()` invocation. Default is a never-set Event so the
        # `is_set()` checks in workers are always cheap and falsy.
        self._cancel: threading.Event = threading.Event()

    def run(
        self,
        mapping: list[MappingRow],
        cancel: threading.Event | None = None,
    ) -> list[RunReport]:
        """Execute the migration. Pass `cancel` to expose a stop button.

        Cancellation is cooperative: workers check the flag before pulling
        the next message and bail with a 'cancelled' outcome. Already
        in-flight uploads run to completion (and have their result recorded
        normally) so resume semantics stay intact. The next `run()` picks
        up exactly where this one stopped, dedupe-driven.
        """
        if not mapping:
            return []

        self._cancel = cancel or threading.Event()

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

            self._reporter.start_run(len(futures))
            try:
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
                    self._reporter.job_completed(rep)
            finally:
                # Always called — even on exceptions — so the reporter
                # can stop its live display and emit a summary.
                self._reporter.finish_run(reports, dict(graph.stats))

        return reports

    def _run_one(self, graph: GraphClient, row: MappingRow) -> RunReport:
        log = logger.bind(ctx=f"job[{row.target_mailbox}|{row.pst_path.name}]")
        report = RunReport(pst_path=row.pst_path, mailbox=row.target_mailbox)
        started = time.time()
        pst_str = str(row.pst_path)

        # Bail before doing anything expensive if we were cancelled before
        # this PST's worker even started (common for queued jobs).
        if self._cancel.is_set():
            report.status = "cancelled"
            report.last_error = "cancelled before start"
            report.elapsed_seconds = time.time() - started
            return report

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
        root = row.target_root_folder or self._cfg.migration.target_root_folder or None

        # Upload in parallel within this mailbox
        with ThreadPoolExecutor(
            max_workers=self._cfg.migration.workers_per_mailbox,
            thread_name_prefix=f"up-{row.target_mailbox.split('@')[0][:6]}",
        ) as up_pool:
            futures = {
                up_pool.submit(self._upload_one, msg, row, folders, uploader, root): msg
                for msg in messages
            }
            for fut in as_completed(futures):
                outcome = fut.result()
                if outcome == "uploaded":
                    report.items_uploaded += 1
                elif outcome == "skipped":
                    report.items_skipped += 1
                elif outcome == "cancelled":
                    report.items_cancelled += 1
                else:
                    report.items_failed += 1

        # Cancellation takes priority over fail_fast — a cancelled run
        # with some failures should be marked 'cancelled' so resume on
        # the next invocation doesn't treat it as a permanent failure.
        if self._cancel.is_set():
            self._state.update_pst_run(
                pst_str,
                row.target_mailbox,
                status="cancelled",
                last_error=f"cancelled after {report.items_uploaded} uploads",
            )
            report.status = "cancelled"
            report.last_error = f"cancelled after {report.items_uploaded} uploads"
        elif report.items_failed and self._cfg.migration.fail_fast:
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
        root: str | None,
    ) -> str:
        # Cooperative cancellation: queued upload futures bail at this gate
        # without touching state. In-flight uploads (already past this check)
        # finish naturally so their state is recorded and resume works.
        if self._cancel.is_set():
            return "cancelled"
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
            folder_id = folders.ensure_path(msg.folder_path, root_folder=root)
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
