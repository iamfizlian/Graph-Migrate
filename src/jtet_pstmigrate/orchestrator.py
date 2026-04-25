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
import time
from collections.abc import Iterable
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

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
from jtet_pstmigrate.graph_client import GraphClient, GraphError
from jtet_pstmigrate.pst_reader import (
    ExtractedMessage,
    ReadpstError,
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
        root: str | None,
    ) -> str:
        pst_str = str(row.pst_path)
        if self._state.is_message_done(row.target_mailbox, msg.dedupe_key):
            self._state.upsert_message(
                mailbox=row.target_mailbox,
                pst_path=pst_str,
                source_path=str(msg.file_path),
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

    def _render_summary(self, reports: Iterable[RunReport], graph: GraphClient) -> None:
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
