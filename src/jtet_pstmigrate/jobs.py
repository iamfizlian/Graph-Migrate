"""Web-safe job execution layer shared by frontends."""

from __future__ import annotations

import threading
import time
import uuid
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal

from jtet_pstmigrate.auth import AppPool
from jtet_pstmigrate.config import MappingRow
from jtet_pstmigrate.log import configure_logging
from jtet_pstmigrate.mapping import load_mapping
from jtet_pstmigrate.orchestrator import Orchestrator, RunReport
from jtet_pstmigrate.selection import SelectionFilters, select_mapping
from jtet_pstmigrate.services import load_config
from jtet_pstmigrate.state import StateStore
from jtet_pstmigrate.validation import ValidationReport, validate_environment

JobKind = Literal[
    "validate",
    "import-mail",
    "import-calendar",
    "import-contacts",
    "import-all",
    "purge-calendar",
    "purge-contacts",
    "purge-mail",
    "reset-state",
    "status",
]
JobStatus = Literal["queued", "running", "done", "failed", "blocked"]
JobCallback = Callable[["JobRecord"], None]

MUTATING_JOBS: set[JobKind] = {
    "import-mail",
    "import-calendar",
    "import-contacts",
    "import-all",
    "purge-calendar",
    "purge-contacts",
    "purge-mail",
    "reset-state",
}
DESTRUCTIVE_JOBS: set[JobKind] = {
    "purge-calendar",
    "purge-contacts",
    "purge-mail",
    "reset-state",
}


@dataclass(slots=True)
class JobSpec:
    kind: JobKind
    config_path: Path | None
    mapping_path: Path | None = None
    filters: SelectionFilters = field(default_factory=SelectionFilters)
    confirmed: bool = False


@dataclass(slots=True)
class JobRecord:
    job_id: str
    kind: JobKind
    status: JobStatus = "queued"
    phase: str = "queued"
    selected_rows: int = 0
    started_at: float | None = None
    finished_at: float | None = None
    reports: list[RunReport] = field(default_factory=list)
    validation: ValidationReport | None = None
    result: dict[str, int] = field(default_factory=dict)
    last_error: str | None = None

    @property
    def elapsed_seconds(self) -> float:
        if self.started_at is None:
            return 0.0
        return (self.finished_at or time.time()) - self.started_at

    @property
    def item_counts(self) -> dict[str, int]:
        return {
            "total": sum(r.items_total for r in self.reports),
            "uploaded": sum(r.items_uploaded for r in self.reports),
            "skipped": sum(r.items_skipped for r in self.reports),
            "failed": sum(r.items_failed for r in self.reports),
        }


def load_selected_rows(spec: JobSpec) -> tuple[list[MappingRow], int]:
    if spec.mapping_path is None:
        return ([], 0)
    all_rows = load_mapping(spec.mapping_path)
    selection = select_mapping(all_rows, spec.filters)
    return (selection.rows, selection.total_rows)


def run_job(spec: JobSpec, callbacks: list[JobCallback] | None = None) -> JobRecord:
    """Run a job synchronously and return its final record."""
    record = JobRecord(job_id=uuid.uuid4().hex[:12], kind=spec.kind)
    callbacks = callbacks or []

    def emit(phase: str) -> None:
        record.phase = phase
        for callback in callbacks:
            callback(record)

    try:
        if spec.kind in DESTRUCTIVE_JOBS and not spec.confirmed:
            record.status = "blocked"
            record.last_error = "Confirmation is required for destructive jobs."
            return record

        record.status = "running"
        record.started_at = time.time()
        emit("loading config")
        cfg = load_config(spec.config_path)
        log_id = f"{spec.kind}_{time.strftime('%Y%m%d_%H%M%S')}"
        configure_logging(cfg.paths.log_dir, cfg.log_level, run_id=log_id)

        emit("loading mapping")
        rows, _ = load_selected_rows(spec)
        record.selected_rows = len(rows)

        if spec.kind == "status":
            record.status = "done"
            emit("status ready")
            return record

        if spec.kind != "validate" and not rows:
            raise ValueError("Selection produced 0 mapping rows.")

        if spec.kind == "validate":
            emit("validating")
            record.validation = validate_environment(cfg, rows)
            record.status = "done" if record.validation.ok else "failed"
            return record

        state = StateStore(cfg.paths.state_dir / "state.sqlite")
        pool = AppPool(cfg.apps)
        orch = Orchestrator(cfg, state, pool)

        if spec.kind == "import-mail":
            emit("importing mail")
            record.reports = orch.run(rows)
        elif spec.kind == "import-calendar":
            emit("importing calendar")
            record.reports = orch.run_calendar(rows)
        elif spec.kind == "import-contacts":
            emit("importing contacts")
            record.reports = orch.run_contacts(rows)
        elif spec.kind == "import-all":
            emit("importing mail")
            record.reports.extend(orch.run(rows))
            emit("importing calendar")
            record.reports.extend(orch.run_calendar(rows))
            emit("importing contacts")
            record.reports.extend(orch.run_contacts(rows))
        elif spec.kind == "purge-calendar":
            emit("purging calendar")
            deleted, missing, errors = orch.purge_calendar(rows)
            record.result = {"deleted": deleted, "missing": missing, "errors": errors}
        elif spec.kind == "purge-contacts":
            emit("purging contacts")
            deleted, missing, errors = orch.purge_contacts(rows)
            record.result = {"deleted": deleted, "missing": missing, "errors": errors}
        elif spec.kind == "purge-mail":
            emit("purging mail")
            deleted, missing, errors = orch.purge_mail(rows)
            record.result = {"deleted": deleted, "missing": missing, "errors": errors}
        elif spec.kind == "reset-state":
            emit("resetting state")
            record.result = _reset_state(state, rows)

        failed_reports = sum(1 for r in record.reports if r.status != "done" or r.items_failed)
        result_errors = record.result.get("errors", 0)
        record.status = "failed" if failed_reports or result_errors else "done"
        return record
    except Exception as e:
        record.status = "failed"
        record.last_error = str(e)
        return record
    finally:
        record.finished_at = time.time()
        emit(record.status)


def _reset_state(state: StateStore, rows: list[MappingRow]) -> dict[str, int]:
    cleared = {"messages": 0, "non_mail_items": 0, "pst_runs": 0, "folder_map": 0}
    affected_mailboxes = {row.target_mailbox for row in rows}
    for row in rows:
        result = state.clear_scope(row.target_mailbox, str(row.pst_path))
        for key, value in result.items():
            cleared[key] += value
    for mailbox in affected_mailboxes:
        cleared["folder_map"] += state.clear_folder_map(mailbox)
    return cleared


class JobManager:
    """Run local web jobs with a single active mutating job by default."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._active_mutating: str | None = None
        self._jobs: dict[str, JobRecord] = {}

    def submit(self, spec: JobSpec) -> JobRecord:
        record = JobRecord(job_id=uuid.uuid4().hex[:12], kind=spec.kind)
        with self._lock:
            if spec.kind in MUTATING_JOBS and self._active_mutating is not None:
                record.status = "blocked"
                record.last_error = f"Mutating job already running: {self._active_mutating}"
                self._jobs[record.job_id] = record
                return record
            if spec.kind in MUTATING_JOBS:
                self._active_mutating = record.job_id
            self._jobs[record.job_id] = record

        thread = threading.Thread(
            target=self._run_in_thread,
            args=(record, spec),
            name=f"pstmigrate-job-{record.job_id}",
            daemon=True,
        )
        thread.start()
        return record

    def _run_in_thread(self, placeholder: JobRecord, spec: JobSpec) -> None:
        spec_record = JobRecord(job_id=placeholder.job_id, kind=spec.kind)

        def sync(update: JobRecord) -> None:
            spec_record.status = update.status
            spec_record.phase = update.phase
            spec_record.selected_rows = update.selected_rows
            spec_record.started_at = update.started_at
            spec_record.finished_at = update.finished_at
            spec_record.reports = update.reports
            spec_record.validation = update.validation
            spec_record.result = update.result
            spec_record.last_error = update.last_error
            with self._lock:
                self._jobs[placeholder.job_id] = spec_record

        final = run_job(spec, callbacks=[sync])
        final.job_id = placeholder.job_id
        with self._lock:
            self._jobs[placeholder.job_id] = final
            if spec.kind in MUTATING_JOBS and self._active_mutating == placeholder.job_id:
                self._active_mutating = None

    def get(self, job_id: str) -> JobRecord | None:
        with self._lock:
            return self._jobs.get(job_id)

    def latest(self, limit: int = 10) -> list[JobRecord]:
        with self._lock:
            return list(self._jobs.values())[-limit:]

