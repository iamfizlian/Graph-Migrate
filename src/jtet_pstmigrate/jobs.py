"""Web-safe job execution layer shared by frontends."""

from __future__ import annotations

import threading
import time
import uuid
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

from jtet_pstmigrate.auth import AppPool
from jtet_pstmigrate.config import MappingRow
from jtet_pstmigrate.entra_setup import (
    PermissionPreset,
    create_migration_apps_with_device_login,
    write_config_for_created_apps,
)
from jtet_pstmigrate.log import configure_logging
from jtet_pstmigrate.mapping import load_mapping
from jtet_pstmigrate.orchestrator import Orchestrator, RunReport
from jtet_pstmigrate.selection import SelectionFilters, select_mapping
from jtet_pstmigrate.services import load_config
from jtet_pstmigrate.state import StateStore
from jtet_pstmigrate.validation import ValidationReport, validate_environment

JobKind = Literal[
    "setup-entra",
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
JobStatus = Literal["queued", "running", "done", "failed", "blocked", "cancelled"]
JobCallback = Callable[["JobRecord"], None]
ProgressUpdate = dict[str, Any]


class JobCancelled(RuntimeError):
    """Raised internally when a user requests a cooperative job stop."""

MUTATING_JOBS: set[JobKind] = {
    "setup-entra",
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
    import_skipped_duplicates: bool = False
    tenant_id: str = ""
    app_count: int = 1
    secret_lifetime_days: int = 180
    permission_preset: PermissionPreset = "full"
    app_prefix: str = "pstmigrate"
    cancel_event: threading.Event | None = field(default=None, repr=False, compare=False)


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
    result: dict[str, Any] = field(default_factory=dict)
    last_error: str | None = None
    progress_total: int = 0
    progress_current: int = 0
    progress_uploaded: int = 0
    progress_skipped: int = 0
    progress_failed: int = 0
    progress_cancelled: int = 0
    activity: str = "Queued"
    events: list[dict[str, str]] = field(default_factory=list)
    _progress_lock: threading.Lock = field(default_factory=threading.Lock, repr=False)

    @property
    def elapsed_seconds(self) -> float:
        if self.started_at is None:
            return 0.0
        return (self.finished_at or time.time()) - self.started_at

    @property
    def item_counts(self) -> dict[str, int]:
        if self.progress_total or self.progress_current:
            return {
                "total": self.progress_total,
                "uploaded": self.progress_uploaded,
                "skipped": self.progress_skipped,
                "failed": self.progress_failed,
                "cancelled": self.progress_cancelled,
            }
        return {
            "total": sum(r.items_total for r in self.reports),
            "uploaded": sum(r.items_uploaded for r in self.reports),
            "skipped": sum(r.items_skipped for r in self.reports),
            "failed": sum(r.items_failed for r in self.reports),
        }

    @property
    def progress_percent(self) -> int:
        if self.progress_total <= 0:
            return 0
        return min(100, int((self.progress_current / self.progress_total) * 100))

    def record_event(self, message: str) -> None:
        with self._progress_lock:
            self.activity = message
            self.events.append({"elapsed": f"{self.elapsed_seconds:0.1f}s", "message": message})
            del self.events[:-30]

    def record_progress(self, update: ProgressUpdate) -> None:
        with self._progress_lock:
            if total := update.get("total"):
                self.progress_total = int(total)
            if total_delta := update.get("total_delta"):
                self.progress_total += int(total_delta)
            if increment := update.get("increment"):
                self.progress_current += int(increment)
            if outcome := update.get("outcome"):
                if outcome == "uploaded":
                    self.progress_uploaded += 1
                elif outcome == "skipped":
                    self.progress_skipped += 1
                elif outcome == "failed":
                    self.progress_failed += 1
                elif outcome == "cancelled":
                    self.progress_cancelled += 1
            if activity := update.get("activity"):
                self.activity = str(activity)
            if message := update.get("message") or update.get("activity"):
                self.events.append({"elapsed": f"{self.elapsed_seconds:0.1f}s", "message": str(message)})
                del self.events[:-30]


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
        record.record_event(phase)
        for callback in callbacks:
            callback(record)

    def progress(update: ProgressUpdate) -> None:
        record.record_progress(update)
        for callback in callbacks:
            callback(record)

    def check_cancelled() -> None:
        if spec.cancel_event is not None and spec.cancel_event.is_set():
            record.record_event("Cancellation requested; stopping after in-flight work finishes")
            raise JobCancelled("Cancelled by user.")

    try:
        if spec.kind in DESTRUCTIVE_JOBS and not spec.confirmed:
            record.status = "blocked"
            record.last_error = "Confirmation is required for destructive jobs."
            return record

        check_cancelled()
        record.status = "running"
        record.started_at = time.time()

        if spec.kind == "setup-entra":
            if spec.config_path is None:
                raise ValueError("A config path is required for Entra setup.")
            if not spec.tenant_id.strip():
                raise ValueError("Tenant ID/domain is required for Entra setup.")

            device_flow: dict[str, Any] = {}

            def capture_device_flow(flow: dict[str, Any]) -> None:
                device_flow.update(
                    {
                        "user_code": flow.get("user_code"),
                        "verification_uri": flow.get("verification_uri"),
                        "verification_uri_complete": flow.get("verification_uri_complete"),
                        "expires_in": flow.get("expires_in"),
                    }
                )
                record.result["device_flow"] = device_flow
                record.record_event(
                    "Open {uri} and enter code {code}".format(
                        uri=device_flow.get("verification_uri")
                        or device_flow.get("verification_uri_complete"),
                        code=device_flow.get("user_code"),
                    )
                )
                for callback in callbacks:
                    callback(record)

            def setup_event(message: str) -> None:
                record.record_event(message)
                for callback in callbacks:
                    callback(record)

            emit("waiting for admin sign-in")
            created_apps = create_migration_apps_with_device_login(
                tenant_id=spec.tenant_id,
                app_prefix=spec.app_prefix,
                app_count=spec.app_count,
                secret_lifetime_days=spec.secret_lifetime_days,
                permission_preset=spec.permission_preset,
                device_flow_callback=capture_device_flow,
                event_callback=setup_event,
            )
            check_cancelled()
            emit("writing config")
            write_config_for_created_apps(spec.config_path, created_apps)
            record.result = {
                "config_path": str(spec.config_path),
                "apps": [
                    {
                        "name": app.name,
                        "client_id": app.client_id,
                        "secret_expires_at": app.secret_expires_at,
                        "permissions": app.permissions,
                    }
                    for app in created_apps
                ],
            }
            record.status = "done"
            return record

        emit("loading config")
        cfg = load_config(spec.config_path)
        log_id = f"{spec.kind}_{time.strftime('%Y%m%d_%H%M%S')}"
        configure_logging(cfg.paths.log_dir, cfg.log_level, run_id=log_id)

        check_cancelled()
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
            check_cancelled()
            emit("validating")
            record.validation = validate_environment(cfg, rows)
            record.status = "done" if record.validation.ok else "failed"
            return record

        state = StateStore(cfg.paths.state_dir / "state.sqlite")
        pool = AppPool(cfg.apps)
        orch = Orchestrator(
            cfg,
            state,
            pool,
            import_skipped_duplicates=spec.import_skipped_duplicates,
            progress_callback=progress,
            cancel_event=spec.cancel_event,
        )

        if spec.kind == "import-mail":
            check_cancelled()
            emit("importing mail")
            record.reports = orch.run(rows)
            check_cancelled()
        elif spec.kind == "import-calendar":
            check_cancelled()
            emit("importing calendar")
            record.reports = orch.run_calendar(rows)
            check_cancelled()
        elif spec.kind == "import-contacts":
            check_cancelled()
            emit("importing contacts")
            record.reports = orch.run_contacts(rows)
            check_cancelled()
        elif spec.kind == "import-all":
            check_cancelled()
            emit("importing mail")
            record.reports.extend(orch.run(rows))
            check_cancelled()
            emit("importing calendar")
            record.reports.extend(orch.run_calendar(rows))
            check_cancelled()
            emit("importing contacts")
            record.reports.extend(orch.run_contacts(rows))
            check_cancelled()
        elif spec.kind == "purge-calendar":
            check_cancelled()
            emit("purging calendar")
            deleted, missing, errors = orch.purge_calendar(rows)
            record.result = {"deleted": deleted, "missing": missing, "errors": errors}
        elif spec.kind == "purge-contacts":
            check_cancelled()
            emit("purging contacts")
            deleted, missing, errors = orch.purge_contacts(rows)
            record.result = {"deleted": deleted, "missing": missing, "errors": errors}
        elif spec.kind == "purge-mail":
            check_cancelled()
            emit("purging mail")
            deleted, missing, errors = orch.purge_mail(rows)
            record.result = {"deleted": deleted, "missing": missing, "errors": errors}
        elif spec.kind == "reset-state":
            check_cancelled()
            emit("resetting state")
            record.result = _reset_state(state, rows)

        failed_reports = sum(1 for r in record.reports if r.status != "done" or r.items_failed)
        result_errors = record.result.get("errors", 0)
        if spec.cancel_event is not None and spec.cancel_event.is_set():
            record.status = "cancelled"
            record.last_error = "Cancelled by user."
        else:
            record.status = "failed" if failed_reports or result_errors else "done"
        return record
    except JobCancelled as e:
        record.status = "cancelled"
        record.last_error = str(e)
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
        self._cancel_events: dict[str, threading.Event] = {}

    def submit(self, spec: JobSpec) -> JobRecord:
        record = JobRecord(job_id=uuid.uuid4().hex[:12], kind=spec.kind)
        if spec.kind in DESTRUCTIVE_JOBS and not spec.confirmed:
            record.status = "blocked"
            record.last_error = "Confirmation is required for destructive jobs."
            with self._lock:
                self._jobs[record.job_id] = record
            return record

        cancel_event = threading.Event()
        spec.cancel_event = cancel_event
        with self._lock:
            if spec.kind in MUTATING_JOBS and self._active_mutating is not None:
                record.status = "blocked"
                record.last_error = f"Mutating job already running: {self._active_mutating}"
                self._jobs[record.job_id] = record
                return record
            if spec.kind in MUTATING_JOBS:
                self._active_mutating = record.job_id
            self._jobs[record.job_id] = record
            self._cancel_events[record.job_id] = cancel_event

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
            spec_record.progress_total = update.progress_total
            spec_record.progress_current = update.progress_current
            spec_record.progress_uploaded = update.progress_uploaded
            spec_record.progress_skipped = update.progress_skipped
            spec_record.progress_failed = update.progress_failed
            spec_record.progress_cancelled = update.progress_cancelled
            spec_record.activity = update.activity
            spec_record.events = list(update.events)
            with self._lock:
                self._jobs[placeholder.job_id] = spec_record

        final = run_job(spec, callbacks=[sync])
        final.job_id = placeholder.job_id
        with self._lock:
            self._jobs[placeholder.job_id] = final
            self._cancel_events.pop(placeholder.job_id, None)
            if spec.kind in MUTATING_JOBS and self._active_mutating == placeholder.job_id:
                self._active_mutating = None

    def cancel(self, job_id: str) -> bool:
        with self._lock:
            record = self._jobs.get(job_id)
            cancel_event = self._cancel_events.get(job_id)
            if record is None or cancel_event is None or record.status not in {"queued", "running"}:
                return False
            cancel_event.set()
            record.phase = "cancelling"
            record.record_event("Stop requested from web UI; waiting for in-flight work to finish")
            self._jobs[job_id] = record
            return True

    def get(self, job_id: str) -> JobRecord | None:
        with self._lock:
            return self._jobs.get(job_id)

    def latest(self, limit: int = 10) -> list[JobRecord]:
        with self._lock:
            return list(self._jobs.values())[-limit:]
