"""State reporting helpers for CLI and web dashboards."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from jtet_pstmigrate.state import StateStore


@dataclass(slots=True)
class DashboardTotals:
    runs: int = 0
    done: int = 0
    failed: int = 0
    skipped: int = 0
    queued: int = 0


class StateQueries:
    """Read-only query facade over the migration state database."""

    def __init__(self, state: StateStore):
        self._state = state

    def dashboard_totals(self) -> DashboardTotals:
        totals = DashboardTotals()
        runs = self._state.all_runs()
        totals.runs = len(runs)
        for run in runs:
            counts = self._state.counts_for_run(run["target_mailbox"], run["pst_path"])
            totals.done += counts["done"]
            totals.failed += counts["failed"]
            totals.skipped += counts["skipped"]
            totals.queued += counts["queued"]
        return totals

    def run_rows(self) -> list[dict[str, str | int]]:
        rows: list[dict[str, str | int]] = []
        for run in self._state.all_runs():
            counts = self._state.counts_for_run(run["target_mailbox"], run["pst_path"])
            rows.append(
                {
                    "mailbox": run["target_mailbox"],
                    "pst": Path(run["pst_path"]).name,
                    "status": run["status"],
                    "total": run["items_total"],
                    "done": counts["done"],
                    "failed": counts["failed"],
                    "skipped": counts["skipped"],
                    "last_error": (run["last_error"] or "")[:160],
                }
            )
        return rows

    def app_breakdown(self) -> dict[str, dict[str, int]]:
        return self._state.app_breakdown()

