"""Quick health view: per-mailbox progress *for the current run only*.

Usage (from Graph-Migrate/):
  .\.venv\Scripts\python.exe _progress.py

Shows only deltas since each PST run's `started_at`, so legacy `failed`
rows from previous broken runs do not pollute the picture.
"""
from __future__ import annotations

import sqlite3
from pathlib import Path

DB = Path(".pstmigrate-state") / "state.sqlite"


def main() -> None:
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row

    # newest pst_run per (mailbox, pst_path)
    runs = conn.execute(
        """
        SELECT target_mailbox, pst_path, status, started_at, finished_at, items_total
        FROM pst_runs
        WHERE (target_mailbox, started_at) IN (
            SELECT target_mailbox, MAX(started_at) FROM pst_runs GROUP BY target_mailbox
        )
        ORDER BY target_mailbox
        """
    ).fetchall()

    print(f"{'mailbox':<14} {'state':<11} {'done':>8} {'fail':>8} {'skip':>6} "
          f"{'pending':>8} {'total':>8}  started_at")
    print("-" * 100)

    for r in runs:
        mbx = r["target_mailbox"].split("@")[0]

        # message rows touched DURING this run
        # (rows the orchestrator inserted/updated for this pst_path)
        counts = conn.execute(
            """
            SELECT
              SUM(CASE WHEN status='done'    THEN 1 ELSE 0 END) AS d,
              SUM(CASE WHEN status='failed'  THEN 1 ELSE 0 END) AS f,
              SUM(CASE WHEN status='skipped' THEN 1 ELSE 0 END) AS s,
              SUM(CASE WHEN status='pending' THEN 1 ELSE 0 END) AS p,
              COUNT(*) AS t
            FROM messages
            WHERE target_mailbox=? AND pst_path=?
            """,
            (r["target_mailbox"], r["pst_path"]),
        ).fetchone()

        d = counts["d"] or 0
        f = counts["f"] or 0
        s = counts["s"] or 0
        p = counts["p"] or 0
        t = counts["t"] or 0

        print(f"{mbx:<14} {r['status']:<11} {d:>8} {f:>8} {s:>6} {p:>8} {t:>8}  "
              f"{r['started_at']}")

    print()
    # global progress against current run
    g = conn.execute(
        """
        SELECT
          SUM(CASE WHEN status='done'    THEN 1 ELSE 0 END) AS d,
          SUM(CASE WHEN status='failed'  THEN 1 ELSE 0 END) AS f,
          SUM(CASE WHEN status='skipped' THEN 1 ELSE 0 END) AS s,
          COUNT(*) AS t
        FROM messages
        """,
    ).fetchone()
    d, f, s, t = g["d"] or 0, g["f"] or 0, g["s"] or 0, g["t"] or 0
    print(f"GLOBAL: done={d}  failed={f}  skipped={s}  total_seen={t}")
    if t:
        pct = 100 * (d + s) / t
        print(f"        {pct:5.1f}% complete (done+skipped)")


if __name__ == "__main__":
    main()
