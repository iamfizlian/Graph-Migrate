"""End-of-migration completeness audit across every mailbox.

For each mailbox, computes:
  - total .eml files seen by the orchestrator (rows in `messages`)
  - done / skipped / failed row counts
  - unique distinct messages (by dedupe_key) actually uploaded
  - dedupe-skip count (rows whose key matches an uploaded done row -> safe)
  - orphan-skip count (rows whose key has NO done row -> potential loss)
  - failure count
  - completeness pct = unique_uploaded / unique_distinct_messages

Run from Graph-Migrate/:
  .\.venv\Scripts\python.exe _final_audit.py
"""
from __future__ import annotations

import sqlite3
from pathlib import Path

DB = Path(".pstmigrate-state") / "state.sqlite"


def main() -> None:
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row

    mailboxes = [
        r["target_mailbox"]
        for r in conn.execute(
            "SELECT DISTINCT target_mailbox FROM messages ORDER BY target_mailbox"
        ).fetchall()
    ]

    if not mailboxes:
        print("No mailboxes found in state.sqlite")
        return

    print(f"{'mailbox':<26} {'total':>7} {'done':>7} {'skip':>7} {'fail':>5} "
          f"{'uniq_done':>10} {'dedupe':>8} {'orphan':>7} {'cov%':>6}")
    print("-" * 110)

    grand = {"total": 0, "done": 0, "skipped": 0, "failed": 0,
             "uniq_done": 0, "dedupe_legit": 0, "orphans": 0}

    rows_out = []
    for mbx in mailboxes:
        c = conn.execute(
            """
            SELECT
              SUM(CASE WHEN status='done'    THEN 1 ELSE 0 END) AS done,
              SUM(CASE WHEN status='skipped' THEN 1 ELSE 0 END) AS skip,
              SUM(CASE WHEN status='failed'  THEN 1 ELSE 0 END) AS fail,
              COUNT(*) AS total
            FROM messages WHERE target_mailbox=?
            """,
            (mbx,),
        ).fetchone()
        total = c["total"] or 0
        done = c["done"] or 0
        skip = c["skip"] or 0
        fail = c["fail"] or 0

        uniq_done = conn.execute(
            "SELECT COUNT(DISTINCT dedupe_key) FROM messages WHERE target_mailbox=? AND status='done'",
            (mbx,),
        ).fetchone()[0]

        dedupe_legit = conn.execute(
            """
            SELECT COUNT(DISTINCT s.dedupe_key)
            FROM messages s
            WHERE s.target_mailbox=? AND s.status='skipped'
              AND EXISTS (
                SELECT 1 FROM messages d
                WHERE d.target_mailbox=s.target_mailbox
                  AND d.dedupe_key=s.dedupe_key
                  AND d.status='done'
              )
            """,
            (mbx,),
        ).fetchone()[0]

        uniq_skip = conn.execute(
            "SELECT COUNT(DISTINCT dedupe_key) FROM messages WHERE target_mailbox=? AND status='skipped'",
            (mbx,),
        ).fetchone()[0]
        orphan = uniq_skip - dedupe_legit

        # Total distinct logical messages = uniq_done + orphans + (failures we can't account for)
        # Coverage = uniq_done / (uniq_done + orphans).
        # Failures aren't "missing" if they ultimately succeeded somewhere, but
        # in this DB any 'failed' row is genuinely missing from destination.
        uniq_failed = conn.execute(
            """
            SELECT COUNT(DISTINCT dedupe_key) FROM messages
            WHERE target_mailbox=? AND status='failed'
              AND NOT EXISTS (
                SELECT 1 FROM messages d
                WHERE d.target_mailbox=messages.target_mailbox
                  AND d.dedupe_key=messages.dedupe_key
                  AND d.status='done'
              )
            """,
            (mbx,),
        ).fetchone()[0]

        denominator = uniq_done + orphan + uniq_failed
        coverage = (100.0 * uniq_done / denominator) if denominator else 100.0

        short = mbx.split("@")[0]
        rows_out.append({
            "name": short, "total": total, "done": done, "skip": skip,
            "fail": fail, "uniq_done": uniq_done, "dedupe_legit": dedupe_legit,
            "orphan": orphan, "uniq_failed": uniq_failed, "coverage": coverage,
        })

        for k in ("total", "done", "skipped", "failed", "uniq_done"):
            grand[k] = grand.get(k, 0)  # ensure exists
        grand["total"] += total
        grand["done"] += done
        grand["skipped"] += skip
        grand["failed"] += fail
        grand["uniq_done"] += uniq_done
        grand["dedupe_legit"] += dedupe_legit
        grand["orphans"] += orphan

    # Print rows; flag any row with orphans>0 or failures>0
    for r in rows_out:
        flag = ""
        if r["orphan"] > 0:
            flag += " !"
        if r["fail"] > 0:
            flag += " F"
        print(f"{r['name']:<26} {r['total']:>7} {r['done']:>7} {r['skip']:>7} "
              f"{r['fail']:>5} {r['uniq_done']:>10} {r['dedupe_legit']:>8} "
              f"{r['orphan']:>7} {r['coverage']:>5.2f}%{flag}")

    print("-" * 110)
    print(f"{'TOTALS':<26} {grand['total']:>7} {grand['done']:>7} {grand['skipped']:>7} "
          f"{grand['failed']:>5} {grand['uniq_done']:>10} {grand['dedupe_legit']:>8} "
          f"{grand['orphans']:>7}")
    print()

    print("Legend:")
    print("  total       = total .eml files extracted from PST")
    print("  done        = messages successfully uploaded (one row per source .eml)")
    print("  skip        = dedupe-skipped (multiple .eml files for same Message-ID)")
    print("  fail        = currently in 'failed' state (lost from destination)")
    print("  uniq_done   = DISTINCT messages by Message-ID actually in the mailbox")
    print("  dedupe      = unique dedupe-keys matched to a done sibling (legit dedupe)")
    print("  orphan      = unique dedupe-keys with NO done row (potentially missing)")
    print("  cov%        = unique_done / (unique_done + orphans + unique_failed)")
    print()
    print("Flags:  '!' = mailbox has orphan-skipped messages, "
          "'F' = mailbox has failed messages")

    if any(r["orphan"] > 0 or r["fail"] > 0 for r in rows_out):
        print()
        print("To inspect orphans/failures for a specific mailbox, edit _orphans.py")
        print("(or rerun _audit_failures.py for failure breakdown).")


if __name__ == "__main__":
    main()
