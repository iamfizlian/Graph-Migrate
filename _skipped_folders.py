"""Show what folders the *skipped* messages came from, per mailbox.

Edit MAILBOX below or pass it on the command line.

Output columns:
  count_skipped      - how many .eml files in this folder were skipped
  also_done_in_folder - how many of those skips have a sibling 'done' row
                        in the SAME folder (vs. uploaded only into a
                        different folder)
  folder             - relative folder path within the PST
"""
from __future__ import annotations

import sqlite3
import sys
from pathlib import PurePath

DB = ".pstmigrate-state\\state.sqlite"

DEFAULT_MAILBOX = "johnd@jteatono365.onmicrosoft.com"
TOP_N = 30


def folder_of(source_path: str) -> str:
    p = PurePath(source_path)
    parts = list(p.parts)
    # Trim everything up to and including ".pstmigrate-work" + the
    # extraction-root dir, leaving just the in-PST folder hierarchy.
    try:
        i = parts.index(".pstmigrate-work")
        # parts[i+1] = extraction subdir (mirrors PST stem)
        # parts[i+2] = pst display name dir produced by readpst
        # parts[i+3:-1] = real folder path inside the PST
        rel = parts[i + 3 : -1]
    except ValueError:
        rel = parts[:-1]
    return "/".join(rel) if rel else "(root)"


def main() -> None:
    mbx = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_MAILBOX
    conn = sqlite3.connect(DB)
    print(f"Mailbox: {mbx}")
    print()

    rows = conn.execute(
        """
        SELECT source_path, dedupe_key
        FROM messages
        WHERE target_mailbox=? AND status='skipped'
        """,
        (mbx,),
    ).fetchall()
    if not rows:
        print("No skipped rows for this mailbox.")
        return

    # For each skipped row, find a 'done' sibling in the same folder
    folder_skipped: dict[str, int] = {}
    folder_skipped_with_done_sibling: dict[str, int] = {}
    by_dedupe_done_folders: dict[str, set[str]] = {}

    # Pre-fetch all done rows for this mailbox (one query, faster than per-row)
    done_rows = conn.execute(
        "SELECT source_path, dedupe_key FROM messages WHERE target_mailbox=? AND status='done'",
        (mbx,),
    ).fetchall()
    for sp, dk in done_rows:
        by_dedupe_done_folders.setdefault(dk, set()).add(folder_of(sp))

    for sp, dk in rows:
        f = folder_of(sp)
        folder_skipped[f] = folder_skipped.get(f, 0) + 1
        if f in by_dedupe_done_folders.get(dk, set()):
            folder_skipped_with_done_sibling[f] = (
                folder_skipped_with_done_sibling.get(f, 0) + 1
            )

    items = sorted(folder_skipped.items(), key=lambda kv: kv[1], reverse=True)
    total = sum(folder_skipped.values())
    print(f"{'count':>7} {'with-done-sibling':>20}  folder")
    print("-" * 100)
    for folder, n in items[:TOP_N]:
        sib = folder_skipped_with_done_sibling.get(folder, 0)
        print(f"{n:>7} {sib:>20}  {folder}")
    if len(items) > TOP_N:
        print(f"... and {len(items) - TOP_N} more folders")
    print()
    print(f"TOTAL skipped: {total}")
    sib_total = sum(folder_skipped_with_done_sibling.values())
    print(f"  of which the SAME folder also has a 'done' copy: {sib_total}")
    print(f"  remaining (skipped here, done in a DIFFERENT folder): {total - sib_total}")


if __name__ == "__main__":
    main()
