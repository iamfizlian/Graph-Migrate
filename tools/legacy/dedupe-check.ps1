@'
import sqlite3
mbx = input("Enter mailbox to check (e.g., user@tenant.onmicrosoft.com): ").strip()
c = sqlite3.connect(r".pstmigrate-state\state.sqlite")
overlap = c.execute("""
    SELECT COUNT(DISTINCT s.dedupe_key)
    FROM messages s
    WHERE s.target_mailbox=? AND s.status='skipped'
      AND EXISTS (
        SELECT 1 FROM messages d
        WHERE d.target_mailbox=s.target_mailbox
          AND d.dedupe_key=s.dedupe_key
          AND d.status='done'
      )
""", (mbx,)).fetchone()[0]
unique_skipped = c.execute("SELECT COUNT(DISTINCT dedupe_key) FROM messages WHERE target_mailbox=? AND status='skipped'", (mbx,)).fetchone()[0]
orphan = unique_skipped - overlap
print(f"\nMailbox: {mbx}")
print(f"unique skipped dedupe_keys      : {unique_skipped}")
print(f"  ...also have a done row       : {overlap}  (legit dedupes)")
print(f"  ...have NO done row           : {orphan}   (orphan skips - investigate)")
print()
print("Sample orphan skipped rows (unique dedupe_keys with NO done row):")
rows = c.execute("""
    SELECT source_path, dedupe_key, last_error
    FROM messages
    WHERE target_mailbox=? AND status='skipped'
      AND NOT EXISTS (
        SELECT 1 FROM messages d
        WHERE d.target_mailbox=messages.target_mailbox
          AND d.dedupe_key=messages.dedupe_key
          AND d.status='done'
      )
    LIMIT 5
""", (mbx,)).fetchall()
for r in rows:
    print(f"  src={r[0][-60:]}  key={r[1][:50]}  err={r[2]}")
'@ | Set-Content -Path "_dedupe_check.py" -Encoding utf8
.\.venv\Scripts\python.exe _dedupe_check.py