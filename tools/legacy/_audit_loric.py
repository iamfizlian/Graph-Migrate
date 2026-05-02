import sqlite3

c = sqlite3.connect(r".pstmigrate-state\state.sqlite")
cur = c.cursor()
cur.execute(
    "SELECT pst_path, target_mailbox, status, items_total, last_error, "
    "datetime(started_at,'unixepoch','localtime') AS started, "
    "datetime(finished_at,'unixepoch','localtime') AS finished "
    "FROM pst_runs WHERE target_mailbox LIKE 'loric%'"
)
cols = [d[0] for d in cur.description]
rows = cur.fetchall()
if not rows:
    print("(no rows for loric)")
for r in rows:
    print(dict(zip(cols, r)))
