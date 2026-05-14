"""Fetch excel_paste text for a single dataset name, write to stdout."""
import sys, sqlite3
DB = r'D:\000. MyWorks\002. DB\process-review.db'
name = sys.argv[1]
con = sqlite3.connect(DB)
con.execute("PRAGMA busy_timeout = 30000")
row = con.execute(
    "SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'",
    (name,)).fetchone()
con.close()
if not row or not row[0]:
    print("__NO_PASTE__")
else:
    sys.stdout.buffer.write(row[0].encode("utf-8"))
