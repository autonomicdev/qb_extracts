"""
Runs qb_invoice_extract.py for the prior 30 days and drops the file on the network share.
"""

import subprocess
import sys
from datetime import date, timedelta

today = date.today()
from_date = (today - timedelta(days=30)).strftime("%Y-%m-%d")
to_date = today.strftime("%Y-%m-%d")

subprocess.run([
    sys.executable, "qb_invoice_extract.py",
    "--from-date", from_date,
    "--to-date",   to_date,
    "--output-dir", r"\\FS-01\Application Data\qb",
], check=True)
