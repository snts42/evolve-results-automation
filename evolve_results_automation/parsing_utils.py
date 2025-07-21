import os
import re
from datetime import datetime
from .logging_utils import log
from .config import REPORTS_BASE

def make_report_folder_path(date_str: str) -> str:
    dt = datetime.strptime(date_str, "%d/%m/%Y")
    year = dt.strftime("%Y")
    month_day = dt.strftime("%m %d")
    folder = os.path.join(REPORTS_BASE, year, month_day)
    os.makedirs(folder, exist_ok=True)
    return folder

def report_filename(row):
    parts = [
        str(row["First name"]).strip(),
        str(row["Last name"]).strip(),
        str(row["Test Name"]).strip(),
        str(row["Result"]).strip()
    ]
    fname = " ".join(parts) + ".pdf"
    fname = "".join(c for c in fname if c not in r'\/:*?"<>|')
    return fname

def unique_row_hash(row):
    fields = [
        "Candidate ref.", "First name", "Last name", "Completed", "Test Name", "Result"
    ]
    return "|".join([str(row.get(f, "")).strip().lower() for f in fields])

def extract_pdf_filename_from_html(html):
    match = re.search(r'([a-f0-9\-]{36}\.pdf)', html)
    if match:
        return match.group(1)
    return None

def get_column_map():
    return {
        "Downloaded At": "Report URL",
        "Report Downloaded At": "Report Download",
        "E-Certificate Sent": "E-Certificate",
        "Certificate Issued": "Certificate"
    }