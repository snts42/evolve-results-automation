import os
import re
import logging
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

from .config import get_reports_base_for_year

def make_report_folder_path(date_str: str) -> str:
    """Create a dated folder path for reports based on exam completion date."""
    dt = datetime.strptime(date_str, "%d/%m/%Y")
    year = dt.year
    month_day = dt.strftime("%m %d")
    reports_base = get_reports_base_for_year(year)
    folder = os.path.join(reports_base, month_day)
    os.makedirs(folder, exist_ok=True)
    return folder

def report_filename(row):
    """Generate a sanitized PDF filename from row data."""
    parts = [
        str(row["First name"]).strip(),
        str(row["Last name"]).strip(),
        str(row["Test Name"]).strip(),
        str(row["Result"]).strip()
    ]
    fname = " ".join(parts) + ".pdf"
    fname = "".join(c for c in fname if c not in r'\/:*?"<>|')
    if len(fname) > 200:
        fname = fname[:196] + ".pdf"
    return fname

def unique_row_hash(row):
    """Generate a unique hash for a result row for deduplication."""
    fields = [
        "Enrolment no.", "First name", "Last name", "Completed", "Test Name", "Result"
    ]
    return "|".join([str(row.get(f, "")).strip().lower() for f in fields])

def extract_pdf_filename_from_html(html):
    """Extract the PDF filename from HTML content."""
    match = re.search(r'([a-f0-9\-]{36}\.pdf)', html, re.IGNORECASE)
    if match:
        return match.group(1)
    return None

def download_pdf(pdf_url, row, completed):
    """Download a PDF report to the appropriate dated folder.
    Returns True if the file was downloaded or already exists on disk."""
    target_dir = make_report_folder_path(completed)
    target_name = report_filename(row)
    save_path = os.path.join(target_dir, target_name)

    if os.path.exists(save_path) and os.path.getsize(save_path) > 0:
        logging.info("PDF already on disk, updating timestamp")
        return True

    try:
        resp = urlopen(Request(pdf_url), timeout=30)
        try:
            with open(save_path, 'wb') as f:
                while True:
                    chunk = resp.read(10240)
                    if not chunk:
                        break
                    f.write(chunk)
            logging.info("  PDF saved")
            return True
        finally:
            resp.close()
    except HTTPError as e:
        logging.warning(f"Failed to download PDF, status: {e.code}")
        return False
    except (URLError, OSError) as e:
        logging.warning(f"Failed to download PDF: {e}")
        return False
