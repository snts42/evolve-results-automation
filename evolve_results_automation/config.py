import os
import sys
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

COLUMNS = [
    "Candidate ref.", "First name", "Last name", "Completed",
    "Test Name", "Result", "Percent", "Duration", "Centre Name",
    "Report URL", "Report Download", "Result Sent", "Result Sent By",
    "E-Certificate", "E-Certificate By", "Certificate", "Certificate By",
    "Comments", "Keycode", "Subject", "PDF Direct Link"
]

ENCRYPTED_CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.enc")
EXCEL_FILE = os.path.join(BASE_DIR, "exam_results.xlsx")
REPORTS_BASE = os.path.join(BASE_DIR, "reports")
LOGS_BASE = os.path.join(BASE_DIR, "logs")

def current_log_path():
    now = datetime.now()
    folder = os.path.join(LOGS_BASE, now.strftime("%Y"), now.strftime("%m %d"))
    os.makedirs(folder, exist_ok=True)
    fname = now.strftime("log_%Y-%m-%d_%H-%M-%S.txt")
    return os.path.join(folder, fname)

LOG_FILE = current_log_path()