import os
import sys
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

# Year-based organization (by exam completion date)
def get_year_folder(year):
    """Get the folder for a specific year."""
    year_folder = os.path.join(BASE_DIR, str(year))
    os.makedirs(year_folder, exist_ok=True)
    return year_folder

def get_excel_file_for_year(year):
    """Get the Excel file path for a specific year."""
    year_folder = get_year_folder(year)
    return os.path.join(year_folder, "exam_results.xlsx")

def get_reports_base_for_year(year):
    """Get the reports base folder for a specific year."""
    year_folder = get_year_folder(year)
    return os.path.join(year_folder, "reports")

def get_logs_base_for_year(year):
    """Get the logs base folder for a specific year."""
    year_folder = get_year_folder(year)
    return os.path.join(year_folder, "logs")

COLUMNS = [
    "Candidate ref.", "First name", "Last name", "Completed",
    "Test Name", "Result", "Percent", "Duration", "Centre Name",
    "Report URL", "Report Download", "Result Sent", "Result Sent By",
    "E-Certificate", "E-Certificate By", "Certificate", "Certificate By",
    "Comments", "Keycode", "Subject", "PDF Direct Link"
]

# Credentials stay at root level (shared across years)
ENCRYPTED_CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.enc")

# Default paths (use current year for logs)
CURRENT_YEAR = datetime.now().year
LOGS_BASE = get_logs_base_for_year(CURRENT_YEAR)

# Excel and reports are determined dynamically based on exam date
# Use get_excel_file_for_year(year) and get_reports_base_for_year(year)

def current_log_path():
    """Get log path for current run (uses current date's year)."""
    now = datetime.now()
    logs_base = get_logs_base_for_year(now.year)
    folder = os.path.join(logs_base, now.strftime("%m %d"))
    os.makedirs(folder, exist_ok=True)
    fname = now.strftime("log_%Y-%m-%d_%H-%M-%S.txt")
    return os.path.join(folder, fname)

LOG_FILE = current_log_path()