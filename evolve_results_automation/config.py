import os
import sys
from datetime import datetime

if getattr(sys, 'frozen', False):
    # Running as a PyInstaller bundled .exe - use the exe's directory
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Running as a Python module (python -m evolve_results_automation)
    # __file__ is .../evolve_results_automation/config.py, so go up two levels
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

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
    # Core data (left)
    "Enrolment no.", "First name", "Last name", "Completed",
    "Test Name", "Result", "Centre Name",
    # Admin fills in (middle)
    "Result Sent", "Result Sent By",
    "E-Certificate sent", "E-Certificate By",
    "Certificate", "Certificate By",
    # Admin notes
    "Comments",
    # Technical/auto (right)
    "Percent", "Duration",
    "Scraping date/time", "PDF report save time",
    "Keycode", "Subject"
]

# Credentials stay at root level (shared across years)
ENCRYPTED_CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.enc")

# Evolve platform URLs
RESULTS_URL = "https://evolve.cityandguilds.com/#TestAdministration/Results"
DOCUMENT_STORE_URL = "https://evolve.cityandguilds.com/secureassess/CustomerData/Evolve/DocumentStore/"

def current_log_path():
    """Get log path for current run (uses current date's year)."""
    now = datetime.now()
    logs_base = get_logs_base_for_year(now.year)
    folder = os.path.join(logs_base, now.strftime("%m %d"))
    os.makedirs(folder, exist_ok=True)
    fname = now.strftime("log_%Y-%m-%d_%H-%M-%S.txt")
    return os.path.join(folder, fname)
