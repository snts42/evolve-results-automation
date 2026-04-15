import os
import sys
import glob
import json
import logging
from datetime import datetime

APP_VER = "v1.3.2"

if getattr(sys, 'frozen', False):
    # Running as a PyInstaller bundled .exe - use the exe's directory
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Running as a Python module (python -m evolve_results_automation)
    # __file__ is .../evolve_results_automation/config.py, so go up two levels
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Year-based organization (by exam completion date)
def _year_folder_path(year):
    """Pure path lookup - no side effects."""
    return os.path.join(BASE_DIR, str(year))

def get_excel_file_for_year(year):
    """Get the Excel file path for a specific year (no mkdir)."""
    return os.path.join(_year_folder_path(year), f"exam_results_{year}.xlsx")

def get_reports_base_for_year(year):
    """Get the reports base folder for a specific year (no mkdir)."""
    return os.path.join(_year_folder_path(year), "reports")

def get_logs_base_for_year(year):
    """Get the logs base folder for a specific year (no mkdir)."""
    return os.path.join(_year_folder_path(year), "logs")

def list_year_folders():
    """Return sorted list of 4-digit year folder names under BASE_DIR (descending)."""
    return sorted(
        [os.path.basename(d)
         for d in glob.glob(os.path.join(BASE_DIR, "[12][0-9][0-9][0-9]"))
         if os.path.isdir(d)],
        reverse=True)


def list_year_excel_files():
    """Return list of (year_str, filepath) for all year Excel files."""
    result = []
    for d in list_year_folders():
        p = get_excel_file_for_year(d)
        if os.path.isfile(p):
            result.append((d, p))
    return result


COLUMNS = [
    # Core data (left)
    "Enrolment no.", "First name", "Last name", "Completed",
    "Test Name", "Result", "Percent", "Centre Name",
    # Admin fills in (middle)
    "Result Sent", "Result Sent By",
    "E-Certificate sent", "E-Certificate By",
    "Certificate", "Certificate By",
    # Admin notes
    "Comments",
    # Technical/auto (right)
    "Duration",
    "Scraping date/time", "PDF report save time",
    "Keycode", "Subject"
]

# Root-level files (shared across years)
ENCRYPTED_CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.enc")
ANALYTICS_FILE = os.path.join(BASE_DIR, "analytics.xlsx")

# User settings (persisted across sessions)
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")

_DEFAULT_SETTINGS = {
    "show_browser": False,
    "notifications": True,
    "schedule_enabled": False,
    "schedule_time": "",
    "start_with_windows": False,
    "minimize_to_tray": False,
}

def load_settings():
    """Load user settings from disk, returning defaults for missing keys."""
    settings = dict(_DEFAULT_SETTINGS)
    try:
        if os.path.isfile(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            if isinstance(saved, dict):
                settings.update(saved)
    except Exception as e:
        logging.debug(f"Could not load settings: {e}")
    return settings

def save_settings(settings: dict):
    """Persist user settings to disk."""
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except Exception as e:
        logging.debug(f"Could not save settings: {e}")

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
