import os
from datetime import datetime

CHROME_DRIVER_PATH = "chromedriver.exe"
ENCRYPTED_CREDENTIALS_FILE = "credentials.enc"
EXCEL_FILE = "exam_results.xlsx"
REPORTS_BASE = "reports"
LOGS_BASE = "logs"

def current_log_path():
    now = datetime.now()
    folder = os.path.join(LOGS_BASE, now.strftime("%Y"), now.strftime("%m %d"))
    os.makedirs(folder, exist_ok=True)
    fname = now.strftime("log_%Y-%m-%d_%H-%M-%S.txt")
    return os.path.join(folder, fname)

LOG_FILE = current_log_path()