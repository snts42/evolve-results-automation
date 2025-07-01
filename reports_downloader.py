import os
import time
import pandas as pd
import requests
from datetime import datetime

EXCEL_FILE = "exam_results.xlsx"
REPORTS_BASE = "reports"

def log(msg):
    print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}")

def make_report_folder_path(date_str):
    dt = datetime.strptime(date_str, "%d/%m/%Y")
    year = dt.strftime("%Y")
    month_day = dt.strftime("%m %d")
    folder = os.path.join(REPORTS_BASE, year, month_day)
    os.makedirs(folder, exist_ok=True)
    return folder

def report_filename(row):
    first_last = f"{str(row['First name']).strip()} {str(row['Last name']).strip()}"
    test = str(row["Test Name"]).strip()
    result = str(row["Result"]).strip()
    fname = f"{first_last} - {test} - {result}.pdf"
    fname = "".join(c for c in fname if c not in r'\/:*?"<>|')
    return fname

def is_valid_link(link):
    return (
        isinstance(link, str) and
        link.strip() != "" and
        link.strip().lower() != "not found" and
        link.strip().startswith("http")
    )

def download_pdf(url, save_path):
    try:
        r = requests.get(url, stream=True, timeout=30)
        if r.status_code == 200:
            with open(save_path, "wb") as f:
                for chunk in r.iter_content(10240):
                    f.write(chunk)
            return True
        else:
            log(f"Download failed: {url} [status {r.status_code}]")
            return False
    except Exception as e:
        log(f"Exception downloading {url}: {e}")
        return False

def main():
    if not os.path.exists(EXCEL_FILE):
        log(f"Excel file not found: {EXCEL_FILE}")
        return

    df = pd.read_excel(EXCEL_FILE, dtype=str)

    # Only process rows with missing 'Report Downloaded At' and a valid PDF link
    mask = (
        (df["Report Downloaded At"].isnull() | (df["Report Downloaded At"] == "")) &
        (df["PDF Direct Link"].apply(is_valid_link))
    )
    to_process = df[mask]

    if to_process.empty:
        log("No new reports downloaded. (All links already marked as downloaded, or links missing)")
        return

    for idx, row in to_process.iterrows():
        url = row["PDF Direct Link"]
        completed = row["Completed"]
        fname = report_filename(row)
        folder = make_report_folder_path(completed)
        save_path = os.path.join(folder, fname)

        log(f"Downloading PDF for: {row['First name']} {row['Last name']} ({row['Test Name']})")
        log(f" -> URL: {url}")
        log(f" -> Save as: {save_path}")

        if download_pdf(url, save_path):
            df.at[idx, "Report Downloaded At"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        else:
            log("Download failed, will try again next time.")

        time.sleep(1)  # polite delay

    # Optional: sort by Completed date before saving
    try:
        df["Completed_dt"] = pd.to_datetime(df["Completed"], format="%d/%m/%Y")
        df.sort_values(by="Completed_dt", inplace=True)
        df.drop(columns=["Completed_dt"], inplace=True)
    except Exception:
        pass

    df.to_excel(EXCEL_FILE, index=False)
    log("Done. Reports downloaded and Excel updated.")

if __name__ == "__main__":
    main()