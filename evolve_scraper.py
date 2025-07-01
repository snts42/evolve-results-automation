import os
import time
import json
import shutil
import re
import requests
import pandas as pd

from datetime import datetime
from typing import List, Dict, Set
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def extract_pdf_filename_from_html(html):
    # Search for .pdf files in the HTML
    match = re.search(r'([a-f0-9\-]{36}\.pdf)', html)
    if match:
        return match.group(1)
    return None

def download_pdf_with_cookies(driver, pdf_url, save_path):
    """Download PDF using session cookies from Selenium driver."""
    session = requests.Session()
    # Transfer cookies from Selenium to requests
    for cookie in driver.get_cookies():
        session.cookies.set(cookie['name'], cookie['value'])
    r = session.get(pdf_url, stream=True)
    if r.status_code == 200:
        with open(save_path, 'wb') as f:
            for chunk in r.iter_content(10240):
                f.write(chunk)
        log(f"PDF downloaded and saved to {save_path}")
        return True
    else:
        log(f"Failed to download PDF. Status: {r.status_code}")
        return False

# ---- CONFIGURATION ----
CHROME_DRIVER_PATH = "chromedriver.exe"
CREDENTIALS_FILE = "credentials.json"
EXCEL_FILE = "exam_results.xlsx"
REPORTS_BASE = "reports"
DOWNLOAD_DIR = os.path.join(os.getcwd(), REPORTS_BASE, "_downloads")  # Temporary download location

# ---- LOGGING ----
def log(msg: str):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"{timestamp} | {msg}")

# ---- EXCEL MANAGEMENT ----
def initialize_excel(filepath: str, columns: List[str]):
    if not os.path.exists(filepath):
        log(f"Excel file not found. Creating {filepath}")
        pd.DataFrame(columns=columns).to_excel(filepath, index=False)

def load_existing_results(filepath: str) -> pd.DataFrame:
    if not os.path.exists(filepath):
        return pd.DataFrame()
    return pd.read_excel(filepath, dtype=str)

def save_results(filepath: str, df: pd.DataFrame):
    df.to_excel(filepath, index=False)

# ---- CREDENTIALS ----
def load_credentials(json_file: str):
    with open(json_file, 'r') as f:
        creds = json.load(f)
    return creds['username'], creds['password']

# ---- SELENIUM SETUP ----
def start_driver(download_dir: str) -> webdriver.Chrome:
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--force-device-scale-factor=0.5")
    # Optionally set browser zoom factor, e.g., 80%: chrome_options.add_argument("--force-device-scale-factor=0.8")
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        #"plugins.always_open_pdf_externally": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    service = Service(executable_path=CHROME_DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

# ---- UTILS ----
def step(msg: str):
    log(f"STEP: {msg}")
    input("\nPress Enter to continue...\n")

def safe_find(driver, by, value, timeout=15):
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    except Exception:
        with open("debug_failed_to_find_element.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception(f"Element not found (timeout): {value}")

def wait_for_element(driver, by, value, timeout=15):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))

def make_report_folder_path(date_str: str) -> str:
    dt = datetime.strptime(date_str, "%d/%m/%Y")
    year = dt.strftime("%Y")
    month_day = dt.strftime("%m %d")
    folder = os.path.join(REPORTS_BASE, year, month_day)
    os.makedirs(folder, exist_ok=True)
    return folder

def report_filename(row: Dict) -> str:
    parts = [
        str(row["First name"]).strip(),
        str(row["Last name"]).strip(),
        str(row["Test Name"]).strip(),
        str(row["Result"]).strip()
    ]
    fname = " - ".join(parts) + ".pdf"
    fname = "".join(c for c in fname if c not in r'\/:*?"<>|')
    return fname

def unique_row_hash(row: Dict) -> str:
    fields = [
        "Candidate ref.", "First name", "Last name", "Completed", "Test Name", "Result"
    ]
    return "|".join([str(row.get(f, "")).strip().lower() for f in fields])

def wait_for_and_move_pdf(download_dir: str, target_dir: str, target_name: str, timeout=60) -> str:
    t0 = time.time()
    before = set(os.listdir(download_dir))
    pdf_path = None

    while time.time() - t0 < timeout:
        files = set(os.listdir(download_dir))
        new_files = files - before
        ready_files = [f for f in new_files if f.endswith('.pdf') and not f.endswith('.crdownload')]
        if ready_files:
            latest_pdf = max(
                (os.path.join(download_dir, f) for f in ready_files),
                key=os.path.getctime
            )
            pdf_path = latest_pdf
            break
        time.sleep(1)

    if not pdf_path:
        raise Exception("PDF did not appear in download directory within timeout.")

    os.makedirs(target_dir, exist_ok=True)
    dest_path = os.path.join(target_dir, target_name)
    shutil.move(pdf_path, dest_path)
    log(f"PDF moved to: {dest_path}")
    return dest_path

# ---- LOGIN ----
def login(driver, username: str, password: str):
    driver.get("https://evolve.cityandguilds.com/Login")
    step("Loaded login page. Check browser.")
    user_box = safe_find(driver, By.ID, "UserName")
    pass_box = safe_find(driver, By.ID, "Password")
    user_box.clear()
    user_box.send_keys(username)
    pass_box.clear()
    pass_box.send_keys(password)
    step("Credentials filled in. Ready to login?")
    login_btn = safe_find(driver, By.XPATH, "//input[@type='submit' and @value='Login']")
    login_btn.click()
    log("Login submitted. Wait for dashboard to load...")

# ---- NAVIGATION ----
def goto_results_tab(driver):
    step("Logged in. Wait for dashboard.")
    test_admin = safe_find(driver, By.XPATH, "//a[@data-id='TestAdministration']")
    test_admin.click()
    step("Clicked 'Test Administration'. Wait for results tab.")
    results_tab = safe_find(driver, By.XPATH, "//a[@href='#TestAdministration/Results']")
    results_tab.click()
    step("Results tab opened. Please check Results table visible.")

# ---- IFRAME MANAGEMENT ----
def switch_to_results_iframe(driver, timeout=15):
    try:
        iframe = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.ID, "TestAdministrationResultsFrame"))
        )
        driver.switch_to.frame(iframe)
        log("Switched to Results iframe.")
    except Exception:
        with open("debug_failed_to_find_iframe.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception("Could not find or switch to Results iframe.")

def switch_to_default(driver):
    driver.switch_to.default_content()
    log("Switched back to default content.")

def reset_and_refresh(driver):
    try:
        refresh_btn = safe_find(driver, By.XPATH, "//i[contains(@class,'dx-icon-refresh')]")
        refresh_btn.click()
        log("Refresh button clicked.")
        step("Refresh done, let table reload.")
    except Exception:
        log("Refresh button not found. Skipping.")

# ---- TABLE SCRAPING ----
def parse_results_table(driver, existing_hashes: Set[str]) -> List[Dict]:
    row_xpath = (
        "//div[contains(@class, 'dx-datagrid-rowsview')]"
        "//table[contains(@class, 'dx-datagrid-table')]"
        "/tbody/tr[contains(@class, 'dx-row') and not(contains(@class, 'dx-freespace-row'))]"
    )
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, row_xpath)))
    rows = driver.find_elements(By.XPATH, row_xpath)
    log(f"Found {len(rows)} data rows in the results table.")

    all_rows = []
    for idx, row in enumerate(rows):
        cells = row.find_elements(By.TAG_NAME, "td")
        if not any(cell.text.strip() for cell in cells) or len(cells) < 12:
            log(f"Skipping row {idx} - too few cells or all blank")
            continue
        data = {
            "Keycode": cells[1].text.strip(),
            "Candidate ref.": cells[2].text.strip(),
            "First name": cells[3].text.strip(),
            "Last name": cells[4].text.strip(),
            "Completed": cells[5].text.strip(),
            "Subject": cells[6].text.strip(),
            "Test Name": cells[7].text.strip(),
            "Result": cells[8].text.strip(),
            "Percent": cells[9].text.strip(),
            "Duration": cells[10].text.strip(),
            "Centre Name": cells[11].text.strip(),
            "Downloaded At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Report Downloaded At": "",
            "Result Sent": "",
            "E-Certificate Sent": "",
            "Certificate Issued": "",
            "Comments": ""
        }
        h = unique_row_hash(data)
        if h in existing_hashes:
            log(f"Skipping row {idx} - already exists in Excel")
            continue
        all_rows.append(data)
    return all_rows

# ---- TABLE ROW SELECTION ----
def select_table_row(driver, row: Dict) -> bool:
    row_xpath = (
        "//div[contains(@class, 'dx-datagrid-rowsview')]"
        "//table[contains(@class, 'dx-datagrid-table')]"
        "/tbody/tr[contains(@class, 'dx-row') and not(contains(@class, 'dx-freespace-row'))]"
    )
    table_rows = driver.find_elements(By.XPATH, row_xpath)
    for tr in table_rows:
        tds = tr.find_elements(By.TAG_NAME, "td")
        if not tds or len(tds) < 12:
            continue
        matches = all(
            tds[idx].text.strip() == str(row[col]).strip()
            for col, idx in [
                ("Candidate ref.", 2), ("First name", 3), ("Last name", 4),
                ("Completed", 5), ("Test Name", 7), ("Result", 8)
            ]
        )
        if matches:
            driver.execute_script("arguments[0].scrollIntoView(true);", tr)
            tr.click()
            log(f"Selected row in table: {row['First name']} {row['Last name']} ({row['Test Name']})")
            time.sleep(0.5)
            return True
    log("Could not find table row to select!")
    return False

def click_candidate_report_button(driver) -> bool:
    try:
        btn = safe_find(driver, By.ID, "button_candidatereport")
        btn.click()
        log("Clicked Candidate Report button.")
        time.sleep(1)
        return True
    except Exception as e:
        log(f"Failed to click Candidate Report button: {e}")
        return False

# ---- MAIN PROCESS ----
def main():
    columns = [
        "Keycode", "Candidate ref.", "First name", "Last name", "Completed",
        "Subject", "Test Name", "Result", "Percent", "Duration", "Centre Name",
        "Downloaded At", "Report Downloaded At", "PDF Direct Link",
        "Result Sent", "E-Certificate Sent", "Certificate Issued", "Comments"
    ]
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    initialize_excel(EXCEL_FILE, columns)
    existing_df = load_existing_results(EXCEL_FILE)
    existing_hashes = set(unique_row_hash(row) for _, row in existing_df.iterrows()) if not existing_df.empty else set()
    username, password = load_credentials(CREDENTIALS_FILE)

    driver = start_driver(DOWNLOAD_DIR)
    try:
        login(driver, username, password)
        goto_results_tab(driver)
        
        switch_to_results_iframe(driver)
        reset_and_refresh(driver)
        step("In Results iframe. All results should fit in one page. Then press Enter.")
        new_rows = parse_results_table(driver, existing_hashes)
        if new_rows:
            new_df = pd.DataFrame(new_rows)
            result_df = pd.concat([existing_df, new_df], ignore_index=True)
            save_results(EXCEL_FILE, result_df)
            log(f"Added {len(new_rows)} new rows to Excel.")
        else:
            result_df = existing_df.copy()
            log("No new rows found.")

        for idx, row in result_df[(result_df["PDF Direct Link"].isnull()) | (result_df["PDF Direct Link"] == "") | (result_df["PDF Direct Link"] == "NOT FOUND")].iterrows():
            try:
                log(f"--- Processing row idx={idx}: {row.to_dict()}")
                step(f"About to process: {row['First name']} {row['Last name']} ({row['Test Name']}).")

                # 1. Select correct table row (inside iframe)
                selected = select_table_row(driver, row)
                if not selected:
                    log(f"Could not select row for candidate {row['First name']} {row['Last name']}. Skipping.")
                    continue

                # 2. Leave iframe and click Candidate Report button
                ok = click_candidate_report_button(driver)
                if not ok:
                    log(f"Candidate Report button failed. Skipping.")
                    switch_to_results_iframe(driver)
                    continue

                step("Candidate report opened. Waiting for PDF filename in toolbar... Press Enter to continue if visible.")
                with open("debug_after_candidate_report.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                log("Dumped HTML after candidate report to debug_after_candidate_report.html")

                html = driver.page_source
                file_name = extract_pdf_filename_from_html(html)

                if file_name:
                    pdf_url = f"https://evolve.cityandguilds.com/secureassess/CustomerData/Evolve/DocumentStore/{file_name}"
                    log(f"PDF direct link: {pdf_url}")
                    # Save link to DataFrame
                    result_df.at[idx, "PDF Direct Link"] = pdf_url
                    save_results(EXCEL_FILE, result_df)
                else:
                    log("No PDF file name found! Skipping.")
                    result_df.at[idx, "PDF Direct Link"] = "NOT FOUND"
                    save_results(EXCEL_FILE, result_df)

                driver.switch_to.default_content()    
                results_tab = safe_find(driver, By.XPATH, "//a[@href='#TestAdministration/Results']")
                results_tab.click()
                step("Results tab opened. Please check Results table visible.")
                switch_to_results_iframe(driver)           

            except Exception as e:
                log(f"Error processing row idx={idx}: {e}")
                try:
                    goto_results_tab(driver)
                    switch_to_results_iframe(driver)
                except Exception as ex:
                    log(f"Failed to recover after error: {ex}")
                    break

    finally:
        result_df["Completed"] = pd.to_datetime(result_df["Completed"], dayfirst=True, errors='coerce')
        result_df = result_df.sort_values("Completed", ascending=True)
        result_df["Completed"] = result_df["Completed"].dt.strftime('%d/%m/%Y')
        save_results(EXCEL_FILE, result_df)

        driver.quit()
        log("Chrome closed.")

if __name__ == "__main__":
    main()