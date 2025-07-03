import os
import time
import json
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
from openpyxl import load_workbook

# ---- CONFIGURATION ----
CHROME_DRIVER_PATH = "chromedriver.exe"
CREDENTIALS_FILE = "credentials.json"
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

def log(msg: str):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    out = f"{timestamp} | {msg}"
    print(out)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(out + "\n")

def initialize_excel(filepath: str, columns: List[str]):
    if not os.path.exists(filepath):
        log(f"Excel file not found. Creating {filepath}")
        df = pd.DataFrame(columns=columns)
        df.to_excel(filepath, index=False)
        autofilter_and_autofit(filepath)

def load_existing_results(filepath: str) -> pd.DataFrame:
    if not os.path.exists(filepath):
        return pd.DataFrame()
    return pd.read_excel(filepath, dtype=str)

def save_results(filepath: str, df: pd.DataFrame):
    df.to_excel(filepath, index=False)
    autofilter_and_autofit(filepath)

def autofilter_and_autofit(filepath: str):
    """Adds autofilter to all columns and adjusts width to fit contents."""
    wb = load_workbook(filepath)
    ws = wb.active
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                cell_len = len(str(cell.value)) if cell.value is not None else 0
                if cell_len > max_length:
                    max_length = cell_len
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = max(10, min(max_length + 2, 60))
    wb.save(filepath)

def load_credentials(json_file: str):
    with open(json_file, 'r') as f:
        creds = json.load(f)
    if isinstance(creds, dict):
        creds = [creds]
    return creds

def start_driver() -> webdriver.Chrome:
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--force-device-scale-factor=0.5")
    service = Service(executable_path=CHROME_DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def safe_find(driver, by, value, timeout=15):
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    except Exception:
        with open("debug_failed_to_find_element.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception(f"Element not found (timeout): {value}")

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
    fname = " ".join(parts) + ".pdf"
    fname = "".join(c for c in fname if c not in r'\/:*?"<>|')
    return fname

def unique_row_hash(row: Dict) -> str:
    fields = [
        "Candidate ref.", "First name", "Last name", "Completed", "Test Name", "Result"
    ]
    return "|".join([str(row.get(f, "")).strip().lower() for f in fields])

def extract_pdf_filename_from_html(html):
    match = re.search(r'([a-f0-9\-]{36}\.pdf)', html)
    if match:
        return match.group(1)
    return None

def login(driver, username: str, password: str):
    driver.get("https://evolve.cityandguilds.com/Login")
    time.sleep(6)
    user_box = safe_find(driver, By.ID, "UserName")
    pass_box = safe_find(driver, By.ID, "Password")
    user_box.clear()
    user_box.send_keys(username)
    pass_box.clear()
    pass_box.send_keys(password)
    login_btn = safe_find(driver, By.XPATH, "//input[@type='submit' and @value='Login']")
    login_btn.click()
    time.sleep(6)
    log("Login submitted. Wait for dashboard to load...")

def goto_results_tab(driver):
    driver.get("https://evolve.cityandguilds.com/#TestAdministration/Results")
    time.sleep(6)
    test_admin = safe_find(driver, By.XPATH, "//a[@data-id='TestAdministration']")
    test_admin.click()
    time.sleep(6)
    results_tab = safe_find(driver, By.XPATH, "//a[@href='#TestAdministration/Results']")
    results_tab.click()
    time.sleep(6)
    log("Results tab opened. Please check Results table visible.")

def switch_to_results_iframe(driver, timeout=15):
    time.sleep(6)
    iframe = safe_find(driver, By.ID, "TestAdministrationResultsFrame")
    driver.switch_to.frame(iframe)
    time.sleep(6)
    log("Switched to Results iframe.")

def switch_to_default(driver):
    driver.switch_to.default_content()
    time.sleep(1)
    log("Switched back to default content.")

def reset_and_refresh(driver):
    refresh_btn = safe_find(driver, By.XPATH, "//i[contains(@class,'dx-icon-refresh')]")
    refresh_btn.click()
    time.sleep(6)
    log("Refresh button clicked. Wait for table reload.")

def parse_results_table(driver, existing_hashes: Set[str], col_map: Dict) -> List[Dict]:
    time.sleep(6)
    row_xpath = (
        "//div[contains(@class, 'dx-datagrid-rowsview')]"
        "//table[contains(@class, 'dx-datagrid-table')]"
        "/tbody/tr[contains(@class, 'dx-row') and not(contains(@class, 'dx-freespace-row'))]"
    )
    rows = driver.find_elements(By.XPATH, row_xpath)
    log(f"Found {len(rows)} data rows in the results table.")
    all_rows = []
    for idx, row in enumerate(rows):
        cells = row.find_elements(By.TAG_NAME, "td")
        if not any(cell.text.strip() for cell in cells) or len(cells) < 12:
            continue
        data = {
            col_map.get("Keycode", "Keycode"): cells[1].text.strip(),
            col_map.get("Candidate ref.", "Candidate ref."): cells[2].text.strip(),
            col_map.get("First name", "First name"): cells[3].text.strip(),
            col_map.get("Last name", "Last name"): cells[4].text.strip(),
            col_map.get("Completed", "Completed"): cells[5].text.strip(),
            col_map.get("Subject", "Subject"): cells[6].text.strip(),
            col_map.get("Test Name", "Test Name"): cells[7].text.strip(),
            col_map.get("Result", "Result"): cells[8].text.strip(),
            col_map.get("Percent", "Percent"): cells[9].text.strip(),
            col_map.get("Duration", "Duration"): cells[10].text.strip(),
            col_map.get("Centre Name", "Centre Name"): cells[11].text.strip(),
            col_map.get("Report URL", "Report URL"): datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            col_map.get("Report Download", "Report Download"): "",
            col_map.get("PDF Direct Link", "PDF Direct Link"): "",
            col_map.get("Result Sent", "Result Sent"): "",
            col_map.get("Result Sent By", "Result Sent By"): "",
            col_map.get("E-Certificate", "E-Certificate"): "",
            col_map.get("E-Certificate By", "E-Certificate By"): "",
            col_map.get("Certificate", "Certificate"): "",
            col_map.get("Certificate By", "Certificate By"): "",
            col_map.get("Comments", "Comments"): ""
        }
        h = unique_row_hash(data)
        if h in existing_hashes:
            continue
        all_rows.append(data)
    return all_rows

def select_table_row(driver, row: Dict) -> bool:
    time.sleep(6)
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
            time.sleep(4)
            log(f"Selected row in table: {row['First name']} {row['Last name']} ({row['Test Name']})")
            return True
    log("Could not find table row to select!")
    return False

def click_candidate_report_button(driver) -> bool:
    btn = safe_find(driver, By.ID, "button_candidatereport")
    btn.click()
    time.sleep(6)
    log("Clicked Candidate Report button.")
    return True

def get_column_map():
    # Maps old/existing column names to new ones if renamed
    return {
        "Downloaded At": "Report URL",
        "Report Downloaded At": "Report Download",
        "E-Certificate Sent": "E-Certificate",
        "Certificate Issued": "Certificate"
    }

def main():
    # Column order and names
    columns = [
        "Candidate ref.", "First name", "Last name", "Completed",
        "Test Name", "Result", "Percent", "Duration", "Centre Name",
        "Report URL", "Report Download", "Result Sent", "Result Sent By",
        "E-Certificate", "E-Certificate By", "Certificate", "Certificate By",
        "Comments", "Keycode", "Subject", "PDF Direct Link"
    ]
    col_map = get_column_map()

    initialize_excel(EXCEL_FILE, columns)
    existing_df = load_existing_results(EXCEL_FILE)

    # Rename any legacy columns
    if not existing_df.empty:
        existing_df.rename(columns=col_map, inplace=True)
        save_results(EXCEL_FILE, existing_df)

    accounts = load_credentials(CREDENTIALS_FILE)
    for idx_acc, account in enumerate(accounts):
        username = account.get("username", "").strip()
        password = account.get("password", "").strip()
        if not username or not password:
            log(f"Credentials missing for account #{idx_acc+1}, skipping this account.")
            continue

        log(f"--- Starting for account #{idx_acc+1}: {username} ---")
        driver = start_driver()
        try:
            try:
                log("Logging in...")
                login(driver, username, password)
            except Exception as e:
                log(f"Login failed for {username}: {e}")
                continue

            goto_results_tab(driver)
            switch_to_results_iframe(driver)
            reset_and_refresh(driver)

            log("Parsing table...")
            # Deduplicate hashes across ALL accounts so far
            result_df = load_existing_results(EXCEL_FILE)
            result_df.rename(columns=col_map, inplace=True)
            existing_hashes = set(unique_row_hash(row) for _, row in result_df.iterrows()) if not result_df.empty else set()

            new_rows = parse_results_table(driver, existing_hashes, col_map)
            if new_rows:
                new_df = pd.DataFrame(new_rows)
                result_df = pd.concat([result_df, new_df], ignore_index=True)
                save_results(EXCEL_FILE, result_df)
                log(f"Added {len(new_rows)} new rows to Excel.")
            else:
                log("No new rows found.")

            # Step 1: Scrape missing PDF links
            for idx, row in result_df[
                (result_df["PDF Direct Link"].isnull()) | (result_df["PDF Direct Link"] == "") | (result_df["PDF Direct Link"] == "NOT FOUND")
            ].iterrows():
                try:
                    log(f"Processing for PDF link: {row['First name']} {row['Last name']} ({row['Test Name']})")
                    select_table_row(driver, row)
                    click_candidate_report_button(driver)
                    time.sleep(6)
                    html = driver.page_source
                    file_name = extract_pdf_filename_from_html(html)
                    if file_name:
                        pdf_url = f"https://evolve.cityandguilds.com/secureassess/CustomerData/Evolve/DocumentStore/{file_name}"
                        log(f"PDF direct link: {pdf_url}")
                        result_df.at[idx, "PDF Direct Link"] = pdf_url
                        save_results(EXCEL_FILE, result_df)
                    else:
                        log("No PDF file name found! Skipping.")
                        result_df.at[idx, "PDF Direct Link"] = "NOT FOUND"
                        save_results(EXCEL_FILE, result_df)
                    driver.switch_to.default_content()
                    results_tab = safe_find(driver, By.XPATH, "//a[@href='#TestAdministration/Results']")
                    results_tab.click()
                    time.sleep(6)
                    switch_to_results_iframe(driver)
                    reset_and_refresh(driver)
                except Exception as e:
                    log(f"Error processing row idx={idx}: {e}")
                    try:
                        driver.switch_to.default_content()
                        goto_results_tab(driver)
                        switch_to_results_iframe(driver)
                    except Exception as ex:
                        log(f"Failed to recover after error: {ex}")
                        break

            # Step 2: Download PDFs
            num_downloaded = 0
            for idx, row in result_df[
                (result_df["PDF Direct Link"].notnull()) &
                (result_df["PDF Direct Link"] != "") &
                (result_df["PDF Direct Link"] != "NOT FOUND") &
                ((result_df["Report Download"].isnull()) | (result_df["Report Download"] == ""))
            ].iterrows():
                try:
                    pdf_url = row["PDF Direct Link"]
                    if not pdf_url or pdf_url == "NOT FOUND":
                        continue
                    completed = row["Completed"]
                    target_dir = make_report_folder_path(completed)
                    target_name = report_filename(row)
                    save_path = os.path.join(target_dir, target_name)
                    r = requests.get(pdf_url, stream=True)
                    if r.status_code == 200:
                        with open(save_path, 'wb') as f:
                            for chunk in r.iter_content(10240):
                                f.write(chunk)
                        log(f"PDF downloaded and saved to {save_path}")
                        result_df.at[idx, "Report Download"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        save_results(EXCEL_FILE, result_df)
                        num_downloaded += 1
                    else:
                        log(f"Failed to download PDF for idx={idx}. Status: {r.status_code}")
                except Exception as e:
                    log(f"Error downloading PDF for idx={idx}: {e}")

            # Rearranging columns as requested
            if not result_df.empty:
                for c in ["Keycode", "Subject", "PDF Direct Link"]:
                    if c in result_df.columns:
                        result_df = pd.concat([result_df.drop([c], axis=1), result_df[[c]]], axis=1)
                save_results(EXCEL_FILE, result_df)

            # Sort by Completed date (oldest first)
            if "Completed" in result_df.columns:
                try:
                    result_df["Completed_sort"] = pd.to_datetime(result_df["Completed"], format="%d/%m/%Y", errors="coerce")
                    result_df = result_df.sort_values(by="Completed_sort", ascending=True).drop(columns="Completed_sort")
                    save_results(EXCEL_FILE, result_df)
                except Exception as e:
                    log(f"Sorting failed: {e}")

            log(f"All done for account {username}. {num_downloaded} new reports downloaded.\n")

        finally:
            driver.quit()
            log("Chrome closed for this account.\n")

if __name__ == "__main__":
    main()