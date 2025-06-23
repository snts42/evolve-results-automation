import time
import json
import os
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ---- CONFIGURATION ----
CHROME_DRIVER_PATH = "chromedriver.exe"
CREDENTIALS_FILE = "credentials.json"
EXCEL_FILE = "exam_results.xlsx"
REPORTS_BASE = "reports"
YEAR = datetime.now().year

# ---- LOGGING ----
def log(msg):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"{timestamp} | {msg}")

# ---- EXCEL MANAGEMENT ----
def initialize_excel(filepath, columns):
    if not os.path.exists(filepath):
        log(f"Excel file not found. Creating {filepath}")
        pd.DataFrame(columns=columns).to_excel(filepath, index=False)

def load_existing_results(filepath):
    if not os.path.exists(filepath):
        return pd.DataFrame()
    return pd.read_excel(filepath, dtype=str)

def save_results(filepath, df):
    df.to_excel(filepath, index=False)

# ---- CREDENTIALS ----
def load_credentials(json_file):
    with open(json_file, 'r') as f:
        creds = json.load(f)
    return creds['username'], creds['password']

# ---- SELENIUM SETUP ----
def start_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    # chrome_options.add_argument("--headless")  # For automation
    service = Service(executable_path=CHROME_DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

# ---- UTILS ----
def step(msg):
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

def get_pdf_original_url_from_shadow_dom(driver, timeout=20):
    import time
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            # 1. Find the shadow host <pdf-viewer id="viewer">
            shadow_host = driver.find_element(By.CSS_SELECTOR, "pdf-viewer#viewer")
            # 2. Get its shadow root
            shadow_root = driver.execute_script('return arguments[0].shadowRoot', shadow_host)
            # 3. Find the <embed> inside the shadow root
            # The <embed> is under #main > #scroller > #size > #contents > embed
            embed = shadow_root.find_element(By.CSS_SELECTOR, '#main #scroller #size #contents embed')
            url = embed.get_attribute("original-url")
            if url and ".pdf" in url.lower():
                return url
        except Exception as e:
            pass  # Not loaded yet
        time.sleep(1)
    return None

def make_report_folder_path(date_str):
    """Given date string like '21/06/2025', returns folder like reports/2025/06 21"""
    dt = datetime.strptime(date_str, "%d/%m/%Y")
    year = dt.strftime("%Y")
    month_day = dt.strftime("%m %d")
    folder = os.path.join(REPORTS_BASE, year, month_day)
    os.makedirs(folder, exist_ok=True)
    return folder

def report_filename(row):
    # Example: Mustafa Eden - 4748-113 Functional Skills English Reading Level 2 - Pass.pdf
    parts = [
        str(row["First name"]).strip(),
        str(row["Last name"]).strip(),
        str(row["Test Name"]).strip(),
        str(row["Result"]).strip()
    ]
    fname = " - ".join(parts) + ".pdf"
    # Clean up forbidden filename chars
    fname = "".join(c for c in fname if c not in r'\/:*?"<>|')
    return fname

def unique_row_hash(row):
    """Returns a string that uniquely identifies a row for deduplication."""
    fields = [
        "Candidate ref.", "First name", "Last name", "Completed", "Test Name", "Result"
    ]
    return "|".join([str(row.get(f, "")).strip().lower() for f in fields])

# ---- LOGIN ----
def login(driver, username, password):
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
    # Must be inside iframe!
    try:
        refresh_btn = safe_find(driver, By.XPATH, "//i[contains(@class,'dx-icon-refresh')]")
        refresh_btn.click()
        log("Refresh button clicked.")
        step("Refresh done, let table reload.")
    except Exception:
        log("Refresh button not found. Skipping.")

# ---- TABLE SCRAPING ----
def parse_results_table(driver, existing_hashes):
    # Must be inside iframe!
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
        driver.execute_script("arguments[0].scrollIntoView(true);", row)
        time.sleep(0.05)
        cells = row.find_elements(By.TAG_NAME, "td")
        cell_texts = [cell.text.strip() for cell in cells]
        if not any(cell_texts) or len(cells) < 12:
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
def select_table_row(driver, row):
    """Selects the row that matches all the unique_fields"""
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
        matches = True
        for col, idx in [("Candidate ref.", 2), ("First name", 3), ("Last name", 4),
                         ("Completed", 5), ("Test Name", 7), ("Result", 8)]:
            if tds[idx].text.strip() != str(row[col]).strip():
                matches = False
                break
        if matches:
            driver.execute_script("arguments[0].scrollIntoView(true);", tr)
            tr.click()
            log(f"Selected row in table: {row['First name']} {row['Last name']} ({row['Test Name']})")
            time.sleep(0.5)
            return True
    log("Could not find table row to select!")
    return False

def click_candidate_report_button(driver):
    try:
        btn = safe_find(driver, By.ID, "button_candidatereport")
        btn.click()
        log("Clicked Candidate Report button.")
        time.sleep(1)
        return True
    except Exception as e:
        log(f"Failed to click Candidate Report button: {e}")
        return False

# ---- PDF DOWNLOAD ----
def download_pdf_from_url(url, save_path):
    import requests
    try:
        r = requests.get(url, stream=True)
        r.raise_for_status()
        with open(save_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
        log(f"PDF downloaded: {save_path}")
        return True
    except Exception as e:
        log(f"Failed to download PDF: {e}")
        return False

# ---- MAIN PROCESS ----
def main():
    columns = [
        "Keycode", "Candidate ref.", "First name", "Last name", "Completed",
        "Subject", "Test Name", "Result", "Percent", "Duration", "Centre Name",
        "Downloaded At", "Report Downloaded At",
        "Result Sent", "E-Certificate Sent", "Certificate Issued", "Comments"
    ]
    initialize_excel(EXCEL_FILE, columns)
    existing_df = load_existing_results(EXCEL_FILE)
    if existing_df.empty:
        existing_hashes = set()
    else:
        existing_hashes = set(unique_row_hash(row) for _, row in existing_df.iterrows())
    username, password = load_credentials(CREDENTIALS_FILE)

    driver = start_driver()
    try:
        login(driver, username, password)
        goto_results_tab(driver)

        switch_to_results_iframe(driver)
        reset_and_refresh(driver)
        step("In Results iframe. Scroll to load all results. Then press Enter.")
        new_rows = parse_results_table(driver, existing_hashes)
        if new_rows:
            new_df = pd.DataFrame(new_rows)
            result_df = pd.concat([existing_df, new_df], ignore_index=True)
            save_results(EXCEL_FILE, result_df)
            log(f"Added {len(new_rows)} new rows to Excel.")
        else:
            result_df = existing_df.copy()
            log("No new rows found.")

        # ----------- DOWNLOAD REPORTS -------------
        # Only switch to iframe ONCE before processing
        for idx, row in result_df[(result_df["Report Downloaded At"].isnull()) | (result_df["Report Downloaded At"] == "")].iterrows():
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
                    # Go back to iframe for next candidate
                    switch_to_results_iframe(driver)
                    continue

                step("Candidate report opened. Wait for PDF viewer to load, then press Enter.")
                # 3. Get PDF original-url from embed
                pdf_url = get_pdf_original_url_from_shadow_dom(driver)
                if not pdf_url:
                    log("Could not find PDF <embed> with original-url. Skipping.")
                    # Go back to results tab if necessary, then to iframe
                    goto_results_tab(driver)
                    switch_to_results_iframe(driver)
                    continue

                log(f"Found PDF URL: {pdf_url}")
                # 4. Download PDF
                folder = make_report_folder_path(row["Completed"])
                fname = report_filename(row)
                fpath = os.path.join(folder, fname)
                downloaded = download_pdf_from_url(pdf_url, fpath)
                # 5. Update Excel
                if downloaded:
                    result_df.at[idx, "Report Downloaded At"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    save_results(EXCEL_FILE, result_df)
                step(f"Processed {row['First name']} {row['Last name']} - move to next candidate?")

                # After download, go back to Results tab, then switch to iframe again for next loop
                goto_results_tab(driver)
                switch_to_results_iframe(driver)
                reset_and_refresh(driver)

            except Exception as e:
                log(f"Error processing row idx={idx}: {e}")
                # Try to recover context for next iteration
                try:
                    goto_results_tab(driver)
                    switch_to_results_iframe(driver)
                except Exception as ex:
                    log(f"Failed to recover after error: {ex}")
                    break

    finally:
        driver.quit()
        log("Chrome closed.")

if __name__ == "__main__":
    main()