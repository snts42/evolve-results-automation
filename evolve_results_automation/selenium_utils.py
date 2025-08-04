import time
import re
import logging
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from .config import CHROME_DRIVER_PATH

def start_driver(headless=True):   
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--force-device-scale-factor=0.5") 
    if headless:  
        chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--log-level=3") 
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging']) 
    service = Service(executable_path=CHROME_DRIVER_PATH)
    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
    except Exception as e:
        print("\nERROR: Failed to start ChromeDriver.\n" \
              "Please ensure that 'chromedriver.exe' matches your installed version of Google Chrome.\n" \
              "You can download the correct version from: https://chromedriver.chromium.org/downloads\n" \
              f"Original error: {e}")
        raise
    return driver

def safe_find(driver, by, value, timeout=15):
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    except Exception:
        with open("debug_failed_to_find_element.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        raise Exception(f"Element not found (timeout): {value}")

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
    logging.info("Login submitted. Wait for dashboard to load...")

def goto_results_tab(driver):
    driver.get("https://evolve.cityandguilds.com/#TestAdministration/Results")
    time.sleep(6)
    test_admin = safe_find(driver, By.XPATH, "//a[@data-id='TestAdministration']")
    test_admin.click()
    time.sleep(6)
    results_tab = safe_find(driver, By.XPATH, "//a[@href='#TestAdministration/Results']")
    results_tab.click()
    time.sleep(6)
    logging.info("Results tab opened. Please check Results table visible.")

def switch_to_results_iframe(driver, timeout=15):
    time.sleep(6)
    iframe = safe_find(driver, By.ID, "TestAdministrationResultsFrame")
    driver.switch_to.frame(iframe)
    time.sleep(6)
    logging.info("Switched to Results iframe.")

def switch_to_default(driver):
    driver.switch_to.default_content()
    time.sleep(1)
    logging.info("Switched back to default content.")

def reset_and_refresh(driver):
    refresh_btn = safe_find(driver, By.XPATH, "//i[contains(@class,'dx-icon-refresh')]")
    refresh_btn.click()
    time.sleep(6)
    logging.info("Refresh button clicked. Wait for table reload.")

def parse_results_table(driver, existing_hashes, col_map):
    time.sleep(6)
    row_xpath = (
        "//div[contains(@class, 'dx-datagrid-rowsview')]"
        "//table[contains(@class, 'dx-datagrid-table')]"
        "/tbody/tr[contains(@class, 'dx-row') and not(contains(@class, 'dx-freespace-row'))]"
    )
    rows = driver.find_elements(By.XPATH, row_xpath)
    logging.info(f"Found {len(rows)} data rows in the results table.")
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
        h = "|".join([str(data.get(f, "")).strip().lower() for f in [
            "Candidate ref.", "First name", "Last name", "Completed", "Test Name", "Result"
        ]])
        if h in existing_hashes:
            continue
        all_rows.append(data)
    return all_rows

def select_table_row(driver, row):
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
            logging.info(f"Selected row in table: {row['First name']} {row['Last name']} ({row['Test Name']})")
            return True
    logging.info("Could not find table row to select!")
    return False

def click_candidate_report_button(driver):
    btn = safe_find(driver, By.ID, "button_candidatereport")
    btn.click()
    time.sleep(6)
    logging.info("Clicked Candidate Report button.")
    return True

def extract_pdf_filename_from_html(html):
    match = re.search(r'([a-f0-9\-]{36}\.pdf)', html)
    if match:
        return match.group(1)
    return None