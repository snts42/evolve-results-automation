import time
import re
import logging
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from .parsing_utils import unique_row_hash

ROW_XPATH = (
    "//div[contains(@class, 'dx-datagrid-rowsview')]"
    "//table[contains(@class, 'dx-datagrid-table')]"
    "/tbody/tr[contains(@class, 'dx-row') and not(contains(@class, 'dx-freespace-row'))]"
)

def start_driver(headless=True):   
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--force-device-scale-factor=0.5") 
    if headless:  
        chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--log-level=3") 
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging']) 
    try:
        driver = webdriver.Chrome(options=chrome_options)
    except Exception as e:
        print("\nERROR: Failed to start Chrome.\n" \
              "Please ensure Google Chrome is installed.\n" \
              "ChromeDriver is managed automatically by Selenium Manager.\n" \
              f"Original error: {e}")
        raise
    return driver

def safe_find(driver, by, value, timeout=15):
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    except Exception as e:
        raise Exception(f"Element not found (timeout): {value}") from e

def login(driver, username: str, password: str):
    driver.get("https://evolve.cityandguilds.com/Login")
    time.sleep(4)
    user_box = safe_find(driver, By.ID, "UserName")
    pass_box = safe_find(driver, By.ID, "Password")
    user_box.clear()
    user_box.send_keys(username)
    pass_box.clear()
    pass_box.send_keys(password)
    login_btn = safe_find(driver, By.XPATH, "//input[@type='submit' and @value='Login']")
    login_btn.click()
    logging.info("Login submitted. Wait for dashboard to load...")
    time.sleep(8)  # Wait for dashboard to fully load after login

def goto_results_tab(driver):
    driver.get("https://evolve.cityandguilds.com/#TestAdministration/Results")
    time.sleep(6)  # Wait for Results page to load
    logging.info("Results tab opened. Please check Results table visible.")

def switch_to_results_iframe(driver, timeout=15):
    time.sleep(5)
    iframe = safe_find(driver, By.ID, "TestAdministrationResultsFrame")
    driver.switch_to.frame(iframe)
    logging.info("Switched to Results iframe.")

def reset_and_refresh(driver):
    refresh_btn = safe_find(driver, By.XPATH, "//i[contains(@class,'dx-icon-refresh')]")
    refresh_btn.click()
    time.sleep(4)
    logging.info("Refresh button clicked. Wait for table reload.")

def parse_results_table(driver, existing_hashes, col_map):
    time.sleep(5)
    rows = driver.find_elements(By.XPATH, ROW_XPATH)
    all_rows = []
    real_rows = 0
    for idx, row in enumerate(rows):
        cells = row.find_elements(By.TAG_NAME, "td")
        if not any(cell.text.strip() for cell in cells) or len(cells) < 12:
            continue
        real_rows += 1
        data = {
            col_map.get("Keycode", "Keycode"): cells[1].text.strip(),
            col_map.get("Enrolment no.", "Enrolment no."): cells[2].text.strip(),
            col_map.get("First name", "First name"): cells[3].text.strip(),
            col_map.get("Last name", "Last name"): cells[4].text.strip(),
            col_map.get("Completed", "Completed"): cells[5].text.strip(),
            col_map.get("Subject", "Subject"): cells[6].text.strip(),
            col_map.get("Test Name", "Test Name"): cells[7].text.strip(),
            col_map.get("Result", "Result"): cells[8].text.strip(),
            col_map.get("Percent", "Percent"): cells[9].text.strip(),
            col_map.get("Duration", "Duration"): cells[10].text.strip(),
            col_map.get("Centre Name", "Centre Name"): cells[11].text.strip(),
            col_map.get("Scraping date/time", "Scraping date/time"): datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            col_map.get("PDF report downloaded", "PDF report downloaded"): "",
            col_map.get("Result Sent", "Result Sent"): "",
            col_map.get("Result Sent By", "Result Sent By"): "",
            col_map.get("E-Certificate sent", "E-Certificate sent"): "",
            col_map.get("E-Certificate By", "E-Certificate By"): "",
            col_map.get("Certificate", "Certificate"): "",
            col_map.get("Certificate By", "Certificate By"): "",
            col_map.get("Comments", "Comments"): ""
        }
        h = unique_row_hash(data)
        if h in existing_hashes:
            continue
        all_rows.append(data)

    logging.info(f"Found {real_rows} data rows in the results table, of which {len(all_rows)} are new.")
    return all_rows

def select_table_row(driver, row):
    table_rows = driver.find_elements(By.XPATH, ROW_XPATH)
    for tr in table_rows:
        tds = tr.find_elements(By.TAG_NAME, "td")
        if not tds or len(tds) < 12:
            continue
        matches = all(
            tds[idx].text.strip() == str(row[col]).strip()
            for col, idx in [
                ("Enrolment no.", 2), ("First name", 3), ("Last name", 4),
                ("Completed", 5), ("Test Name", 7), ("Result", 8)
            ]
        )
        if matches:
            driver.execute_script("arguments[0].scrollIntoView(true);", tr)
            tr.click()
            time.sleep(2)  # Wait after clicking row
            logging.info(f"Selected row in table: {row['First name']} {row['Last name']} ({row['Test Name']})")
            return True
    # Row not found - return immediately without logging (too noisy)
    return False

def click_candidate_report_button(driver):
    btn = safe_find(driver, By.ID, "button_candidatereport")
    btn.click()
    logging.info("Clicked Candidate Report button.")
    time.sleep(2)
    return True

def get_total_pages(driver):
    """Get total number of pages from pagination control."""
    try:
        pages_count = driver.find_element(By.CLASS_NAME, "dx-pages-count")
        total = int(pages_count.text)
        return total
    except Exception:
        return 1  # Only 1 page if pagination not found

def click_next_page(driver):
    """Click the Next button to go to next page. Returns True if successful, False if already on last page."""
    try:
        next_btn = driver.find_element(By.CSS_SELECTOR, ".dx-navigate-button.dx-next-button")
        
        # Check if disabled
        if "dx-button-disable" in next_btn.get_attribute("class"):
            return False  # Already on last page
        
        next_btn.click()
        time.sleep(4)  # Wait for page to load
        return True
    except Exception as e:
        logging.warning(f"Failed to click next page: {e}")
        return False