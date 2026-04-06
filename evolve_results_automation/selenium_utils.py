import time
import logging
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from .config import RESULTS_URL
from .parsing_utils import unique_row_hash

ROW_XPATH = (
    "//div[contains(@class, 'dx-datagrid-rowsview')]"
    "//table[contains(@class, 'dx-datagrid-table')]"
    "/tbody/tr[contains(@class, 'dx-row') and not(contains(@class, 'dx-freespace-row'))]"
)

# Column index mapping for the results table (0-based td indices)
COL_INDEX = {
    "Keycode": 1, "Enrolment no.": 2, "First name": 3, "Last name": 4,
    "Completed": 5, "Subject": 6, "Test Name": 7, "Result": 8,
    "Percent": 9, "Duration": 10, "Centre Name": 11,
}

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
        logging.error(f"Failed to start Chrome. Ensure Google Chrome is installed. Error: {e}")
        raise
    return driver

def safe_find(driver, by, value, timeout=15):
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    except Exception as e:
        raise Exception(f"Element not found (timeout): {value}") from e

def login(driver, username: str, password: str):
    driver.get(RESULTS_URL)
    time.sleep(5)
    user_box = safe_find(driver, By.ID, "UserName")
    pass_box = safe_find(driver, By.ID, "Password")
    user_box.clear()
    user_box.send_keys(username)
    pass_box.clear()
    pass_box.send_keys(password)
    login_btn = safe_find(driver, By.XPATH, "//input[@type='submit' and @value='Login']")
    login_btn.click()
    time.sleep(3)
    # Check for login failure
    errors = driver.find_elements(By.CLASS_NAME, "validation-summary-errors")
    if errors:
        error_text = errors[0].text.strip()
        raise Exception(f"Login failed for {username}: {error_text}")
    logging.info("Login submitted")

def switch_to_results_iframe(driver, wait=10, timeout=15):
    time.sleep(wait)
    iframe = safe_find(driver, By.ID, "TestAdministrationResultsFrame")
    driver.switch_to.frame(iframe)

def reset_and_refresh(driver):
    logging.info("Refreshing table...")
    refresh_btn = safe_find(driver, By.XPATH, "//i[contains(@class,'dx-icon-refresh')]")
    refresh_btn.click()
    time.sleep(5)

def parse_results_table(driver, existing_hashes):
    rows = driver.find_elements(By.XPATH, ROW_XPATH)
    all_rows = []
    real_rows = 0
    for idx, row in enumerate(rows):
        cells = row.find_elements(By.TAG_NAME, "td")
        if not any(cell.text.strip() for cell in cells) or len(cells) < 12:
            continue
        real_rows += 1
        data = {
            col: cells[idx].text.strip() for col, idx in COL_INDEX.items()
        }
        data.update({
            "Scraping date/time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "PDF report save time": "",
            "Result Sent": "",
            "Result Sent By": "",
            "E-Certificate sent": "",
            "E-Certificate By": "",
            "Certificate": "",
            "Certificate By": "",
            "Comments": ""
        })
        h = unique_row_hash(data)
        if h in existing_hashes:
            continue
        all_rows.append(data)
    logging.info(f"Found {real_rows} data rows, {len(all_rows)} new")
    return all_rows

def select_table_row(driver, row):
    table_rows = driver.find_elements(By.XPATH, ROW_XPATH)
    for tr in table_rows:
        tds = tr.find_elements(By.TAG_NAME, "td")
        if not tds or len(tds) < 12:
            continue
        match_cols = ["Enrolment no.", "First name", "Last name", "Completed", "Test Name", "Result"]
        matches = all(
            tds[COL_INDEX[col]].text.strip() == str(row[col]).strip()
            for col in match_cols
        )
        if matches:
            driver.execute_script("arguments[0].scrollIntoView(true);", tr)
            tr.click()
            time.sleep(0.5)
            return True
    return False

def click_candidate_report_button(driver):
    btn = safe_find(driver, By.ID, "button_candidatereport")
    btn.click()
    time.sleep(5)

def get_total_pages(driver):
    """Get total number of pages from pagination control."""
    try:
        pages_count = driver.find_element(By.CLASS_NAME, "dx-pages-count")
        total = int(pages_count.text)
        return total
    except Exception:
        return 1  # Only 1 page if pagination not found

def _get_current_page(driver):
    """Get the currently selected page number from pagination."""
    try:
        selected = driver.find_element(By.CSS_SELECTOR, ".dx-page.dx-selection")
        return int(selected.text)
    except Exception:
        return None

def click_next_page(driver):
    """Click the Next button to go to next page. Returns True if successful, False if already on last page."""
    page_before = _get_current_page(driver)
    target_page = (page_before or 1) + 1

    # Strategy 1: Click the target page number directly (most reliable)
    try:
        pages = driver.find_elements(By.CSS_SELECTOR, ".dx-page")
        for p in pages:
            if p.text.strip() == str(target_page):
                driver.execute_script("arguments[0].scrollIntoView(true);", p)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", p)
                logging.info(f"Navigated to page {target_page}")
                time.sleep(10)
                return True
    except Exception as e:
        logging.debug(f"Page number click failed: {e}")

    # Strategy 2: Click the Next button via JavaScript
    try:
        next_btn = driver.find_element(By.CSS_SELECTOR, ".dx-navigate-button.dx-next-button")
        if "dx-button-disable" in (next_btn.get_attribute("class") or ""):
            logging.info("Next button is disabled - already on last page")
            return False
        driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", next_btn)
        logging.info(f"Navigated to next page")
        time.sleep(10)
        return True
    except Exception as e:
        logging.warning(f"Next button click failed: {e}")
        return False

def navigate_to_results(driver):
    """Navigate back to results page and switch into the iframe."""
    driver.switch_to.default_content()
    driver.refresh()
    switch_to_results_iframe(driver)

def set_date_filter_to_previous_month_start(driver, timeout=10):
    """
    Set the date filter start to the 1st of the month before the calendar's current month.
    E.g. if calendar shows February, navigate back to January and select the 1st.
    
    Strategy:
    1. Click on dx-filter-range-content to open the overlay
    2. Use JS to click the dropdown button inside the overlay to open calendar
    3. Navigate back ONE month
    4. Click the 1st of that month
    5. Close the overlay
    
    Args:
        driver: Selenium WebDriver instance
        timeout: Maximum wait time for elements
    """
    try:
        logging.info("Setting date filter to start from 1st of previous month...")
        
        # Step 1: Find and click the filter range content to open overlay
        filter_content = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CLASS_NAME, "dx-filter-range-content"))
        )
        filter_content.click()
        time.sleep(2)
        
        # Step 2: Find the start date container and click the dropdown button
        # Search directly from document (not scoped to overlay wrapper which may be empty)
        diag = driver.execute_script("""
            // Diagnostic: check what elements exist in the DOM
            var overlayWrappers = document.querySelectorAll('.dx-datagrid-filter-range-overlay');
            var overlayContents = document.querySelectorAll('.dx-overlay-content.dx-datagrid');
            var startContainers = document.querySelectorAll('.dx-datagrid-filter-range-start');
            var info = 'wrappers:' + overlayWrappers.length + 
                       ' contents:' + overlayContents.length + 
                       ' starts:' + startContainers.length;
            
            // Find the start container directly from document
            var startContainer = document.querySelector('.dx-datagrid-filter-range-start');
            if (!startContainer) return 'START_NOT_FOUND|' + info;
            
            // Find the dropdown button
            var btn = startContainer.querySelector('.dx-dropdowneditor-button');
            if (!btn) return 'BUTTON_NOT_FOUND|' + info + '|' + startContainer.innerHTML.substring(0, 200);
            
            btn.click();
            return 'CLICKED_BUTTON|' + info;
        """)
        if not diag.startswith('CLICKED_BUTTON'):
            logging.error(f"Could not click dropdown button. Diagnostic: {diag}")
            return False
        
        time.sleep(2)
        
        # Step 3: Check if calendar is now visible (aria-expanded should be true)
        expanded = driver.execute_script("""
            var startContainer = document.querySelector('.dx-datagrid-filter-range-start');
            var input = startContainer.querySelector('input.dx-texteditor-input');
            return input ? input.getAttribute('aria-expanded') : 'INPUT_NOT_FOUND';
        """)
        if expanded != 'true':
            # Try clicking the input field to trigger the calendar
            logging.info("Calendar not expanded, trying to click input field...")
            driver.execute_script("""
                var startContainer = document.querySelector('.dx-datagrid-filter-range-start');
                var input = startContainer.querySelector('input.dx-texteditor-input');
                if (input) { input.focus(); input.click(); }
            """)
            time.sleep(2)
        
        # Step 4: Find the calendar and read current month
        # Calendar popup is rendered outside the start container, search from document
        result = driver.execute_script("""
            var calendars = document.querySelectorAll('.dx-calendar');
            if (calendars.length === 0) return 'CALENDAR_NOT_FOUND|count:0';
            
            // Use the last calendar (most recently opened)
            var calendar = calendars[calendars.length - 1];
            var caption = calendar.querySelector('.dx-calendar-caption-button');
            var captionText = caption ? caption.textContent : 'NO_CAPTION';
            return 'CALENDAR_FOUND:' + captionText + '|count:' + calendars.length;
        """)
        if not result.startswith('CALENDAR_FOUND'):
            logging.error(f"Calendar not found: {result}")
            return False
        
        # Step 5: Navigate back exactly ONE month
        driver.execute_script("""
            var calendars = document.querySelectorAll('.dx-calendar');
            var calendar = calendars[calendars.length - 1];
            var prevBtn = calendar.querySelector('.dx-calendar-navigator-previous-month');
            if (prevBtn) prevBtn.click();
        """)
        time.sleep(1)
        
        # Read the new month/year after navigating back
        new_caption = driver.execute_script("""
            var calendars = document.querySelectorAll('.dx-calendar');
            var calendar = calendars[calendars.length - 1];
            var caption = calendar.querySelector('.dx-calendar-caption-button');
            return caption ? caption.textContent : '';
        """)
        # Parse the target month/year from the caption (e.g. "January 2026")
        try:
            target_date = datetime.strptime(new_caption.strip(), "%B %Y")
            target_value = target_date.strftime("%Y/%m/01")
        except ValueError:
            logging.error(f"Could not parse calendar caption: {new_caption}")
            return False
        
        # Step 6: Click the 1st of the target month
        clicked = driver.execute_script(f"""
            var calendars = document.querySelectorAll('.dx-calendar');
            var calendar = calendars[calendars.length - 1];
            var cell = calendar.querySelector('td.dx-calendar-cell[data-value="{target_value}"]');
            if (!cell) return 'CELL_NOT_FOUND';
            cell.click();
            return 'CLICKED';
        """)
        if clicked != 'CLICKED':
            logging.error(f"Could not click 1st of {new_caption}: {clicked}")
            return False
        
        time.sleep(10)
        logging.info("Date filter updated")
        return True
        
    except Exception as e:
        logging.error(f"Failed to set date filter: {e}")
        return False