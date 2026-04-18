import time
import logging
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from .config import RESULTS_URL
from .parsing_utils import unique_row_hash

# JS snippet to find the last (most recently opened) calendar in the DOM
_JS_LAST_CALENDAR = "var calendars = document.querySelectorAll('.dx-calendar'); var calendar = calendars[calendars.length - 1];"

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

MATCH_COLS = ["Enrolment no.", "First name", "Last name", "Completed", "Test Name", "Result"]

def _get_screen_size():
    """Get primary screen resolution (width, height) using Windows API."""
    try:
        import ctypes
        user32 = ctypes.windll.user32
        return user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
    except Exception:
        return 1920, 1080  # Safe fallback

def start_driver(headless=True):   
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    if headless:  
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--window-size=3840,2160")
        chrome_options.add_argument("--force-device-scale-factor=0.5")
    else:
        sw, sh = _get_screen_size()
        # Scale factor shrinks the viewport so all 50 datagrid rows fit in the DOM.
        # DevExtreme virtualises rows outside the visible area - too little zoom-out
        # means rows are removed from the DOM and can't be scraped.
        scale = 0.5 if sh >= 1080 else 0.24
        chrome_options.add_argument(f"--force-device-scale-factor={scale}")
        logging.debug(f"Screen resolution: {sw}x{sh}, scale factor: {scale}")
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
    # Maintenance check: if login fields are missing after page load, site may be down
    if not driver.find_elements(By.ID, "UserName"):
        raise ConnectionError(
            "E-volve login page did not load. The site may be under maintenance. "
            "Try again later or check https://evolve.cityandguilds.com manually.")
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

def switch_to_results_iframe(driver, wait=10):
    time.sleep(wait)
    iframe = safe_find(driver, By.ID, "TestAdministrationResultsFrame")
    driver.switch_to.frame(iframe)

def reset_and_refresh(driver):
    logging.info("Loading results...")
    refresh_btn = safe_find(driver, By.XPATH, "//i[contains(@class,'dx-icon-refresh')]")
    refresh_btn.click()
    time.sleep(5)

def parse_results_table(driver, existing_hashes):
    for attempt in range(3):
        try:
            rows = driver.find_elements(By.XPATH, ROW_XPATH)
            all_rows = []
            page_hashes = set()
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if not any(cell.text.strip() for cell in cells) or len(cells) < 12:
                    continue
                data = {
                    col: cells[ci].text.strip() for col, ci in COL_INDEX.items()
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
                page_hashes.add(h)
                if h in existing_hashes:
                    continue
                all_rows.append(data)
            return all_rows, page_hashes
        except StaleElementReferenceException:
            if attempt < 2:
                time.sleep(5)
                if attempt == 1:
                    logging.warning("Results still loading, retrying...")
            else:
                raise

def select_table_row(driver, row):
    table_rows = driver.find_elements(By.XPATH, ROW_XPATH)
    for tr in table_rows:
        tds = tr.find_elements(By.TAG_NAME, "td")
        if not tds or len(tds) < 12:
            continue
        matches = all(
            tds[COL_INDEX[col]].text.strip() == str(row[col]).strip()
            for col in MATCH_COLS
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
    for attempt in range(3):
        try:
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
                        time.sleep(10)
                        return True
            except StaleElementReferenceException:
                raise
            except Exception as e:
                logging.debug(f"Page number click failed: {e}")

            # Strategy 2: Click the Next button via JavaScript
            next_btn = driver.find_element(By.CSS_SELECTOR, ".dx-navigate-button.dx-next-button")
            if "dx-button-disable" in (next_btn.get_attribute("class") or ""):
                return False
            driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", next_btn)
            time.sleep(10)
            return True
        except StaleElementReferenceException:
            if attempt < 2:
                logging.warning(f"Stale element in click_next_page (attempt {attempt + 1}), retrying after 5s...")
                time.sleep(5)
            else:
                logging.warning("Next page navigation failed after 3 attempts")
                return False

def navigate_to_results(driver):
    """Navigate back to results page and switch into the iframe."""
    driver.switch_to.default_content()
    driver.refresh()
    switch_to_results_iframe(driver)

def set_date_filter(driver, months_back=1, timeout=10):
    """
    Set the date filter start to the 1st of the month N months before the
    calendar's current month.

    Args:
        driver: Selenium WebDriver instance
        months_back: How many months to navigate back (1-60)
        timeout: Maximum wait time for elements
    """
    try:
        logging.info(f"Filtering results from last {months_back} month(s)...")
        
        # Step 1: Find and click the filter range content to open overlay
        filter_content = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CLASS_NAME, "dx-filter-range-content"))
        )
        filter_content.click()
        time.sleep(2)
        
        # Step 2: Find the start date container and click the dropdown button.
        # Search directly from document, not scoped to an overlay wrapper which
        # may be empty before the calendar is rendered.
        result = driver.execute_script("""
            var startContainer = document.querySelector('.dx-datagrid-filter-range-start');
            if (!startContainer) return 'START_NOT_FOUND';
            var btn = startContainer.querySelector('.dx-dropdowneditor-button');
            if (!btn) return 'BUTTON_NOT_FOUND';
            btn.click();
            return 'CLICKED';
        """)
        if result != 'CLICKED':
            logging.error(f"Could not click dropdown button: {result}")
            return False
        
        time.sleep(2)
        
        # Step 4: Find the calendar and read current month
        # Calendar popup is rendered outside the start container, search from document
        result = driver.execute_script(f"""
            {_JS_LAST_CALENDAR}
            if (!calendar) return 'CALENDAR_NOT_FOUND|count:0';
            var caption = calendar.querySelector('.dx-calendar-caption-button');
            var captionText = caption ? caption.textContent : 'NO_CAPTION';
            return 'CALENDAR_FOUND:' + captionText + '|count:' + calendars.length;
        """)
        if not result.startswith('CALENDAR_FOUND'):
            logging.error(f"Calendar not found: {result}")
            return False
        
        # Step 5: Navigate back months
        click_delay = 0.5 if months_back > 12 else 1
        for i in range(months_back):
            driver.execute_script(f"""
                {_JS_LAST_CALENDAR}
                var prevBtn = calendar.querySelector('.dx-calendar-navigator-previous-month');
                if (prevBtn) prevBtn.click();
            """)
            time.sleep(click_delay)
            if months_back > 12 and (i + 1) % 12 == 0:
                logging.info(f"  Navigated back {i + 1}/{months_back} months...")
        
        # Read the new month/year after navigating back
        new_caption = driver.execute_script(f"""
            {_JS_LAST_CALENDAR}
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
            {_JS_LAST_CALENDAR}
            var cell = calendar.querySelector('td.dx-calendar-cell[data-value="{target_value}"]');
            if (!cell) return 'CELL_NOT_FOUND';
            cell.click();
            return 'CLICKED';
        """)
        if clicked != 'CLICKED':
            logging.error(f"Could not click 1st of {new_caption}: {clicked}")
            return False
        
        # Scale wait time with date range size
        if months_back <= 3:
            wait = 10
        elif months_back <= 24:
            wait = 15
        else:
            wait = 25
        time.sleep(wait)
        
        # Dismiss date filter overlay by clicking the grid body
        driver.execute_script("""
            var grid = document.querySelector('.dx-datagrid-rowsview');
            if (grid) grid.click();
        """)
        time.sleep(1)
        
        logging.info(f"Date filter set to 1st {new_caption.strip()}")
        return True
        
    except Exception as e:
        logging.error(f"Failed to set date filter: {e}")
        return False


def handle_duplicate_page(driver, page_num, page_hashes, prev_page_hashes, scrape_fn):
    """Detect and recover from E-volve serving the same page twice.
    scrape_fn(driver, page_num) should re-scrape and return new page_hashes.
    Returns resolved page_hashes, or None if unrecoverable (caller should break)."""
    if not (page_hashes and prev_page_hashes and page_hashes == prev_page_hashes):
        return page_hashes

    # Step 1: wait and re-read
    logging.warning(f"Page {page_num} looks identical to previous, waiting for load...")
    time.sleep(10)
    page_hashes = scrape_fn(driver, page_num)
    if page_hashes != prev_page_hashes:
        return page_hashes

    # Step 2: full table refresh
    logging.warning(f"Page {page_num} still identical, refreshing table...")
    try:
        reset_and_refresh(driver)
        for _ in range(page_num - 1):
            if not click_next_page(driver):
                break
        page_hashes = scrape_fn(driver, page_num)
    except Exception as e:
        logging.warning(f"Recovery refresh failed: {e}")

    if page_hashes != prev_page_hashes:
        return page_hashes

    # Step 3: give up
    logging.warning(f"Evolve pagination bug: page {page_num} could not be resolved. "
                    f"Stopping pagination early.")
    return None
