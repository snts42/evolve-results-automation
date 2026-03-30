"""
Filter management utilities for E-volve platform.
Handles date range filters on the results table.
"""
import time
import logging
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


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
        logging.info(f"Current filter range: {filter_content.text}")
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
        logging.info(f"Step 2 result: {diag}")
        
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
        logging.info(f"Calendar aria-expanded: {expanded}")
        
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
        logging.info(f"Step 4 result: {result}")
        
        if not result.startswith('CALENDAR_FOUND'):
            logging.error(f"Calendar not found: {result}")
            return False
        
        current_month = result.split(':')[1].split('|')[0]
        logging.info(f"Calendar showing: {current_month}")
        
        # Step 5: Navigate back exactly ONE month
        driver.execute_script("""
            var calendars = document.querySelectorAll('.dx-calendar');
            var calendar = calendars[calendars.length - 1];
            var prevBtn = calendar.querySelector('.dx-calendar-navigator-previous-month');
            if (prevBtn) prevBtn.click();
        """)
        time.sleep(0.7)
        
        # Read the new month/year after navigating back
        new_caption = driver.execute_script("""
            var calendars = document.querySelectorAll('.dx-calendar');
            var calendar = calendars[calendars.length - 1];
            var caption = calendar.querySelector('.dx-calendar-caption-button');
            return caption ? caption.textContent : '';
        """)
        logging.info(f"Navigated to: {new_caption}")
        
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
        logging.info(f"1st of month click result: {clicked}")
        
        if clicked != 'CLICKED':
            logging.error(f"Could not click 1st of {new_caption}: {clicked}")
            return False
        
        time.sleep(1)
        logging.info(f"Selected 1st of {new_caption}")
        
        # Step 7: Close the overlay by clicking outside
        driver.execute_script("document.querySelector('.dx-filter-range-content').click();")
        time.sleep(1)
        
        logging.info("Date filter updated successfully")
        return True
        
    except Exception as e:
        logging.error(f"Failed to set date filter: {e}")
        import traceback
        logging.error(traceback.format_exc())
        return False


def clear_date_filter(driver, timeout=10):
    """
    Clear the date filter by clicking the "Clear" option in the filter menu.
    
    Args:
        driver: Selenium WebDriver instance
        timeout: Maximum wait time for elements
    """
    try:
        logging.info("Clearing date filter...")
        
        # Find the filter menu icon (the icon with filter operation)
        filter_menu_icon = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CLASS_NAME, "dx-icon-filter-operation-between"))
        )
        
        # Click to open the context menu
        filter_menu_icon.click()
        time.sleep(0.5)
        
        # Find and click "Clear" option
        clear_option = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, "//span[text()='Clear']"))
        )
        clear_option.click()
        time.sleep(1)
        
        logging.info("Date filter cleared")
        return True
        
    except Exception as e:
        logging.error(f"Failed to clear date filter: {e}")
        return False
