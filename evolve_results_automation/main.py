import os
import time
import glob
import logging
import pandas as pd
import requests
from datetime import datetime
from selenium.webdriver.common.by import By

from evolve_results_automation.gui_tk import EvolveGUI

from evolve_results_automation.config import (
   ENCRYPTED_CREDENTIALS_FILE, COLUMNS, BASE_DIR, get_excel_file_for_year
)
from evolve_results_automation.excel_utils import (
    initialize_excel, load_existing_results, save_results
)
from evolve_results_automation.selenium_utils import (
    start_driver, login, goto_results_tab, switch_to_results_iframe,
    reset_and_refresh, parse_results_table, select_table_row,
    click_candidate_report_button, safe_find, get_total_pages,
    click_next_page
)

from evolve_results_automation.secure_credentials import SecureCredentialManager

from evolve_results_automation.logging_utils import setup_logger
setup_logger()

from evolve_results_automation.parsing_utils import make_report_folder_path, report_filename, extract_pdf_filename_from_html, unique_row_hash, get_column_map
from evolve_results_automation.filter_utils import set_date_filter_to_previous_month_start

from dataclasses import dataclass

@dataclass
class ProcessingStats:
    accounts_processed: int = 0
    new_rows_added: int = 0
    pdfs_downloaded: int = 0
    errors_encountered: int = 0

class EvolveAutomation:
    def __init__(self, headless: bool, master_password: str, selected_username: str = None):
        self.headless = headless
        self.master_password = master_password
        self.selected_username = selected_username
        self.stats = ProcessingStats()
        self.columns = COLUMNS
        self.col_map = get_column_map()

    def _save_year_to_excel(self, year, rows_by_year, silent=False):
        """Save rows for a given year to Excel, merging with existing data and deduplicating."""
        if year not in rows_by_year:
            return
        year_rows = rows_by_year[year]
        excel_file = get_excel_file_for_year(year)
        initialize_excel(excel_file, self.columns)

        # Load existing data (may contain rows from previous accounts)
        existing_df = load_existing_results(excel_file)
        if not existing_df.empty:
            existing_df.rename(columns=self.col_map, inplace=True)

        # Create DataFrame from current rows (includes PDF links from processing)
        new_df = pd.DataFrame(year_rows)

        # Merge existing + new, then deduplicate
        if not existing_df.empty:
            combined = pd.concat([existing_df, new_df], ignore_index=True)
        else:
            combined = new_df.copy()

        # Remove garbage rows missing core fields
        core_fields = ["Completed", "First name", "Last name"]
        for cf in core_fields:
            if cf in combined.columns:
                combined = combined[combined[cf].notna() & (combined[cf] != "")]

        # Deduplicate using unique_row_hash, keep last (in-memory rows have latest PDF links)
        combined['_hash'] = combined.apply(lambda r: unique_row_hash(r), axis=1)
        combined = combined.drop_duplicates(subset='_hash', keep='last')
        combined = combined.drop(columns='_hash')

        # Rearrange columns
        for c in ["Keycode", "Subject"]:
            if c in combined.columns:
                combined = pd.concat([combined.drop([c], axis=1), combined[[c]]], axis=1)

        # Sort by completion date
        if "Completed" in combined.columns:
            try:
                combined["Completed_sort"] = pd.to_datetime(combined["Completed"], format="%d/%m/%Y", errors="coerce")
                combined = combined.sort_values(by="Completed_sort", ascending=True).drop(columns="Completed_sort")
            except Exception as e:
                logging.warning(f"Sorting failed for year {year}: {e}")

        save_results(excel_file, combined)
        if not silent:
            logging.info(f"Saved {len(combined)} rows to {year}/exam_results.xlsx")

    def run(self):
        # We'll initialize Excel files dynamically as we encounter different years
        # Load existing data from all years to build hash set
        # Note: We'll load year-specific files as needed during processing

        accounts = SecureCredentialManager(ENCRYPTED_CREDENTIALS_FILE).decrypt_credentials(self.master_password)
        
        # Filter accounts if specific username selected
        if self.selected_username:
            accounts = [acc for acc in accounts if acc.get("username", "").strip() == self.selected_username]
            if not accounts:
                logging.error(f"Selected account '{self.selected_username}' not found in credentials")
                return self.stats
        
        for idx_acc, account in enumerate(accounts):
            username = account.get("username", "").strip()
            password = account.get("password", "").strip()
            if not username or not password:
                logging.warning(f"Credentials missing for account #{idx_acc+1}, skipping this account.")
                continue
            logging.info(f"--- Starting for account #{idx_acc+1}: {username} ---")
            driver = start_driver(headless=self.headless)
            try:
                login(driver, username, password)
                goto_results_tab(driver)
                switch_to_results_iframe(driver)
                reset_and_refresh(driver)
                
                # Set date filter start to 1st of previous month
                set_date_filter_to_previous_month_start(driver)
                time.sleep(2)  # Wait for filter to apply and table to reload
                
                # Get total pages
                total_pages = get_total_pages(driver)
                logging.info(f"Found {total_pages} page(s) to scrape")
                
                # Build hash set from all existing year Excel files to prevent duplicates
                # Also load existing rows needing PDFs into rows_by_year for resume
                # (matches legacy behavior: ALL rows checked for missing PDFs)
                existing_hashes = set()
                rows_by_year = {}
                pdf_resume_count = 0
                for excel_path in glob.glob(os.path.join(BASE_DIR, '*', 'exam_results.xlsx')):
                    ef = load_existing_results(excel_path)
                    if not ef.empty:
                        ef.rename(columns=self.col_map, inplace=True)
                        for _, r in ef.iterrows():
                            # Skip garbage rows missing core fields
                            completed_val = str(r.get("Completed", "")).strip()
                            if not completed_val or completed_val == "nan":
                                continue
                            first_name = str(r.get("First name", "")).strip()
                            if not first_name or first_name == "nan":
                                continue
                            
                            existing_hashes.add(unique_row_hash(r))
                            # If row needs PDF, add to rows_by_year for processing
                            report_dl = str(r.get("PDF report downloaded", "")).strip()
                            if not report_dl or report_dl in ("", "nan"):
                                try:
                                    yr = datetime.strptime(completed_val, "%d/%m/%Y").year
                                except (ValueError, TypeError):
                                    continue  # Skip rows with unparseable dates
                                if yr not in rows_by_year:
                                    rows_by_year[yr] = []
                                rows_by_year[yr].append(r.to_dict())
                                pdf_resume_count += 1
                logging.info(f"Loaded {len(existing_hashes)} existing hashes from Excel files")
                if pdf_resume_count > 0:
                    logging.info(f"Found {pdf_resume_count} existing rows needing PDF download (resuming)")
                
                # Process each page
                for page_num in range(1, total_pages + 1):
                    logging.info(f"Processing page {page_num}/{total_pages}")
                    
                    # Scrape table on current page
                    new_rows = parse_results_table(driver, existing_hashes, self.col_map)
                    if new_rows:
                        # Group rows by year based on completion date
                        for row in new_rows:
                            completed_date = row.get("Completed", "")
                            try:
                                year = datetime.strptime(completed_date, "%d/%m/%Y").year
                            except (ValueError, TypeError):
                                year = datetime.now().year  # Fallback to current year
                            
                            if year not in rows_by_year:
                                rows_by_year[year] = []
                            rows_by_year[year].append(row)
                            existing_hashes.add(unique_row_hash(row))
                        
                        logging.info(f"Found {len(new_rows)} new rows from page {page_num}")
                        self.stats.new_rows_added += len(new_rows)
                        
                        # Save scraped rows to Excel immediately (before PDF processing)
                        for yr in rows_by_year:
                            self._save_year_to_excel(yr, rows_by_year)
                    else:
                        logging.info(f"No new rows found on page {page_num}")
                    
                    # Collect all rows from rows_by_year for PDF processing
                    all_current_rows = []
                    for year_rows in rows_by_year.values():
                        all_current_rows.extend(year_rows)
                    
                    if not all_current_rows:
                        continue
                    
                    result_df = pd.DataFrame(all_current_rows)
                    
                    # Get rows needing PDF processing: missing PDF report downloaded timestamp
                    no_dl = (result_df["PDF report downloaded"].isnull()) | (result_df["PDF report downloaded"] == "")
                    all_pdf_needed = result_df[no_dl]
                    
                    # Process all rows needing PDFs via web UI
                    for idx, row in all_pdf_needed.iterrows():
                        try:
                            # Skip rows with invalid core fields
                            completed = str(row.get("Completed", "")).strip()
                            if not completed or completed == "nan":
                                continue
                            
                            # Try to select the row - if it's not on this page, skip it
                            if not select_table_row(driver, row):
                                continue  # Row not on this page, skip
                            
                            logging.info(f"Processing PDF for: {row['First name']} {row['Last name']} ({row['Test Name']})")
                            click_candidate_report_button(driver)
                            time.sleep(2)
                            html = driver.page_source
                            file_name = extract_pdf_filename_from_html(html)
                            if file_name:
                                pdf_url = f"https://evolve.cityandguilds.com/secureassess/CustomerData/Evolve/DocumentStore/{file_name}"
                                logging.info(f"PDF URL extracted: {pdf_url}")
                                # Download the PDF immediately
                                completed = row["Completed"]
                                target_dir = make_report_folder_path(completed)
                                target_name = report_filename(row)
                                save_path = os.path.join(target_dir, target_name)
                                
                                # Check if PDF already exists on disk
                                if os.path.exists(save_path):
                                    logging.info(f"PDF exists on disk, stamping timestamp")
                                else:
                                    r = requests.get(pdf_url, stream=True)
                                    if r.status_code == 200:
                                        with open(save_path, 'wb') as f:
                                            for chunk in r.iter_content(10240):
                                                f.write(chunk)
                                        logging.info(f"PDF downloaded and saved to {save_path}")
                                        self.stats.pdfs_downloaded += 1
                                    else:
                                        logging.warning(f"Failed to download PDF for idx={idx}. Status: {r.status_code}")
                                        continue
                                
                                dl_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                result_df.at[idx, "PDF report downloaded"] = dl_time
                                all_current_rows[idx]["PDF report downloaded"] = dl_time
                            else:
                                logging.warning("No PDF file name found! Skipping.")
                            
                            # Incremental save after each PDF processed (silent to avoid log spam)
                            row_year = datetime.strptime(row["Completed"], "%d/%m/%Y").year
                            self._save_year_to_excel(row_year, rows_by_year, silent=True)

                            # Navigate back to results tab (stay on same page)
                            driver.switch_to.default_content()
                            results_tab = safe_find(driver, By.XPATH, "//a[@href='#TestAdministration/Results']")
                            results_tab.click()
                            time.sleep(3)
                            switch_to_results_iframe(driver)
                            time.sleep(2)  # Wait for iframe to load, we're still on the same page
                        except Exception as e:
                            logging.error(f"Error processing row idx={idx}: {e}")
                            self.stats.errors_encountered += 1
                            try:
                                driver.switch_to.default_content()
                                goto_results_tab(driver)
                                switch_to_results_iframe(driver)
                                time.sleep(2)
                            except Exception as ex:
                                logging.error(f"Failed to recover after error: {ex}")
                                break
                    
                    # Move to next page (if not last)
                    if page_num < total_pages:
                        if not click_next_page(driver):
                            logging.warning(f"Failed to navigate to page {page_num + 1}")
                            break

                # Final save for all years (logs row counts)
                for year in rows_by_year:
                    self._save_year_to_excel(year, rows_by_year)

                logging.info(f"All done for account {username}.")
                self.stats.accounts_processed += 1
            except Exception as e:
                logging.error(f"Error processing account {username}: {e}")
                self.stats.errors_encountered += 1
            finally:
                driver.quit()
                logging.info("Chrome closed for this account.\n")
        return self.stats

def main():
    """Main entry point."""
    gui = EvolveGUI()
    gui.run()

if __name__ == "__main__":
    main()