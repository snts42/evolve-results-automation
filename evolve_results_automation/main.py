import os
import logging
from dataclasses import dataclass
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

from evolve_results_automation.config import (
    APP_VER, ENCRYPTED_CREDENTIALS_FILE, DOCUMENT_STORE_URL
)
from evolve_results_automation.excel_utils import (
    save_year_to_excel, load_all_existing_data
)
from evolve_results_automation.selenium_utils import (
    start_driver, login, switch_to_results_iframe,
    reset_and_refresh, parse_results_table, select_table_row,
    click_candidate_report_button, get_total_pages,
    click_next_page, navigate_to_results, set_date_filter_to_previous_month_start
)

from evolve_results_automation.secure_credentials import SecureCredentialManager

from evolve_results_automation.logging_utils import setup_logger
from evolve_results_automation.parsing_utils import make_report_folder_path, report_filename, extract_pdf_filename_from_html, unique_row_hash

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
        self._driver = None

    def run(self):
        """Top-level entry point: setup, decrypt accounts, iterate."""
        setup_logger()
        mode = "headless" if self.headless else "non-headless (browser visible)"
        logging.info(f"Evolve Results Automation {APP_VER} | Mode: {mode}")
        accounts = SecureCredentialManager(ENCRYPTED_CREDENTIALS_FILE).decrypt_credentials(self.master_password)

        # Filter accounts if specific username selected
        if self.selected_username:
            accounts = [acc for acc in accounts if acc.get("username", "").strip() == self.selected_username]
            if not accounts:
                logging.error(f"Selected account '{self.selected_username}' not found in credentials")
                return self.stats

        if not accounts:
            logging.warning("No accounts found in credentials")
            return self.stats

        account_names = ", ".join(a.get("username", "?").strip() for a in accounts)
        logging.info(f"Processing {len(accounts)} account(s): {account_names}")

        account_reports = []
        for idx_acc, account in enumerate(accounts):
            username = account.get("username", "").strip()
            password = account.get("password", "").strip()
            if not username or not password:
                logging.warning(f"Credentials missing, skipping account #{idx_acc+1}")
                continue
            logging.info(f"--- Starting for account #{idx_acc+1}: {username} ---")
            # Snapshot stats before this account
            rows_before = self.stats.new_rows_added
            pdfs_before = self.stats.pdfs_downloaded
            errs_before = self.stats.errors_encountered
            driver = start_driver(headless=self.headless)
            self._driver = driver
            try:
                self._process_account(driver, username, password)
                self.stats.accounts_processed += 1
            except Exception as e:
                logging.error(f"Error processing account {username}: {e}")
                self.stats.errors_encountered += 1
            finally:
                self._driver = None
                try:
                    driver.quit()
                except Exception:
                    pass
                acct_rows = self.stats.new_rows_added - rows_before
                acct_pdfs = self.stats.pdfs_downloaded - pdfs_before
                acct_errs = self.stats.errors_encountered - errs_before
                account_reports.append((username, acct_rows, acct_pdfs, acct_errs))
                logging.info(f"Account {username}: {acct_rows} new results, {acct_pdfs} PDFs, {acct_errs} error(s)")
                logging.info("Chrome closed for this account\n")
        # Final summary (include per-account breakdown if multiple accounts)
        if len(account_reports) > 1:
            logging.info("--- Run Summary ---")
            for name, rows, pdfs, errs in account_reports:
                logging.info(f"  {name}: {rows} new results, {pdfs} PDFs, {errs} error(s)")
        logging.info(f"Run complete: {self.stats.accounts_processed} account(s), "
                     f"{self.stats.new_rows_added} new results, "
                     f"{self.stats.pdfs_downloaded} PDFs downloaded, "
                     f"{self.stats.errors_encountered} error(s)")
        return self.stats

    def _process_account(self, driver, username, password):
        """Login, set filters, scrape all pages, and download PDFs for one account."""
        login(driver, username, password)
        switch_to_results_iframe(driver)
        reset_and_refresh(driver)

        # Set date filter start to 1st of previous month
        if not set_date_filter_to_previous_month_start(driver):
            logging.warning("Date filter may not have been set correctly")

        # Get total pages
        total_pages = get_total_pages(driver)
        logging.info(f"Found {total_pages} page(s) to scrape")

        # Build hash set from all existing year Excel files to prevent duplicates
        existing_hashes, rows_by_year, _ = load_all_existing_data()

        # Process each page
        for page_num in range(1, total_pages + 1):
            logging.info(f"Processing page {page_num}/{total_pages}")

            # Scrape table and group new rows by year
            self._scrape_page(driver, page_num, existing_hashes, rows_by_year)

            # Download PDFs for all rows that still need them
            self._process_page_pdfs(driver, rows_by_year)

            # Move to next page (if not last) - MUST be at end of loop
            if page_num < total_pages:
                if not click_next_page(driver):
                    logging.warning(f"Failed to navigate to page {page_num + 1}")
                    break

        # Final save for all years (logs row counts)
        for year in rows_by_year:
            save_year_to_excel(year, rows_by_year)

        logging.info(f"All done for account {username}")

    def _scrape_page(self, driver, page_num, existing_hashes, rows_by_year):
        """Scrape the current results page and group new rows by year."""
        new_rows = parse_results_table(driver, existing_hashes)
        if new_rows:
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
                save_year_to_excel(yr, rows_by_year)
        else:
            logging.info(f"No new rows found on page {page_num}")

    def _process_page_pdfs(self, driver, rows_by_year):
        """Download PDFs for all rows across all years that still need them."""
        pdf_needed = [
            row for year_rows in rows_by_year.values()
            for row in year_rows
            if not row.get("PDF report save time")
        ]

        if not pdf_needed:
            return

        for row in pdf_needed:
            try:
                completed = str(row.get("Completed", "")).strip()
                if not completed:
                    continue

                # Try to select the row - if it's not on this page, skip it
                if not select_table_row(driver, row):
                    continue

                logging.info(f"Processing PDF for: {row['First name']} {row['Last name']} ({row['Test Name']})")
                click_candidate_report_button(driver)
                html = driver.page_source
                file_name = extract_pdf_filename_from_html(html)
                if file_name:
                    pdf_url = f"{DOCUMENT_STORE_URL}{file_name}"
                    downloaded = self._download_pdf(pdf_url, row, completed)

                    if downloaded:
                        dl_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        row["PDF report save time"] = dl_time
                        row_year = datetime.strptime(row["Completed"], "%d/%m/%Y").year
                        save_year_to_excel(row_year, rows_by_year, silent=True)
                else:
                    logging.warning("No PDF file name found, skipping")

                navigate_to_results(driver)
            except Exception as e:
                logging.error(f"Error processing PDF for {row.get('First name', '?')} {row.get('Last name', '?')}: {e}")
                self.stats.errors_encountered += 1
                try:
                    navigate_to_results(driver)
                except Exception as ex:
                    logging.error(f"Failed to recover after error: {ex}")
                    break

    def _download_pdf(self, pdf_url, row, completed):
        """Download a single PDF. Returns True if downloaded or already on disk."""
        target_dir = make_report_folder_path(completed)
        target_name = report_filename(row)
        save_path = os.path.join(target_dir, target_name)

        if os.path.exists(save_path) and os.path.getsize(save_path) > 0:
            logging.info(f"PDF already on disk, updating timestamp")
            return True

        try:
            resp = urlopen(Request(pdf_url), timeout=30)
            try:
                with open(save_path, 'wb') as f:
                    while True:
                        chunk = resp.read(10240)
                        if not chunk:
                            break
                        f.write(chunk)
                logging.info(f"PDF saved: {target_name}")
                self.stats.pdfs_downloaded += 1
                return True
            finally:
                resp.close()
        except HTTPError as e:
            logging.warning(f"Failed to download PDF, status: {e.code}")
            return False
        except (URLError, OSError) as e:
            logging.warning(f"Failed to download PDF: {e}")
            return False