import logging
from dataclasses import dataclass
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.error import URLError

from evolve_results_automation.config import (
    APP_VER, ENCRYPTED_CREDENTIALS_FILE, DOCUMENT_STORE_URL, RESULTS_URL
)
from evolve_results_automation.excel_utils import (
    save_year_to_excel, load_all_existing_data, regenerate_analytics
)
from evolve_results_automation.selenium_utils import (
    start_driver, login, switch_to_results_iframe,
    reset_and_refresh, parse_results_table, select_table_row,
    click_candidate_report_button, get_total_pages,
    click_next_page, navigate_to_results, set_date_filter,
    handle_duplicate_page
)
from evolve_results_automation.secure_credentials import SecureCredentialManager
from evolve_results_automation.logging_utils import setup_logger
from evolve_results_automation.parsing_utils import (
    extract_pdf_filename_from_html, unique_row_hash, download_pdf
)

@dataclass
class ProcessingStats:
    accounts_processed: int = 0
    new_rows_added: int = 0
    pdfs_downloaded: int = 0
    errors_encountered: int = 0
    pdfs_skipped: int = 0


def compute_pdf_cutoff_date(months_back: int) -> datetime:
    total_back = months_back + 1
    now = datetime.now()
    month = now.month - total_back
    cutoff_year = now.year + (month - 1) // 12
    cutoff_month = (month - 1) % 12 + 1
    return datetime(cutoff_year, cutoff_month, 1)

class EvolveAutomation:
    def __init__(self, headless: bool, master_password: str, selected_username: str = None, stop_event=None,
                 months_back: int = 1, skip_pdfs: bool = False, scheduled: bool = False):
        self.headless = headless
        self.master_password = master_password
        self.selected_username = selected_username
        self.stats = ProcessingStats()
        self._driver = None
        self._stop_event = stop_event
        self._months_back = months_back
        self._skip_pdfs = skip_pdfs
        self._scheduled = scheduled

    def run(self):
        """Top-level entry point: setup, decrypt accounts, iterate."""
        setup_logger()
        mode = "(browser not visible)" if self.headless else "browser visible"
        trigger = "scheduled" if self._scheduled else "manual"
        logging.info(f"Evolve Results Automation {APP_VER} | {mode} | {trigger} run")
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

        self._preflight_check()

        account_reports = []
        for idx_acc, account in enumerate(accounts):
            if self._stop_event and self._stop_event.is_set():
                logging.info("Automation stopped by user")
                break
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

            max_attempts = 2
            for attempt in range(1, max_attempts + 1):
                if self._stop_event and self._stop_event.is_set():
                    break
                driver = start_driver(headless=self.headless)
                self._driver = driver
                try:
                    self._process_account(driver, username, password)
                    self.stats.accounts_processed += 1
                    break  # success - no retry needed
                except Exception as e:
                    if attempt < max_attempts:
                        logging.warning(f"Attempt {attempt} failed for {username}: {e}")
                        logging.info(f"Retrying with a fresh browser...")
                    else:
                        logging.error(f"Error processing account {username} (attempt {attempt}/{max_attempts}): {e}")
                        self.stats.errors_encountered += 1
                finally:
                    self._driver = None
                    try:
                        driver.quit()
                        logging.info("Chrome closed")
                    except Exception:
                        pass

            acct_rows = self.stats.new_rows_added - rows_before
            acct_pdfs = self.stats.pdfs_downloaded - pdfs_before
            acct_errs = self.stats.errors_encountered - errs_before
            account_reports.append((username, acct_rows, acct_pdfs, acct_errs))
            logging.info(f"Account {username}: {acct_rows} new results, {acct_pdfs} PDFs, {acct_errs} error(s)")
        # Final summary (include per-account breakdown if multiple accounts)
        if len(account_reports) > 1:
            logging.info("--- Run Summary ---")
            for name, rows, pdfs, errs in account_reports:
                logging.info(f"  {name}: {rows} new results, {pdfs} PDFs, {errs} error(s)")
        # Regenerate analytics once after all accounts (runs even on early stop)
        if self.stats.new_rows_added > 0 or self.stats.pdfs_downloaded > 0:
            logging.info("Updating analytics...")
            regenerate_analytics()

        logging.info(f"Run complete: {self.stats.accounts_processed} account(s), "
                     f"{self.stats.new_rows_added} new results, "
                     f"{self.stats.pdfs_downloaded} PDFs downloaded, "
                     f"{self.stats.errors_encountered} error(s)")
        return self.stats

    def _preflight_check(self):
        """Test connectivity to E-volve before starting automation."""
        try:
            resp = urlopen(Request(RESULTS_URL.split('#')[0],
                                   headers={"User-Agent": "Mozilla/5.0"}), timeout=10)
            resp.close()
        except (URLError, OSError) as e:
            raise ConnectionError(
                f"Cannot reach E-volve. Check your internet connection. ({e})") from e

    def _process_account(self, driver, username, password):
        """Login, set filters, scrape all pages, and download PDFs for one account."""
        login(driver, username, password)

        if self._stop_event and self._stop_event.is_set():
            logging.info("Automation stopped by user")
            return

        switch_to_results_iframe(driver)
        reset_and_refresh(driver)

        if self._stop_event and self._stop_event.is_set():
            logging.info("Automation stopped by user")
            return

        existing_hashes, rows_by_year, pdf_resume_count = load_all_existing_data(silent=True)

        if not set_date_filter(driver, self._months_back):
            logging.warning("Date filter may not have been set correctly")

        total_pages = get_total_pages(driver)
        pdf_note = f" ({pdf_resume_count} pending PDF reports)" if pdf_resume_count > 0 and not self._skip_pdfs else ""
        logging.info(f"Loaded {len(existing_hashes)} previous results, {total_pages} page(s) to check{pdf_note}")

        prev_page_hashes = set()
        for page_num in range(1, total_pages + 1):
            if self._stop_event and self._stop_event.is_set():
                logging.info("Automation stopped by user - saving progress")
                break
            logging.info(f"Checking page {page_num}/{total_pages}")

            page_hashes = self._scrape_page(driver, page_num, existing_hashes, rows_by_year)
            scrape_fn = lambda d, p: self._scrape_page(d, p, existing_hashes, rows_by_year)
            page_hashes = handle_duplicate_page(
                driver, page_num, page_hashes, prev_page_hashes, scrape_fn)
            if page_hashes is None:
                break
            prev_page_hashes = page_hashes

            if not self._skip_pdfs:
                self._process_page_pdfs(driver, rows_by_year)

            # Move to next page (if not last) - MUST be at end of loop
            if page_num < total_pages:
                if not click_next_page(driver):
                    logging.warning(f"Failed to navigate to page {page_num + 1}, skipping remaining pages")
                    self.stats.errors_encountered += 1
                    break

        for year in rows_by_year:
            save_year_to_excel(year, rows_by_year, silent=True)
        logging.info(f"Finished account {username}")

    def _scrape_page(self, driver, page_num, existing_hashes, rows_by_year):
        """Scrape the current results page and group new rows by year.
        Returns the set of page hashes for duplicate page detection."""
        new_rows, page_hashes = parse_results_table(driver, existing_hashes)
        if new_rows:
            # Track how many new rows go to each year
            new_per_year = {}
            for row in new_rows:
                completed_date = row.get("Completed", "")
                try:
                    year = datetime.strptime(completed_date, "%d/%m/%Y").year
                except (ValueError, TypeError):
                    year = datetime.now().year
                    logging.warning(f"Unparseable date '{completed_date}' for {row.get('First name', '?')} {row.get('Last name', '?')}, defaulting to {year}")

                if year not in rows_by_year:
                    rows_by_year[year] = []
                rows_by_year[year].append(row)
                existing_hashes.add(unique_row_hash(row))
                new_per_year[year] = new_per_year.get(year, 0) + 1

            logging.info(f"Found {len(new_rows)} new result(s) on page {page_num}")
            self.stats.new_rows_added += len(new_rows)

            # Save scraped rows to Excel immediately (before PDF processing)
            for yr in sorted(new_per_year.keys(), reverse=True):
                save_year_to_excel(yr, rows_by_year, silent=True)
                logging.info(f"  Saved {new_per_year[yr]} new result(s) to {yr}/exam_results_{yr}.xlsx")
        else:
            logging.info(f"No new results on page {page_num}")
        return page_hashes

    def _process_page_pdfs(self, driver, rows_by_year):
        """Download PDFs for all rows across all years that still need them."""
        # Filter to rows within the current date filter range to avoid
        # trying to select rows not visible in the table.
        # The E-volve calendar opens one month behind current, then navigates
        # back months_back more, so the effective range is months_back + 1.
        cutoff = compute_pdf_cutoff_date(self._months_back)

        pdf_needed = []
        skipped = 0
        for year_rows in rows_by_year.values():
            for row in year_rows:
                if row.get("PDF report save time"):
                    continue
                try:
                    dt = datetime.strptime(row.get("Completed", ""), "%d/%m/%Y")
                    if dt < cutoff:
                        skipped += 1
                        continue
                except (ValueError, TypeError):
                    continue
                pdf_needed.append(row)

        if skipped:
            logging.info(f"Skipped {skipped} PDF(s) outside current date range")
        if not pdf_needed:
            return

        for row in pdf_needed:
            if self._stop_event and self._stop_event.is_set():
                break
            try:
                completed = str(row.get("Completed", "")).strip()
                if not completed:
                    continue

                # Try to select the row - if it's not on this page, skip it
                if not select_table_row(driver, row):
                    continue

                result = row.get('Result', '').strip()
                logging.info(f"Downloading PDF: {row['First name']} {row['Last name']} ({row['Test Name']}) - {result} ({completed})")
                click_candidate_report_button(driver)
                html = driver.page_source
                file_name = extract_pdf_filename_from_html(html)
                if file_name:
                    pdf_url = f"{DOCUMENT_STORE_URL}{file_name}"
                    downloaded = download_pdf(pdf_url, row, completed)

                    if downloaded:
                        self.stats.pdfs_downloaded += 1
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
                    remaining = sum(
                        1 for r in pdf_needed
                        if not r.get("PDF report save time"))
                    if remaining:
                        self.stats.pdfs_skipped += remaining
                        logging.warning(f"{remaining} PDF(s) skipped due to unrecoverable navigation error")
                    break
