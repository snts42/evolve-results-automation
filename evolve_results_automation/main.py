import os
import sys
import logging
import pandas as pd
from datetime import datetime

from evolve_results_automation.config import (
    CHROME_DRIVER_PATH, ENCRYPTED_CREDENTIALS_FILE, EXCEL_FILE, REPORTS_BASE, LOG_FILE
)
from evolve_results_automation.excel_utils import (
    initialize_excel, load_existing_results, save_results, autofilter_and_autofit
)
from evolve_results_automation.selenium_utils import (
    start_driver, login, goto_results_tab, switch_to_results_iframe,
    reset_and_refresh, parse_results_table, select_table_row,
    click_candidate_report_button, extract_pdf_filename_from_html, safe_find
)
from selenium.webdriver.common.by import By
from evolve_results_automation.secure_credentials import load_secure_credentials
from evolve_results_automation.gui import EvolveGUI

import requests

from evolve_results_automation.logging_utils import setup_logger, log
setup_logger()


def make_report_folder_path(date_str: str) -> str:
    dt = datetime.strptime(date_str, "%d/%m/%Y")
    year = dt.strftime("%Y")
    month_day = dt.strftime("%m %d")
    folder = os.path.join(REPORTS_BASE, year, month_day)
    os.makedirs(folder, exist_ok=True)
    return folder

def report_filename(row):
    parts = [
        str(row["First name"]).strip(),
        str(row["Last name"]).strip(),
        str(row["Test Name"]).strip(),
        str(row["Result"]).strip()
    ]
    fname = " ".join(parts) + ".pdf"
    fname = "".join(c for c in fname if c not in r'\/:*?"<>|')
    return fname

def unique_row_hash(row):
    fields = [
        "Candidate ref.", "First name", "Last name", "Completed", "Test Name", "Result"
    ]
    return "|".join([str(row.get(f, "")).strip().lower() for f in fields])

from dataclasses import dataclass
from typing import List, Dict

@dataclass
class ProcessingStats:
    accounts_processed: int = 0
    new_rows_added: int = 0
    pdfs_downloaded: int = 0
    errors_encountered: int = 0

class EvolveAutomation:
    def __init__(self, headless: bool, master_password: str):
        self.headless = headless
        self.master_password = master_password
        self.stats = ProcessingStats()
        self.columns = [
            "Candidate ref.", "First name", "Last name", "Completed",
            "Test Name", "Result", "Percent", "Duration", "Centre Name",
            "Report URL", "Report Download", "Result Sent", "Result Sent By",
            "E-Certificate", "E-Certificate By", "Certificate", "Certificate By",
            "Comments", "Keycode", "Subject", "PDF Direct Link"
        ]
        self.col_map = {
            "Downloaded At": "Report URL",
            "Report Downloaded At": "Report Download",
            "E-Certificate Sent": "E-Certificate",
            "Certificate Issued": "Certificate"
        }

    def run(self):
        initialize_excel(EXCEL_FILE, self.columns)
        existing_df = load_existing_results(EXCEL_FILE)
        if not existing_df.empty:
            existing_df.rename(columns=self.col_map, inplace=True)
            save_results(EXCEL_FILE, existing_df)

        accounts = load_secure_credentials(ENCRYPTED_CREDENTIALS_FILE, master_password=self.master_password)
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
                # Deduplicate hashes across ALL accounts so far
                result_df = load_existing_results(EXCEL_FILE)
                result_df.rename(columns=self.col_map, inplace=True)
                existing_hashes = set(unique_row_hash(row) for _, row in result_df.iterrows()) if not result_df.empty else set()
                new_rows = parse_results_table(driver, existing_hashes, self.col_map)
                if new_rows:
                    new_df = pd.DataFrame(new_rows)
                    result_df = pd.concat([result_df, new_df], ignore_index=True)
                    save_results(EXCEL_FILE, result_df)
                    logging.info(f"Added {len(new_rows)} new rows to Excel.")
                    self.stats.new_rows_added += len(new_rows)
                else:
                    logging.info("No new rows found.")

                # Scrape PDF links and download+update Excel as soon as found
                pdf_rows = result_df[
                    (result_df["PDF Direct Link"].isnull()) | 
                    (result_df["PDF Direct Link"] == "") | 
                    (result_df["PDF Direct Link"] == "NOT FOUND")
                ]
                
                for idx, row in pdf_rows.iterrows():
                    try:
                        logging.info(f"Processing for PDF link: {row['First name']} {row['Last name']} ({row['Test Name']})")
                        select_table_row(driver, row)
                        click_candidate_report_button(driver)
                        driver.implicitly_wait(6)
                        html = driver.page_source
                        file_name = extract_pdf_filename_from_html(html)
                        if file_name:
                            pdf_url = f"https://evolve.cityandguilds.com/secureassess/CustomerData/Evolve/DocumentStore/{file_name}"
                            logging.info(f"PDF direct link: {pdf_url}")
                            result_df.at[idx, "PDF Direct Link"] = pdf_url
                            # Download the PDF immediately
                            completed = row["Completed"]
                            target_dir = make_report_folder_path(completed)
                            target_name = report_filename(row)
                            save_path = os.path.join(target_dir, target_name)
                            r = requests.get(pdf_url, stream=True)
                            if r.status_code == 200:
                                with open(save_path, 'wb') as f:
                                    for chunk in r.iter_content(10240):
                                        f.write(chunk)
                                logging.info(f"PDF downloaded and saved to {save_path}")
                                result_df.at[idx, "Report Download"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                self.stats.pdfs_downloaded += 1
                            else:
                                logging.warning(f"Failed to download PDF for idx={idx}. Status: {r.status_code}")
                            # Update Excel immediately after each PDF process
                            save_results(EXCEL_FILE, result_df)
                        else:
                            logging.warning("No PDF file name found! Skipping.")
                            result_df.at[idx, "PDF Direct Link"] = "NOT FOUND"
                            save_results(EXCEL_FILE, result_df)
                        driver.switch_to.default_content()
                        # Reopen results tab/iframe after report
                        results_tab = safe_find(driver, By.XPATH, "//a[@href='#TestAdministration/Results']")
                        results_tab.click()
                        driver.implicitly_wait(6)
                        switch_to_results_iframe(driver)
                        reset_and_refresh(driver)
                    except Exception as e:
                        logging.error(f"Error processing row idx={idx}: {e}")
                        self.stats.errors_encountered += 1
                        try:
                            driver.switch_to.default_content()
                            goto_results_tab(driver)
                            switch_to_results_iframe(driver)
                        except Exception as ex:
                            logging.error(f"Failed to recover after error: {ex}")
                            break

                # Optional: Rearranging and sorting after processing
                if not result_df.empty:
                    for c in ["Keycode", "Subject", "PDF Direct Link"]:
                        if c in result_df.columns:
                            result_df = pd.concat([result_df.drop([c], axis=1), result_df[[c]]], axis=1)
                    save_results(EXCEL_FILE, result_df)

                if "Completed" in result_df.columns:
                    try:
                        result_df["Completed_sort"] = pd.to_datetime(result_df["Completed"], format="%d/%m/%Y", errors="coerce")
                        result_df = result_df.sort_values(by="Completed_sort", ascending=True).drop(columns="Completed_sort")
                        save_results(EXCEL_FILE, result_df)
                    except Exception as e:
                        logging.warning(f"Sorting failed: {e}")

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
    import os
    from colorama import Fore, Style
    import getpass
    from evolve_results_automation.config import ENCRYPTED_CREDENTIALS_FILE, EXCEL_FILE
    gui = EvolveGUI()

    # --- File permission and lock checks ---
    files_to_check = [(EXCEL_FILE, 'Excel results file'),
                      (ENCRYPTED_CREDENTIALS_FILE, 'Encrypted credentials file')]
    for path, label in files_to_check:
        if os.path.exists(path):
            try:
                with open(path, 'a+b'):
                    pass
            except Exception as e:
                print(Fore.RED + f"\nERROR: Cannot access {label} ({path}): {e}" + Style.RESET_ALL)
                print(Fore.YELLOW + f"Please close any programs using this file and try again." + Style.RESET_ALL)
                input(Fore.CYAN + "\nPress Enter to continue..." + Style.RESET_ALL)
                sys.exit(1)

    gui.show_banner()
    print(Fore.CYAN + "\nIMPORTANT: If you forget your master password, you can delete the 'credentials.enc' file and restart the program to set a new master password. This will " + Fore.RED + "ERASE all saved credentials." + Style.RESET_ALL)
    print(Fore.GREEN + "="*100 + Style.RESET_ALL)

    master_password = None

    # Onboarding: If encrypted credentials file does not exist, prompt to add credentials and set master password
    if not os.path.exists(ENCRYPTED_CREDENTIALS_FILE):
        gui.show_credentials_menu()  # This will allow user to add credentials and set master password
        # After adding, check again
        if not os.path.exists(ENCRYPTED_CREDENTIALS_FILE):
            print(Fore.RED + "No credentials were added. Exiting." + Style.RESET_ALL)
            sys.exit(1)
    else:
        # Loop until valid master password is entered
        from evolve_results_automation.secure_credentials import SecureCredentialManager
        manager = SecureCredentialManager(ENCRYPTED_CREDENTIALS_FILE)
        while True:
            master_password = getpass.getpass(Fore.YELLOW + "Enter master password: " + Style.RESET_ALL)
            try:
                # Try to decrypt credentials to validate password
                manager.decrypt_credentials(master_password)
                break  # Success!
            except Exception as e:
                print(Fore.RED + f"\n❌ Invalid master password: {e}" + Style.RESET_ALL)
        gui.master_password = master_password

    try:
        while True:
            choice = gui.show_main_menu(master_password=master_password)
            if choice == 'run_automation':
                headless = gui.setup_automation(master_password=master_password)
                if headless is not None:
                    automation = EvolveAutomation(headless, master_password)
                    stats = automation.run()
                    gui.show_summary(stats)
                    print(Fore.CYAN + "\nPress Enter to return to main menu..." + Style.RESET_ALL)
                    input()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\nFatal error: {e}")
        logging.error(f"Fatal error in main: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()