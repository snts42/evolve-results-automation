import os
import sys
import logging
import pandas as pd
from datetime import datetime
from colorama import init, Fore, Style

from evolve_results_automation.config import (
    CHROME_DRIVER_PATH, CREDENTIALS_FILE, EXCEL_FILE, REPORTS_BASE, LOG_FILE
)
from evolve_results_automation.excel_utils import (
    initialize_excel, load_existing_results, save_results, autofilter_and_autofit
)
from evolve_results_automation.selenium_utils import (
    start_driver, login, goto_results_tab, switch_to_results_iframe,
    reset_and_refresh, parse_results_table, select_table_row,
    click_candidate_report_button, extract_pdf_filename_from_html, safe_find
)
from evolve_results_automation.credentials_utils import load_credentials

import requests

from evolve_results_automation.logging_utils import setup_logger, log
setup_logger()

def pre_run_check():
    from colorama import init, Fore, Style
    import os, sys

    init(autoreset=True)
    ascii_logo = f"""{Fore.GREEN}{Style.BRIGHT}
 _____ _   _ _____ _     _   _ _____    ___  _   _ _____ ________  ___ ___ _____ _____ _____ _   _ 
|  ___| | | |  _  | |   | | | |  ___|  / _ \| | | |_   _|  _  |  \/  |/ _ \_   _|_   _|  _  | \ | |
| |__ | | | | | | | |   | | | | |__   / /_\ \ | | | | | | | | | .  . / /_\ \| |   | | | | | |  \| |
|  __|| | | | | | | |   | | | |  __|  |  _  | | | | | | | | | | |\/| |  _  || |   | | | | | | . ` |
| |___\ \_/ | \_/ / |___\ \_/ / |___  | | | | |_| | | | \ \_/ / |  | | | | || |  _| |_\ \_/ / |\  |
\____/ \___/ \___/\_____/\___/\____/  \_| |_/\___/  \_/  \___/\_|  |_|_| |_/\_/  \___/ \___/\_| \_/
{Style.RESET_ALL}
{Fore.WHITE}{Style.BRIGHT}
                  Evolve Results Automation Tool by snts42
{Style.RESET_ALL}
"""
    print(ascii_logo)
    print(Fore.GREEN + "=" * 100)
    print(Fore.WHITE + "WELCOME TO THE EVOLVE RESULTS AUTOMATION TOOL!")
    print(Fore.GREEN + "=" * 100 + Style.RESET_ALL)
    print(
        Fore.GREEN +
        "Automate the retrieval and download of exam results\n"
        "from the City & Guilds Evolve platform."
        f"\n{Fore.GREEN}Make sure your {Fore.WHITE}'credentials.json'{Fore.GREEN} file is present and up to date."
        f"\nIf not, copy {Fore.WHITE}'credentials_example.json'{Fore.GREEN} and edit it.\n"
    )
    if not os.path.exists(CREDENTIALS_FILE):
        print(
            Fore.RED
            + f"\nERROR: '{CREDENTIALS_FILE}' not found. Please create this file before running the tool."
        )
        input(Fore.CYAN + "Press Enter to exit..." + Style.RESET_ALL)
        sys.exit(1)

    else:
        try:
            accounts = load_credentials(CREDENTIALS_FILE)
            print(Fore.GREEN + f"\nFound '{CREDENTIALS_FILE}' with {len(accounts)} credential(s):")
            for idx, acc in enumerate(accounts, 1):
                user = acc.get('username', '(missing username)')
                print(Fore.CYAN + f"  [{idx}] {user}")
        except Exception as e:
            print(Fore.RED + f"\nERROR: Unable to load credentials: {e}")
            input(Fore.CYAN + "Press Enter to exit..." + Style.RESET_ALL)
            sys.exit(1)
    
    while True:
        headless_input = input(
            Fore.GREEN + "\nRun in HEADLESS (no browser window) mode? [Y/n]: " + Style.RESET_ALL
        ).strip().lower()
        if headless_input in ("y", "yes"):
            headless = True
            break
        elif headless_input in ("n", "no"):
            headless = False
            break
        else:
            print(Fore.YELLOW + "Please enter 'Y' or 'N'." + Style.RESET_ALL)
    return headless

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

def main():
    headless = pre_run_check()

    columns = [
        "Candidate ref.", "First name", "Last name", "Completed",
        "Test Name", "Result", "Percent", "Duration", "Centre Name",
        "Report URL", "Report Download", "Result Sent", "Result Sent By",
        "E-Certificate", "E-Certificate By", "Certificate", "Certificate By",
        "Comments", "Keycode", "Subject", "PDF Direct Link"
    ]
    col_map = {
        "Downloaded At": "Report URL",
        "Report Downloaded At": "Report Download",
        "E-Certificate Sent": "E-Certificate",
        "Certificate Issued": "Certificate"
    }

    initialize_excel(EXCEL_FILE, columns)
    existing_df = load_existing_results(EXCEL_FILE)

    if not existing_df.empty:
        existing_df.rename(columns=col_map, inplace=True)
        save_results(EXCEL_FILE, existing_df)

    accounts = load_credentials(CREDENTIALS_FILE)

    for idx_acc, account in enumerate(accounts):
        username = account.get("username", "").strip()
        password = account.get("password", "").strip()
        if not username or not password:
            logging.warning(f"Credentials missing for account #{idx_acc+1}, skipping this account.")
            continue

        logging.info(f"--- Starting for account #{idx_acc+1}: {username} ---")
        driver = start_driver(headless=headless)
        try:
            login(driver, username, password)
            goto_results_tab(driver)
            switch_to_results_iframe(driver)
            reset_and_refresh(driver)

            # Deduplicate hashes across ALL accounts so far
            result_df = load_existing_results(EXCEL_FILE)
            result_df.rename(columns=col_map, inplace=True)
            existing_hashes = set(unique_row_hash(row) for _, row in result_df.iterrows()) if not result_df.empty else set()

            new_rows = parse_results_table(driver, existing_hashes, col_map)
            if new_rows:
                new_df = pd.DataFrame(new_rows)
                result_df = pd.concat([result_df, new_df], ignore_index=True)
                save_results(EXCEL_FILE, result_df)
                logging.info(f"Added {len(new_rows)} new rows to Excel.")
            else:
                logging.info("No new rows found.")

            # Scrape PDF links and download+update Excel as soon as found
            for idx, row in result_df[
                (result_df["PDF Direct Link"].isnull()) | 
                (result_df["PDF Direct Link"] == "") | 
                (result_df["PDF Direct Link"] == "NOT FOUND")
            ].iterrows():
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
                    try:
                        driver.switch_to.default_content()
                        goto_results_tab(driver)
                        switch_to_results_iframe(driver)
                    except Exception as ex:
                        logging.error(f"Failed to recover after error: {ex}")
                        break

            # Optional: Rearranging and sorting after processing (as before)
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
        finally:
            driver.quit()
            logging.info("Chrome closed for this account.\n")

    print(Fore.CYAN + "\nAll done! Press Enter to exit..." + Style.RESET_ALL)
    input()

if __name__ == "__main__":
    main()