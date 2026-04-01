# Evolve Results Automation

## Introduction

**Evolve Results Automation** is an open-source tool for exam centre staff to securely automate the download of candidate results and reports from the City & Guilds Evolve platform, which lacks official API support.

- Designed for internal/admin use only
- Automates browser actions (Selenium)
- All credentials are stored securely using AES-256 encryption
- Provides a user-friendly .exe for non-technical users (no Python required)

## Features

- Logs into Evolve and navigates to the Results section
- Downloads all results data into an Excel spreadsheet (`exam_results.xlsx`)
- Extracts and saves direct PDF links for candidate reports
- Downloads all candidate reports as PDFs into a structured folder by date
- Supports multiple Evolve accounts (multi-credential access)
- Robust error handling and logging
- Encrypted credential management via a modern desktop GUI (CustomTkinter)

## Quick Start (.exe Users)

1. Download the latest `.exe` from [Releases](https://github.com/snts42/evolve-results-automation/releases/)
2. Run the `.exe` — a desktop GUI will open
3. Set a master password on first launch, then add your Evolve credentials in the **Accounts** tab
4. Click **Start Automation** from the **Dashboard** tab
5. Results are saved to `YYYY/exam_results.xlsx`, PDFs to `YYYY/reports/MM DD/`, logs to `YYYY/logs/MM DD/`

> **Note:** ChromeDriver is managed automatically. Just make sure Google Chrome is installed.

> **If you forget your master password:**
> Delete the `credentials.enc` file and restart the program to re-onboard.

## For Python Users / Developers

1. Clone/download this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run:
   ```bash
   python -m evolve_results_automation
   ```

## Dependencies

- Python >= 3.13.3
- selenium >= 4.33.0
- pandas >= 2.3.0
- requests >= 2.32.3
- openpyxl >= 3.1.5
- cryptography >= 45.0.5
- customtkinter >= 5.2.2
- Pillow >= 10.0.0
- colorama >= 0.4.6

## Security & Credential Management

- **All credentials are stored encrypted** (AES-256, master password required)
- No plaintext credentials are ever written to disk
- You must remember your master password; if lost, delete `credentials.enc` to reset
- Master password is prompted only once per session
- No legacy `.json` plain-text credentials logic remains

## ChromeDriver

ChromeDriver is managed automatically by Selenium Manager (built into Selenium 4.6+). You do **not** need to download or update ChromeDriver manually — it is resolved automatically based on your installed Chrome version.

If you encounter issues starting Chrome, ensure:
- Google Chrome is installed on your machine
- You have internet access (for the initial ChromeDriver download)

## Versioning & Release Notes

- **Current release:** v1.0.0
- Modern desktop GUI (CustomTkinter), encrypted credentials, `.exe` support, automatic ChromeDriver management
- Year-based folder organisation for Excel results, PDF reports, and logs
- All previous `.py` scripts and plaintext credential logic are obsolete

## License

MIT. See [LICENSE](LICENSE) for details.

---

**Author:**  
Alex Santonastaso | [santonastaso.codes](https://santonastaso.codes)