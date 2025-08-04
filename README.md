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
- Encrypted credential management  via CLI or GUI menu

## Quick Start (.exe Users)

1. Download the latest `.exe` from [Releases](https://github.com/snts42/evolve-results-automation/releases/) 
2. Place `chromedriver.exe` (matching your installed Chrome version) in the same folder as the `.exe`
3. Run the `.exe` and follow the prompts to set a master password and add your Evolve credentials
4. Start the automation from the menu
5. Results will be saved to `exam_results.xlsx`, PDFs in `reports/YYYY/MM DD/`, and logs in `logs/YYYY/MM DD/`

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
   python -m evolve_results_automation.main
   ```

## Dependencies

- Python >= 3.13.3
- selenium >= 4.33.0
- pandas >= 2.3.0
- requests >= 2.32.3
- openpyxl >= 3.1.5
- colorama >= 0.4.6
- cryptography >= 45.0.5

## Security & Credential Management

- **All credentials are stored encrypted** (AES-256, master password required)
- No plaintext credentials are ever written to disk
- You must remember your master password; if lost, delete `credentials.enc` to reset
- Master password is prompted only once per session
- No legacy `.json` plain-text credentials logic remains

## ChromeDriver Troubleshooting

- You **must** use a `chromedriver.exe` version that matches your installed Google Chrome version
- If the automation fails to start ChromeDriver, you will see a clear error message
- Download the correct ChromeDriver from: https://chromedriver.chromium.org/downloads

## Versioning & Release Notes

- **Current release:** v0.2.0
- This is a major security and usability update (encrypted credentials, .exe support, no legacy code)
- All previous `.py` scripts and plaintext credential logic are obsolete

## License

MIT. See [LICENSE](LICENSE) for details.

---

**Author:**  
Alex Santonastaso | [santonastaso.codes](https://santonastaso.codes)