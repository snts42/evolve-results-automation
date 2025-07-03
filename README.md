# Evolve Results Automation

## Introduction

This is an **unofficial, open-source tool** to automate the retrieval and download of exam results and candidate reports from the City & Guilds Evolve platform, which lacks official API support.  
It is designed to assist centre/admin staff in downloading all candidate results and reports in a consistent, auditable, and efficient manner—saving hours of manual work and reducing human error.

**This tool is intended for internal/admin use only.**  
It operates by automating browser actions and is for use by those with legitimate access to the Evolve platform. Never run this tool on untrusted machines. The project respects privacy, your data, and the Terms of Service of the sites you access.

## Features

- Logs in and navigates to the Evolve Results section.
- Downloads all result data into an Excel spreadsheet (`exam_results.xlsx`).
- Extracts and saves direct PDF links for candidate reports.
- Downloads all candidate reports as PDFs into an organised folder structure by date.
- Supports multiple Evolve accounts (multi-credential).
- Automatically logs each session with timestamped log files in a structured folder.
- Comprehensive error handling for reliability and auditability.
- Excel file is auto-filtered, columns auto-sized, and data is always kept up to date.

## Which Script Should I Use?

- **`evolve_scraper+downloader.py`**  
  The main, fully-automated script—handles scraping results and downloading PDF reports in one go.  
  **Use this for normal operation.**

- **`evolve_scraper.py`** & **`reports_downloader.py`**  
  These are legacy and development scripts, originally used for separate scraping and downloading steps.  
  They lack full automation, multi-account support, advanced logging, Excel formatting, and other new features.  
  Provided for reference or testing only.

## Usage Overview

1. Clone/download the repository.
2. Install Python dependencies (see below).
3. Place `chromedriver.exe` (matching your Chrome version) in the project root.
4. Add your credentials in `credentials.json` (supports one or more accounts).
5. Run `evolve_scraper+downloader.py`.
6. Results will be saved to `exam_results.xlsx`, PDF reports in `reports/YYYY/MM DD/`, and logs in `logs/YYYY/MM DD/`.

## Dependencies

- Python 3.7 or later (tested on 3.13.3)
- Selenium
- pandas
- requests
- openpyxl
- Chrome + ChromeDriver

## Security, Disclaimer, and Responsible Use

- **For internal/admin use only.**
- Use with your own credentials; never share or publish your credentials file.
- The authors and contributors are not affiliated with City & Guilds or Evolve.  
- **You are solely responsible** for use, compliance, and any consequences.
- This tool is an assistant and does not replace your professional judgment or compliance obligations.
- All actions are logged for your audit and troubleshooting.

## License

MIT. See [LICENSE](LICENSE) for details.

---

**Author:**  
Alex Santonastaso | [santonastaso.codes](https://santonastaso.codes)