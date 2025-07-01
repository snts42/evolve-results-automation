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
- Comprehensive logging and error handling for auditability and troubleshooting.

## Usage

There are **two scripts** included for modularity and ease of deployment:

1. **`evolve_scraper.py`**  
   - Logs in, navigates, and scrapes all exam results and PDF report URLs.
   - Outputs everything to `exam_results.xlsx`.

2. **`reports_downloader.py`**  
   - Reads the Excel, downloads all linked candidate reports, and organises them in folders by year/month/day.
   - Updates the Excel to track downloaded files.

> **Note:**  
> At present, the report scraping and download process is still **semi-manual for testing**:  
> The tool will prompt the user to press Enter at various steps (such as confirming page loads or after certain browser actions).  
> This ensures stability during development and troubleshooting.  
> A fully automated, "headless" version is planned for future releases, once all edge cases and are resolved.

### Steps

1. **Clone the repository.**  
2. **Install dependencies.**  
3. **Download and place ChromeDriver** (matching your Chrome version) in the project root.  
4. **Prepare credentials:**  
   Edit **`credentials.json`** and enter your Evolve username and password.  
5. **Run `evolve_scraper.py`** to extract results and links.  
6. **Run `reports_downloader.py`** to download all reports.

## Dependencies

- Python 3.7 or later (tested on 3.13.3)
- [Selenium](https://pypi.org/project/selenium/)
- [pandas](https://pypi.org/project/pandas/)
- [requests](https://pypi.org/project/requests/)
- [openpyxl](https://pypi.org/project/openpyxl/)
- Chrome + [ChromeDriver](https://sites.google.com/chromium.org/driver/)

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