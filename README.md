# E-volve SecureAssess Automation

An open-source desktop tool for exam centre staff to automate the download of candidate results and PDF reports from the City & Guilds E-volve SecureAssess platform.

![Windows](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/python-3.13%2B-yellow)
![License](https://img.shields.io/badge/license-MIT-green)

## Features

- **Desktop GUI** built with CustomTkinter (light mode, City & Guilds branding)
- **Automated scraping** of the E-volve Results table via Selenium (headless or visible browser)
- **PDF report downloads** with direct links extracted from the candidate report page
- **Multi-account support** with per-account automation
- **AES-256 encrypted credentials** protected by a master password
- **Year-based folder organisation** for Excel results, PDF reports, and logs
- **Pagination support** across multi-page result sets
- **Date filtering** automatically set to the 1st of the previous month
- **PDF resume** across runs (skips already-downloaded reports)
- **Automatic ChromeDriver management** via Selenium Manager

## Quick Start

### Download the .exe (recommended)

1. Download `EvolveResultsAutomation.exe` from the [latest release](https://github.com/snts42/evolve-results-automation/releases/latest)
2. Place it in a dedicated folder (data files are saved alongside the exe)
3. Run it and set a master password on first launch
4. Open **Settings** to add your E-volve SecureAssess login(s)
5. Click **Run Automation**

> **Requires:** Google Chrome installed. ChromeDriver is downloaded automatically.

### Run from source

```bash
git clone https://github.com/snts42/evolve-results-automation.git
cd evolve-results-automation
pip install -r requirements.txt
python -m evolve_results_automation
```

## Output Structure

All data is organised by exam completion year:

```
EvolveResultsAutomation.exe
credentials.enc
2026/
  exam_results.xlsx
  reports/
    03 15/
      FirstName_LastName_TestName.pdf
  logs/
    03 15/
      log_2026-03-15_09-30-00.txt
```

## Security

- Credentials are encrypted with **AES-256** using a master password you set on first launch
- No plaintext credentials are ever written to disk
- The master password is prompted once per session
- If you forget your master password, delete `credentials.enc` and restart to set a new one

## Dependencies

All managed via `requirements.txt`:

- **selenium** - browser automation
- **pandas** - data processing
- **requests** - PDF downloads
- **openpyxl** - Excel read/write
- **cryptography** - AES-256 credential encryption
- **customtkinter** - desktop GUI framework
- **colorama** - terminal colour output

## License

MIT. See [LICENSE](LICENSE.md) for details.

---

**Disclaimer:** This is an unofficial tool and is not affiliated with, endorsed by, or associated with City & Guilds. E-volve and SecureAssess are trademarks of The City and Guilds of London Institute.

**Author:** Alex Santonastaso | [santonastaso.codes](https://santonastaso.codes)