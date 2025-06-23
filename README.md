# Evolve Results Automation

## Introduction

This is an **unofficial, open-source tool** to automate the retrieval and download of exam results and candidate reports from the City & Guilds Evolve platform, which lacks official API support.  
It is designed to assist centre/admin staff in downloading all candidate results and reports in a consistent, auditable, and efficient manner—saving hours of manual work and reducing human error.

**This tool is intended for internal/admin use only.**  
It operates by automating browser actions and is for use by those with legitimate access to the Evolve platform. Never run this tool on untrusted machines. The project respects privacy, your data, and the Terms of Service of the sites you access.

## Features

- Logs in and navigates to the Evolve Results section.
- Downloads all result data into an Excel spreadsheet (`exam_results.xlsx`).
- (Work in progress) Downloads all candidate reports as PDFs into an organized folder structure.
- Comprehensive logging and error handling for auditability and troubleshooting.

## Usage

1. **Clone the repository:**  

2. **Install dependencies:** 

3. **Download and place ChromeDriver:**  
Ensure your version of `chromedriver.exe` matches your Chrome browser and is placed in the project root.

4. **Prepare credentials:**  
Edit `credentials.json` and enter your Evolve username and password.

5. **Run the script:**  

The script will prompt you for manual steps as needed. Automated mode can be added in future updates.

## Configuration

- Configurable parameters are at the top of `evolve_scraper.py`.
- Results are saved in `exam_results.xlsx`.
- Candidate reports are downloaded into the `reports/` directory, organized by date.

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

---