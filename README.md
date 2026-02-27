# TimesheetsProcessor

A modular Python application that automates the processing of employee timesheet attachments from Gmail. It downloads, parses, and organizes timesheet data into structured Excel workbooks.

## Features

- **Job-1-Download**: Searches Gmail for emails with timesheet attachments from employees, downloads them into organized per-employee folders, and writes metadata CSVs
- **Job-2-Parse**: Parses downloaded timesheet files (PDF, PNG, XLSX, ZIP) using text extraction and OCR, writes structured data to per-employee Excel workbooks with one tab per year
- **Job-3-MoveToProcessedTag**: Moves processed emails in Gmail from inbox to a custom label, closing the processing loop

### Supported Timesheet Formats

- Oracle HCM
- PeopleSoft HRCMS
- MBO Partners
- Unanet Time-List
- Unanet EaZyTyme
- Calendar grid view
- Techno-Comp Excel (.xlsx)
- ZIP archives containing any of the above

## Prerequisites

- Python 3.10+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) (for image-based timesheet parsing)
- Google Cloud project with Gmail API enabled
- OAuth2 credentials for Gmail API

### Install Tesseract

```bash
# macOS
brew install tesseract

# Ubuntu/Debian
sudo apt-get install tesseract-ocr

# Windows
# Download from https://github.com/UB-Mannheim/tesseract/wiki
```

## Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/YOUR_USERNAME/TimesheetsProcessor.git
   cd TimesheetsProcessor
   ```

2. **Install Python dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up Gmail API credentials**
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a project and enable the Gmail API
   - Create OAuth 2.0 credentials (Desktop application)
   - Download the credentials JSON file
   - Save it as `config/credentials.json`
   - See `config/credentials.json.example` for the expected format

4. **Configure the application**
   - Edit `config/config.yaml` to set your date range and preferences
   - Edit `config/EmployeesList.md` to add your employees (name, email, project)

## Usage

```bash
# Step 1: Download timesheet attachments from Gmail
python -m src.jobs.Job_1_Download

# Step 2: Parse downloaded timesheets and write to Excel
python -m src.jobs.Job_2_Parse

# Step 2 (alt): Re-process previously unparseable files
python -m src.jobs.Job_2_Parse --reparse

# Step 3: Move processed emails to "claudeprocessed" Gmail label
python -m src.jobs.Job_3_MoveToProcessedTag
```

## Project Structure

```
TimesheetsProcessor/
├── config/
│   ├── config.yaml                  # Main configuration
│   ├── EmployeesList.md             # Employee list (name, email, project)
│   ├── credentials.json             # Google OAuth2 credentials (not tracked)
│   └── credentials.json.example     # Example credentials format
├── src/
│   ├── common/                      # Module #0 — shared foundation
│   │   ├── config_loader.py         # Config file parser
│   │   ├── logger_setup.py          # Logging with rotation
│   │   ├── employee_parser.py       # Employee list parser
│   │   ├── gmail_auth.py            # OAuth2 authentication
│   │   ├── gmail_client.py          # Gmail API wrapper
│   │   ├── file_manager.py          # File/folder operations
│   │   ├── timesheet_parser.py      # PDF/Image/XLSX parsing
│   │   └── excel_writer.py          # Excel workbook creation
│   └── jobs/                        # Job executables
│       ├── Job_1_Download.py        # Download attachments from Gmail
│       ├── Job_2_Parse.py           # Parse timesheet files
│       └── Job_3_MoveToProcessedTag.py  # Move processed emails
├── plans/
│   └── implementation-plan.md       # Detailed technical specification
├── requirements.txt
├── .gitignore
└── MyPlan.md                        # Project requirements
```

## How It Works

1. **Configuration**: Reads `config/config.yaml` for date ranges, batch sizes, and paths
2. **Employee List**: Parses `config/EmployeesList.md` (Markdown table) for employee names, emails, and projects
3. **Gmail Search**: Uses Gmail API to search for emails from each employee with attachments within the configured date range
4. **Download**: Saves attachments to per-employee folders with batch metadata CSVs
5. **Parsing**: Extracts text from PDF/images using pdfplumber and Tesseract OCR, then uses regex patterns to find date ranges and hours
6. **Excel Output**: Creates/updates per-employee Excel workbooks with one tab per year, prepopulated with weekly or monthly date ranges
7. **Label Management**: Moves processed emails to a "claudeprocessed" Gmail label (never deletes)

## License

MIT License
