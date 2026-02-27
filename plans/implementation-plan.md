# Timesheets Processor - Implementation Plan

> **Version**: 1.3 | **Date**: 2026-02-26

---

## Changelog

| Version | Change | Description |
|---|---|---|
| v1.01 | **Modular executables** | Each major job is now a separate standalone executable script, not a single monolithic `main.py` |
| v1.01 | **Module #0 (Init)** | Config, logging, employee parsing, Gmail auth, and file management are shared foundation modules used by all jobs |
| v1.01 | **Job-1-Download** | The download workflow is named `Job-1-Download` and lives in its own executable |
| v1.01 | **Project structure** | Reorganized into `src/common/` (shared #0 modules) and `src/jobs/` (per-job executables) |
| v1.01 | **Future-proof** | Architecture supports adding Job-2, Job-3, etc. as independent executables that reuse common modules |
| v1.02 | ~~Employees subfolder~~ | ~~Added employees/ intermediary subfolder~~ — **Reverted in v1.03** |
| v1.03 | **Direct folder creation** | Employee folders are created **directly** under `EMPLOYEES_PARENT_DIR_PATH` — no `employees/` intermediary. Path: `{EMPLOYEES_PARENT_DIR_PATH}/{EmployeeName_ProjectName}/`. Removed `employees_dir` from Config dataclass. |
| v1.04 | **Job-2-Parse** | New job: parse downloaded timesheet files (PDF/PNG), extract structured data (hours, dates, project), write to per-employee Excel files with one tab per year, prepopulated with weekly/monthly date ranges. Unparseable files moved to `unableToParse/` folder. |
| v1.21 | **Post-parse file management** | After successfully parsing a timesheet file, rename it with start and end dates (e.g., `2025-03-02_2025-03-08_original.pdf`) and move from `downloaded/` to `processed/` folder. Unparseable files still go to `unableToParse/`. |
| v1.22 | **Project/client extraction** | Extract project/client identifier from timesheet content. For PeopleSoft timesheets: use the category line (e.g., "VLO Category 2B", "Full Service Cat1"). For Oracle HCM timesheets: no project field available in text — field left empty. Extracted project written to the "Project/Client" column in Excel. Supports employees working on multiple projects simultaneously (different timesheets may have different project values). |
| v1.3 | **Enhanced parser + new formats** | Added PyMuPDF OCR fallback for image-based PDFs, upgraded OCR preprocessing (2x upscale + adaptive threshold), added support for 6 new timesheet formats (MBO Partners, Unanet Time-List, Unanet EaZyTyme, PeopleSoft via OCR, calendar grid view, Techno-Comp .xlsx). Added .xlsx and .zip file support. Added `--reparse` flag to reprocess `unableToParse/` files. Email signature images (<300px) auto-skipped. OCR error tolerance for common misreads. |
| v1.3 | **Job-3-MoveToProcessedTag** | New job: loop through employee folders, read download CSVs, move processed emails from inbox to "claudeprocessed" Gmail label/tag using Gmail API. Never deletes emails. After all emails in a CSV are moved, the CSV itself is moved to `processed/` folder. |

---

## Table of Contents

1. [Project Overview](#1-project-overview)
2. [Architecture: Modular Executables](#2-architecture-modular-executables)
3. [Technology Stack](#3-technology-stack)
4. [Project Structure](#4-project-structure)
5. [Configuration Design](#5-configuration-design)
6. [Authentication (Gmail API)](#6-authentication-gmail-api)
7. [Module #0 — Common/Init Breakdown](#7-module-0--commoninit-breakdown)
8. [Job-1-Download Breakdown](#8-job-1-download-breakdown)
9. [Detailed Flow: Job-1-Download](#9-detailed-flow-job-1-download)
10. [Job-2-Parse Breakdown](#10-job-2-parse-breakdown)
11. [Detailed Flow: Job-2-Parse](#11-detailed-flow-job-2-parse)
12. [Job-3-MoveToProcessedTag Breakdown](#12-job-3-movetoprocessedtag-breakdown)
13. [Detailed Flow: Job-3-MoveToProcessedTag](#13-detailed-flow-job-3-movetoprocessedtag)
14. [Error Handling Strategy](#14-error-handling-strategy)
15. [Logging Strategy](#15-logging-strategy)
16. [Implementation Phases](#16-implementation-phases)
17. [Critical Gotchas & Constraints](#17-critical-gotchas--constraints)

---

## 1. Project Overview

Build a modular Python application where each major job is a **separate executable**. All jobs share a common initialization module (#0) for configuration, logging, authentication, and file management.

Current jobs:
- **Job-1-Download**: Search Gmail for employee timesheet emails, download attachments into organized folders, write metadata CSVs
- **Job-2-Parse**: Parse downloaded timesheet files (PDF/PNG/XLSX/ZIP), extract structured data (hours, dates, project), write results to per-employee Excel files. Supports `--reparse` mode for re-processing failed files.
- **Job-3-MoveToProcessedTag**: Move processed emails in Gmail from inbox to "claudeprocessed" label/tag, then move download CSVs to processed/

**Hard rule**: NEVER delete emails. Gmail API scope uses **modify** (needed by Job-3 to move emails to labels) but emails are never deleted.

---

## 2. Architecture: Modular Executables

```
                    ┌──────────────────────────────────┐
                    │       Module #0 (Common/Init)     │
                    │                                   │
                    │  config_loader.py                  │
                    │  logger_setup.py                   │
                    │  employee_parser.py                │
                    │  gmail_auth.py                     │
                    │  gmail_client.py                   │
                    │  file_manager.py                   │
                    │  timesheet_parser.py               │
                    │  excel_writer.py                   │
                    └──┬──────────┬──────────┬──────────┘
                       │          │          │
          ┌────────────▼──┐  ┌────▼────────┐  ┌──▼──────────────────┐
          │ Job-1-Download │  │ Job-2-Parse │  │ Job-3-MoveToProcessed│
          │ (executable)   │  │ (executable)│  │ Tag (executable)     │
          └────────────────┘  └─────────────┘  └──────────────────────┘
```

**Key principle**: Each job is a standalone script that:
1. Imports from `src.common.*` for shared functionality
2. Has its own `main()` function with job-specific orchestration
3. Can be run independently: `python -m src.jobs.Job_1_Download`

---

## 3. Technology Stack

| Component | Choice | Reason |
|---|---|---|
| Language | Python 3.10+ | Gmail API client library, widely supported |
| Gmail API | `google-api-python-client` | Official Google SDK |
| Auth | `google-auth`, `google-auth-oauthlib` | OAuth2 flow for Gmail |
| Config | `PyYAML` | Read `config.yaml` |
| Logging | `logging` + `RotatingFileHandler` | Built-in, rotation support |
| CSV | `csv` (stdlib) | Write download metadata |
| PDF parsing | `pdfplumber` | Best table extraction, preserves layout, built on pdfminer.six |
| PDF rendering | `pymupdf` (fitz) | Render image-based PDF pages to images at 300 DPI for OCR fallback |
| Image OCR | `pytesseract` + `Pillow` | Text extraction from PNG timesheet images |
| Image preprocessing | `opencv-python` | 2x upscale + adaptive Gaussian threshold for better OCR |
| Excel read/write | `openpyxl` | Read/write .xlsx with multiple tabs, styling, date formatting |

### `requirements.txt`
```
google-api-python-client>=2.100.0
google-auth>=2.23.0
google-auth-oauthlib>=1.1.0
google-auth-httplib2>=0.1.1
pyyaml>=6.0.1
pdfplumber>=0.10.0
pytesseract>=0.3.10
Pillow>=10.0.0
opencv-python>=4.8.0
openpyxl>=3.1.0
pymupdf>=1.23.0
```

### System Dependency (OCR)
```bash
# macOS
brew install tesseract
```

---

## 4. Project Structure

```
TimesheetsProcessor/
├── config/
│   ├── config.yaml                  # Main configuration
│   ├── EmployeesList.md             # Default employee list
│   └── credentials.json             # Google OAuth2 credentials (gitignored)
├── logs/
│   └── timesheets_processor.log     # Auto-created, rotated
├── JohnSmith_ProjectA/                 # Employee folders created directly under EMPLOYEES_PARENT_DIR_PATH
│   ├── downloaded/                  # PROCESSING_FOLDER_NAME — raw attachments from Gmail
│   │   ├── <msgId>_timesheet.pdf
│   │   └── <msgId>_IMG_1234.png
│   ├── processed/                   # Successfully parsed files, renamed with date ranges
│   │   └── 2025-03-02_2025-03-08_<msgId>_timesheet.pdf
│   ├── unableToParse/               # Files that Job-2-Parse could not extract hours from
│   │   └── <msgId>_corrupted.pdf
│   ├── JohnSmith_ProjectA.xlsx         # Structured timesheet Excel (one tab per year)
│   ├── download-1.csv               # Batch run #1 metadata
│   └── download-2.csv               # Batch run #2 metadata
├── JaneDoe_ProjectBeta/
│   └── ...
├── src/
│   ├── __init__.py
│   ├── common/                      # ── Module #0 (shared by all jobs) ──
│   │   ├── __init__.py              #   Exports init_app() convenience function
│   │   ├── config_loader.py         #   Read & validate config.yaml
│   │   ├── employee_parser.py       #   Parse EmployeesList.md
│   │   ├── gmail_auth.py            #   OAuth2 authentication
│   │   ├── gmail_client.py          #   Search emails, download attachments
│   │   ├── file_manager.py          #   Folder creation, sanitization, CSV
│   │   ├── logger_setup.py          #   Logging with rotation
│   │   ├── timesheet_parser.py      #   PDF text extraction + OCR for images
│   │   └── excel_writer.py          #   Excel workbook creation, year tabs, date ranges
│   └── jobs/                        # ── Job executables ──
│       ├── __init__.py
│       ├── Job_1_Download.py        #   Job-1-Download: email attachment downloader
│       ├── Job_2_Parse.py           #   Job-2-Parse: parse timesheets → Excel
│       └── Job_3_MoveToProcessedTag.py  #   Job-3: move processed emails to Gmail label
├── plans/
│   └── implementation-plan.md       # This file
├── token.pickle                     # Auto-generated OAuth token (gitignored)
├── requirements.txt
├── .gitignore
└── MyPlan.md                        # Original requirements
```

### Running Jobs

```bash
# Run Job-1-Download (search Gmail, download attachments)
python -m src.jobs.Job_1_Download

# Run Job-2-Parse (parse downloaded timesheets, write Excel)
python -m src.jobs.Job_2_Parse

# Run Job-2-Parse in reparse mode (re-process unableToParse/ files)
python -m src.jobs.Job_2_Parse --reparse

# Run Job-3-MoveToProcessedTag (move emails to "claudeprocessed" label)
python -m src.jobs.Job_3_MoveToProcessedTag
```

---

## 5. Configuration Design

### 5.1 `config/config.yaml`

```yaml
# Timesheets Processor Configuration
START_DATE: "12-01-2024"                       # MM-DD-YYYY (required)
END_DATE: ""                                   # MM-DD-YYYY (defaults to current date)
EMPLOYEES_PARENT_DIR_PATH: ""                  # Defaults to working directory
EMPLOYEES_LIST_FILE: ""                        # Defaults to config/EmployeesList.md
EMAIL_BATCH_SIZE: 5                            # Number of emails to retrieve at a time
PROCESSING_FOLDER_NAME: "downloaded"           # Subfolder for attachments

# Gmail OAuth
GMAIL_CREDENTIALS_FILE: "config/credentials.json"
GMAIL_TOKEN_FILE: "token.pickle"
GMAIL_USER_ID: "me"                            # 'me' = authenticated user
```

### 5.2 `config/EmployeesList.md`

Expected format (Markdown table):

```markdown
| Employee Name | Employee Email         | Project Name   |
|---------------|------------------------|----------------|
| John Smith | john.smith@example.com | ProjectA    |
```

### 5.3 Config Loader Logic (`config_loader.py`)

```
1. Read config/config.yaml using yaml.safe_load()
2. Parse START_DATE as MM-DD-YYYY → datetime object (required, fail if missing)
3. Parse END_DATE as MM-DD-YYYY → datetime object (default: today)
4. EMPLOYEES_PARENT_DIR_PATH → default to os.getcwd() (employee folders created directly here)
5. EMPLOYEES_LIST_FILE → default to "config/EmployeesList.md"
6. EMAIL_BATCH_SIZE → default to 5, must be positive integer
7. PROCESSING_FOLDER_NAME → default to "downloaded"
8. Return a validated Config dataclass
```

---

## 6. Authentication (Gmail API)

### 6.1 Scope

```python
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
```

**Modify scope** is required by Job-3-MoveToProcessedTag to add/remove labels on emails (move to "claudeprocessed" label). This scope does NOT allow permanent deletion. Emails are NEVER deleted — only moved between labels.

> **Note**: If upgrading from a previous read-only token, delete `token.pickle` and re-authenticate to get the new scope.

### 6.2 OAuth2 Flow (`gmail_auth.py`)

```
1. Check if token.pickle exists and is valid
2. If expired → refresh using refresh_token
3. If no token → run InstalledAppFlow.run_local_server()
   (opens browser for user consent on first run)
4. Save refreshed/new token to token.pickle
5. Return authenticated gmail service object
```

### 6.3 Pre-requisites (Manual Setup)

Before first run, user must:
1. Go to Google Cloud Console → create project
2. Enable Gmail API
3. Create OAuth 2.0 Client ID (Desktop application)
4. Download `credentials.json` → place in `config/credentials.json`
5. First run will open browser for consent → generates `token.pickle`

---

## 7. Module #0 — Common/Init Breakdown

All modules in `src/common/` are shared infrastructure. Every job imports from here.

### 7.1 `logger_setup.py` — Logging

- Creates `logs/` directory if not exists
- Configures `RotatingFileHandler`: 5 MB max, 7 backup files
- Console handler for INFO+, file handler for DEBUG+
- Format: `%(asctime)s | %(levelname)-8s | %(module)s | %(message)s`

### 7.2 `config_loader.py` — Configuration

- Reads and validates `config/config.yaml`
- Returns a `Config` dataclass with typed fields
- Validates date formats (MM-DD-YYYY)
- Resolves default values for optional params

### 7.3 `employee_parser.py` — Employee List Parser

- Reads the Markdown file specified by `EMPLOYEES_LIST_FILE`
- Parses the Markdown table rows
- Returns list of `Employee` dataclasses:
  ```python
  @dataclass
  class Employee:
      name: str           # "John Smith"
      email: str          # "john.smith@example.com"
      project: str        # "__"
      folder_name: str    # "JohnSmith_ProjectA" (sanitized)
  ```
- Sanitization: strip spaces, dots, special chars from `name + "_" + project`

### 7.4 `gmail_auth.py` — Authentication

- OAuth2 flow as described in Section 6
- Returns `googleapiclient.discovery.Resource` (Gmail service)

### 7.5 `gmail_client.py` — Gmail Operations

**`search_emails(service, sender_email, start_date, end_date, batch_size)`**
- Builds Gmail query: `from:{email} after:{YYYY/MM/DD} before:{YYYY/MM/DD} has:attachment`
- Calls `users().messages().list()` with `maxResults=batch_size`
- Returns list of message ID stubs (respects EMAIL_BATCH_SIZE limit)
- **Note**: Gmail date query uses `YYYY/MM/DD` format (slashes, not dashes)
- **Note**: `before:` is exclusive, so add 1 day to END_DATE for inclusive range

**`get_message_details(service, message_id)`**
- Calls `users().messages().get()` with `format='full'`
- Extracts: message ID, internalDate (ms → datetime), attachment parts
- Returns structured metadata dict

**`download_attachment(service, message_id, attachment_id, filename, output_dir)`**
- Calls `users().messages().attachments().get()`
- Decodes with `base64.urlsafe_b64decode()` (NOT standard b64)
- Saves to `output_dir/{message_id}_{filename}`
- Returns saved file path and size

**`_get_attachment_parts(payload)`**
- Recursively walks MIME parts tree
- Finds parts with `filename` and `attachmentId`
- Returns list of attachment metadata dicts

### 7.6 `file_manager.py` — File & Folder Operations

**`ensure_employee_folders(parent_dir, folder_name, processing_folder)`**
- Creates `{parent_dir}/{folder_name}/` if not exists (e.g. `JohnSmith_ProjectA/`)
- Creates `{parent_dir}/{folder_name}/{processing_folder}/` if not exists
- Returns both paths

**`sanitize_folder_name(name, project)`** (lives in employee_parser.py)
- Concatenates name + "_" + project
- Removes spaces, dots, and special characters (keep alphanumeric + underscore)

**`get_next_batch_number(employee_folder)`**
- Scans for existing `download-*.csv` files
- Returns next integer (1 if none exist)

**`write_download_csv(employee_folder, batch_number, records)`**
- Creates/writes `download-{batch_number}.csv`
- Columns: `gmail_message_id, num_attachments, email_datetime`
- Each row = one email processed

### 7.7 `common/__init__.py` — Convenience Init

Provides an `init_app()` function that every job calls to bootstrap:

```python
def init_app():
    """
    Standard initialization for all jobs.
    Returns (logger, config, employees, gmail_service).
    """
    logger = setup_logging()
    config = load_config()
    employees = parse_employees(config.employees_list_file)
    service = authenticate(config.gmail_credentials_file, config.gmail_token_file)
    return logger, config, employees, service
```

This ensures consistent startup across all jobs while keeping each job's code focused on its own logic.

### 7.8 `timesheet_parser.py` — PDF, Image & Excel Text Extraction

Extracts raw text from timesheet files so Job-2-Parse can parse structured data. Supports 8 distinct timesheet formats across PDF, image, Excel, and ZIP files.

**Supported Extensions**: `.pdf`, `.png`, `.jpg`, `.jpeg`, `.xlsx`, `.zip`

**`extract_text(file_path)`**
- Routes to the appropriate extractor based on file extension
- `.pdf` → `_extract_text_from_pdf()`, `.png/.jpg/.jpeg` → `_extract_text_from_image()`
- `.xlsx` and `.zip` → returns empty string (handled separately by dedicated parsers)

**`_extract_text_from_pdf(file_path)`**
- Primary: Uses `pdfplumber` to extract text from each page
- **OCR fallback**: If pdfplumber returns empty text (image-based PDFs), falls back to `_ocr_pdf_pages()` which renders each page at 300 DPI via PyMuPDF (fitz) and runs OCR

**`_ocr_pdf_pages(file_path)`**
- Opens PDF with PyMuPDF (`fitz`)
- Renders each page to a pixmap at 300 DPI
- Converts to PIL Image → runs `_preprocess_for_ocr()` + `pytesseract`
- Returns combined text from all pages

**`_extract_text_from_image(file_path)`**
- **Size filter**: Images smaller than 300px in width or height are skipped (returns empty string) — filters out email signature logos (LinkedIn, BBB, NMSDC)
- Preprocesses with `_preprocess_for_ocr()` and runs `pytesseract` with PSM 6

**`_preprocess_for_ocr(img)`**
- 2x upscale with cubic interpolation for better OCR accuracy
- Converts to grayscale
- Adaptive Gaussian threshold (block size 31, C=10)
- Returns preprocessed PIL Image

**`parse_xlsx_timesheet(file_path)`**
- Directly parses Techno-Comp style Excel timesheets (cell-based extraction)
- Reads employee name from row 7, period dates from row 13
- Reads weekly hours from rows 18 and 24 (Sun-Sat columns)
- Returns `TimesheetFileData` with entries

**`extract_from_zip(zip_path, dest_dir)`**
- Extracts supported files from ZIP archives
- Skips macOS metadata (`__MACOSX/`, `._` prefixed files) and temp files (`~$`)
- Returns list of extracted file paths

**`parse_all_entries(text)`**
- Parses all timesheet entries from raw text (handles multi-entry files)
- Supports 8 timesheet formats:
  - **Oracle HCM Cloud**: `DD-Mon-YYYY - DD-Mon-YYYY` date ranges with `Total Hours:XX.XX`
  - **PeopleSoft HRCMS**: `From Sunday MM/DD/YYYY to Saturday MM/DD/YYYY` with `Reported Hours XX.XXX`
  - **MBO Partners**: `Timesheet for Mon DD - Mon DD, YYYY` with standalone `Total XX.XX`
  - **Unanet Time-List**: `MM/DD/YYYY — MM/DD/YYYY  40  LOCKED` inline format
  - **Unanet EaZyTyme**: `(MM/DD/YYYY - MM/DD/YYYY)` in title with `Totals: 40.00`
  - **Calendar grid view**: Monthly calendar with weekly totals at row ends
  - **Techno-Comp Excel**: `.xlsx` files parsed directly via `parse_xlsx_timesheet()`
  - **ZIP archives**: Extracted via `extract_from_zip()`, inner files parsed by format
- **Hours search chain** (in order of precedence): `Total Hours` → `Reported Hours` → `Totals:` → inline hours after dates → `Total XX.XX`
- **Hours sanity check**: Rejects hours > 168 (weekly) or > 744 (monthly) to filter OCR-corrupted values
- **OCR error tolerance**: Handles common OCR misreads:
  - `F` dropped from "From" (`F?rom` pattern)
  - `'` instead of `/` in dates (normalized before parsing)
  - `§` junk characters between keywords and values (`\W*` tolerant)
  - `S` read as `8`, `<` read as `4` in calendar grids
- Returns a `TimesheetFileData` dataclass:
  ```python
  @dataclass
  class TimesheetEntry:
      start_date: date
      end_date: date
      hours: float
      period_type: str            # "weekly" or "monthly"
      status: Optional[str]       # "Approved", "Submitted", etc.

  @dataclass
  class TimesheetFileData:
      employee_name: Optional[str]
      project: Optional[str]      # Extracted project/client identifier
      entries: List[TimesheetEntry]
      raw_text: str
  ```

**`_parse_calendar_view(text)`**
- Detects month+year header (e.g., "June 2025" or "June2025")
- Pre-cleans OCR errors: `S`→`8`, `<N`→`4N`
- Extracts daily hours from rows with 8+ numbers (7 days + weekly total)
- Builds week boundaries from the calendar month
- Returns list of `TimesheetEntry` objects

### 7.9 `excel_writer.py` — Excel Workbook Management

Manages per-employee Excel files with year-based tabs.

**`get_or_create_workbook(filepath)`**
- Loads existing `.xlsx` or creates new `Workbook`

**`get_or_create_year_sheet(workbook, year, mode)`**
- Gets existing tab named `"{year}"` or creates and prepopulates it
- `mode="weekly"`: rows are Sun-Sat week ranges for entire year
- `mode="monthly"`: rows are calendar months (Jan-Dec)

**`get_weekly_ranges(year)`**
- Generates all Sun-Sat week ranges that overlap with the given year
- Returns list of `(week_start_date, week_end_date)` tuples

**`get_monthly_ranges(year)`**
- Generates `(first_day, last_day)` for each month Jan-Dec
- Returns 12 tuples

**Weekly sheet layout:**

| Week Start (Sun) | Week End (Sat) | Hours Worked | Project/Client | Notes |
|---|---|---|---|---|
| 12/29/2024 | 01/04/2025 | | | |
| 01/05/2025 | 01/11/2025 | 40 | ProjectAlpha | |
| ... | ... | | | |

**Monthly sheet layout:**

| Month | Start Date | End Date | Hours Worked | Project/Client | Notes |
|---|---|---|---|---|---|
| January | 01/01/2025 | 01/31/2025 | 160 | ProjectAlpha | |
| February | 02/01/2025 | 02/28/2025 | | | |
| ... | ... | ... | | | |

**`update_weekly_row(worksheet, week_start_date, hours, project)`**
- Finds the row where column A matches `week_start_date`
- Updates Hours Worked and Project/Client columns
- Returns `True` if match found, `False` if not

**`update_monthly_row(worksheet, year, month, hours, project)`**
- Finds the row for the given month in the year tab
- Updates Hours Worked and Project/Client columns

**`save_workbook(workbook, filepath)`**
- Sorts tabs chronologically by year
- Saves to disk

---

## 8. Job-1-Download Breakdown

### File: `src/jobs/Job_1_Download.py`

**Purpose**: Search Gmail for emails with attachments from each employee, download attachments to organized folders, write metadata CSV.

**Run command**: `python -m src.jobs.Job_1_Download`

### Pseudocode

```
Job_1_Download.main():
    1. Call init_app() → get logger, config, employees, gmail_service
    2. Log: "Job-1-Download started"
    3. For each employee:
        a. Log: "Processing employee: {name} ({email})"
        b. Search Gmail for matching emails (from, dates, has:attachment)
        c. If no emails found → log & skip
        d. Ensure employee folder + processing subfolder exist
        e. Determine next batch number
        f. Initialize CSV records list
        g. For each email message:
            i.   Get full message details
            ii.  Extract attachment parts
            iii. Download each attachment to processing folder
            iv.  Append record: {message_id, attachment_count, email_datetime}
            v.   Log: "Downloaded {n} attachments from message {id}"
        h. Write download-{batch}.csv
        i. Log: "Completed employee: {name}, {n} emails, {m} attachments"
    4. Log: "Job-1-Download completed"
```

---

## 9. Detailed Flow: Job-1-Download

```
┌──────────────────────────────────────────────────┐
│           Job-1-Download START                    │
└──────────────────┬───────────────────────────────┘
                   ▼
         ┌─────────────────────┐
         │  init_app()         │
         │  - Setup Logging    │
         │  - Load config.yaml │
         │  - Parse Employees  │
         │  - Auth Gmail       │
         └────────┬────────────┘
                  ▼
      ┌───────────────────────┐
      │ For each Employee     │◄────────────────┐
      └───────────┬───────────┘                 │
                  ▼                             │
      ┌───────────────────────┐                 │
      │ Build Gmail query:    │                 │
      │ from: + dates +       │                 │
      │ has:attachment        │                 │
      └───────────┬───────────┘                 │
                  ▼                             │
      ┌───────────────────────┐                 │
      │ messages.list()       │                 │
      │ (max EMAIL_BATCH_SIZE)│                 │
      └───────────┬───────────┘                 │
                  ▼                             │
           ┌──────────┐  No                    │
           │ Results? ├──────► Log & Skip ─────┤
           └─────┬────┘                         │
                 │ Yes                          │
                 ▼                              │
      ┌───────────────────────┐                 │
      │ Create/verify folders │                 │
      │ Get next batch #      │                 │
      └───────────┬───────────┘                 │
                  ▼                             │
      ┌───────────────────────┐                 │
      │ For each message:     │                 │
      │  - get full message   │                 │
      │  - find attachments   │                 │
      │  - download each      │                 │
      │  - record metadata    │                 │
      └───────────┬───────────┘                 │
                  ▼                             │
      ┌───────────────────────┐                 │
      │ Write download CSV    │                 │
      └───────────┬───────────┘                 │
                  ▼                             │
      ┌───────────────────────┐                 │
      │ Next employee?        ├─── Yes ─────────┘
      └───────────┬───────────┘
                  │ No
                  ▼
         ┌─────────────────────┐
         │ Job-1-Download DONE │
         └─────────────────────┘
```

---

## 10. Job-2-Parse Breakdown

### File: `src/jobs/Job_2_Parse.py`

**Purpose**: Parse downloaded timesheet files (PDF/PNG/XLSX/ZIP) from each employee's `downloaded/` folder, extract structured hours data, and write results to a per-employee Excel workbook with one tab per year.

**Run commands**:
- `python -m src.jobs.Job_2_Parse` — parse files from `downloaded/` folders
- `python -m src.jobs.Job_2_Parse --reparse` — re-process files from `unableToParse/` folders

### New Common Modules Required

| Module | Purpose |
|---|---|
| `src/common/timesheet_parser.py` | PDF text extraction (pdfplumber + PyMuPDF OCR fallback), image OCR (pytesseract + adaptive threshold), Excel parsing, ZIP extraction, regex-based data extraction for 8 formats |
| `src/common/excel_writer.py` | Create/update Excel workbooks, year tabs with weekly/monthly date ranges, update rows by date |

### Pseudocode

```
Job_2_Parse.main():
    1. Parse CLI args (--reparse flag)
    2. Setup logging, load config (does NOT need Gmail auth)
    3. Log: "Job-2-Parse started" (with --reparse mode indicator if applicable)
    4. Set processing_folder = "unableToParse" if --reparse, else config.processing_folder_name
    5. Scan EMPLOYEES_PARENT_DIR_PATH for employee subfolders containing processing_folder
    6. For each employee subfolder:
        a. Log: "Parsing timesheets for: {folder_name}"
        b. Extract any .zip files first:
           - Extract supported files from zip to processing folder
           - Move processed zip to processed/ (or unableToParse/ if empty)
        c. Find all supported files in {folder}/{processing_folder}/
        d. If no files → log & skip
        e. Determine Excel output path: {folder}/{folder_name}.xlsx
        f. For each file:
            i.   Route by type:
                 - .xlsx → parse_xlsx_timesheet() (direct cell extraction)
                 - .pdf → pdfplumber → OCR fallback if empty text
                 - .png/.jpg/.jpeg → size filter (skip <300px) → OCR
            ii.  Parse all entries: dates, hours, project/client, status
                 - Hours sanity check: reject > 168 (weekly) / > 744 (monthly)
            iii. If parsing fails (no entries with hours/dates extracted):
                 - Move file to {folder}/unableToParse/ (create if needed)
                 - Log WARNING: "Unable to parse {filename}, moved to unableToParse/"
                 - Continue to next file
            iv.  Open/create Excel workbook, update matching rows
            v.   Rename file with start and end dates, move to {folder}/processed/
            vi.  Log: "Parsed {filename}: {n} entries, moved to processed/"
        g. Save Excel workbook
        h. Log: "Completed {folder_name}: {parsed} files parsed, {failed} failed"
    7. Log: "Job-2-Parse completed"
```

### Key Design Decisions

| Decision | Choice | Reason |
|---|---|---|
| **PDF library** | `pdfplumber` + `pymupdf` fallback | pdfplumber for text-based PDFs; PyMuPDF renders image-based PDFs to images at 300 DPI for OCR |
| **OCR preprocessing** | 2x upscale + adaptive Gaussian threshold | Significantly better accuracy on colored UI screenshots vs simple Otsu threshold |
| **Image size filter** | Skip images < 300px wide or tall | Filters email signature logos without affecting real timesheets |
| **Excel read** | `openpyxl` | Reads Techno-Comp .xlsx timesheets directly from cell values |
| **Excel write** | `openpyxl` | Only Python lib that supports both read and write of `.xlsx` |
| **ZIP support** | `zipfile` (stdlib) | Extracts .xlsx files from zip archives, skips macOS metadata |
| **Weekly vs Monthly detection** | Date span: ≤7 days = weekly, >7 = monthly | Simple heuristic from date range |
| **Hours sanity check** | Reject > 168 weekly / > 744 monthly | Catches OCR-corrupted values |
| **Unparseable files** | Move to `unableToParse/` | Preserves files, doesn't delete. `--reparse` flag re-processes them after parser improvements |
| **Excel filename** | `{employee_folder_name}.xlsx` | Matches folder name, one file per employee |
| **Tab naming** | Year as string: `"2025"`, `"2026"` | Simple, sortable, one tab per year |
| **No Gmail auth needed** | `init_app()` not used directly | Job-2 only reads local files, uses `init_light()` (logging + config only) |

---

## 11. Detailed Flow: Job-2-Parse

```
┌──────────────────────────────────────────────────┐
│           Job-2-Parse START                       │
└──────────────────┬───────────────────────────────┘
                   ▼
         ┌─────────────────────┐
         │  Setup Logging      │
         │  Load config.yaml   │
         └────────┬────────────┘
                  ▼
      ┌───────────────────────────┐
      │ Scan EMPLOYEES_PARENT_DIR │
      │ for employee subfolders   │
      └───────────┬───────────────┘
                  ▼
      ┌───────────────────────┐
      │ For each subfolder    │◄──────────────────┐
      └───────────┬───────────┘                   │
                  ▼                               │
      ┌───────────────────────┐                   │
      │ List files in         │                   │
      │ {folder}/downloaded/  │                   │
      └───────────┬───────────┘                   │
                  ▼                               │
           ┌──────────┐  No                      │
           │ Files?   ├──────► Log & Skip ───────┤
           └─────┬────┘                           │
                 │ Yes                            │
                 ▼                                │
      ┌───────────────────────┐                   │
      │ For each file:        │◄───────┐          │
      └───────────┬───────────┘        │          │
                  ▼                    │          │
      ┌───────────────────────┐        │          │
      │ Extract text          │        │          │
      │ (PDF→pdfplumber       │        │          │
      │  PNG→pytesseract)     │        │          │
      └───────────┬───────────┘        │          │
                  ▼                    │          │
      ┌───────────────────────┐        │          │
      │ Parse: dates, hours,  │        │          │
      │ project, period type  │        │          │
      └───────────┬───────────┘        │          │
                  ▼                    │          │
           ┌──────────┐  No           │          │
           │Parseable?├───► Move to   │          │
           └─────┬────┘  unableToParse│          │
                 │ Yes      │         │          │
                 ▼          └─────────┤          │
      ┌───────────────────────┐       │          │
      │ Open/create Excel     │       │          │
      │ Get/create year tab   │       │          │
      │ Update matching row   │       │          │
      │ with hours + project  │       │          │
      └───────────┬───────────┘       │          │
                  ▼                   │          │
      ┌───────────────────────┐       │          │
      │ Rename file with      │       │          │
      │ start_end dates       │       │          │
      │ Move to processed/    │       │          │
      └───────────┬───────────┘       │          │
                  ▼                   │          │
      ┌───────────────────────┐       │          │
      │ Next file?            ├─Yes───┘          │
      └───────────┬───────────┘                  │
                  │ No                           │
                  ▼                              │
      ┌───────────────────────┐                  │
      │ Save Excel workbook   │                  │
      └───────────┬───────────┘                  │
                  ▼                              │
      ┌───────────────────────┐                  │
      │ Next subfolder?       ├─── Yes ──────────┘
      └───────────┬───────────┘
                  │ No
                  ▼
         ┌─────────────────────┐
         │ Job-2-Parse DONE    │
         └─────────────────────┘
```

---

## 12. Job-3-MoveToProcessedTag Breakdown

### File: `src/jobs/Job_3_MoveToProcessedTag.py`

**Purpose**: After timesheets have been downloaded (Job-1) and parsed (Job-2), move the corresponding Gmail emails from inbox to a "claudeprocessed" label/tag. This keeps the inbox clean and marks emails as handled without deleting them.

**Run command**: `python -m src.jobs.Job_3_MoveToProcessedTag`

### Prerequisites
- Gmail API scope must be `gmail.modify` (not `gmail.readonly`)
- Download CSVs from Job-1 must exist in employee folders (contain `gmail_message_id` values)
- The "claudeprocessed" Gmail label must exist (Job-3 creates it if missing)

### Pseudocode

```
Job_3_MoveToProcessedTag.main():
    1. Call init_app() → get logger, config, employees, gmail_service
    2. Log: "Job-3-MoveToProcessedTag started"
    3. Ensure "claudeprocessed" label exists in Gmail (create if not)
    4. Scan EMPLOYEES_PARENT_DIR_PATH for employee subfolders
    5. For each employee subfolder:
        a. Log: "Processing: {folder_name}"
        b. Find all download-*.csv files in the folder (not in processed/)
        c. If no CSVs → log & skip
        d. For each CSV:
            i.   Read CSV rows
            ii.  For each gmail_message_id:
                 - Remove INBOX label from the email
                 - Add "claudeprocessed" label to the email
                 - Log: "Moved message {id} to claudeprocessed"
            iii. After all messages in CSV are moved:
                 - Move the CSV file to {folder}/processed/
                 - Log: "Moved {csv_name} to processed/"
        e. Log: "Completed {folder_name}: {n} emails moved"
    6. Log: "Job-3-MoveToProcessedTag completed"
```

### Key Design Decisions

| Decision | Choice | Reason |
|---|---|---|
| **Gmail scope** | `gmail.modify` | Required to add/remove labels; cannot delete emails with this scope |
| **Label name** | `claudeprocessed` | Distinct custom label — won't conflict with existing user labels |
| **Label creation** | Auto-create if missing | First-run convenience; uses `users().labels().create()` |
| **Move = relabel** | Remove INBOX, add custom label | Gmail "move" is label-based; email remains accessible via label |
| **CSV processing** | Move CSV to processed/ after all its emails are moved | Ensures atomicity — partial failures leave CSV in place for retry |
| **Never delete** | Hard rule | Emails are only relabeled, never trashed or permanently deleted |
| **Gmail auth needed** | Uses `init_app()` (full init) | Needs Gmail service for API calls |

### Gmail API Calls Used

```python
# Create label (if not exists)
service.users().labels().create(userId='me', body={
    'name': 'claudeprocessed',
    'labelListVisibility': 'labelShow',
    'messageListVisibility': 'show'
}).execute()

# Move email: remove INBOX label, add claudeprocessed label
service.users().messages().modify(userId='me', id=message_id, body={
    'removeLabelIds': ['INBOX'],
    'addLabelIds': [label_id]
}).execute()
```

---

## 13. Detailed Flow: Job-3-MoveToProcessedTag

```
┌──────────────────────────────────────────────────────┐
│      Job-3-MoveToProcessedTag START                   │
└──────────────────┬───────────────────────────────────┘
                   ▼
         ┌─────────────────────┐
         │  init_app()         │
         │  - Setup Logging    │
         │  - Load config.yaml │
         │  - Parse Employees  │
         │  - Auth Gmail       │
         └────────┬────────────┘
                  ▼
         ┌─────────────────────┐
         │ Ensure              │
         │ "claudeprocessed"   │
         │ label exists        │
         └────────┬────────────┘
                  ▼
      ┌───────────────────────────┐
      │ Scan EMPLOYEES_PARENT_DIR │
      │ for employee subfolders   │
      └───────────┬───────────────┘
                  ▼
      ┌───────────────────────┐
      │ For each subfolder    │◄──────────────────┐
      └───────────┬───────────┘                   │
                  ▼                               │
      ┌───────────────────────┐                   │
      │ Find download-*.csv   │                   │
      │ (not in processed/)   │                   │
      └───────────┬───────────┘                   │
                  ▼                               │
           ┌──────────┐  No                      │
           │ CSVs?    ├──────► Log & Skip ───────┤
           └─────┬────┘                           │
                 │ Yes                            │
                 ▼                                │
      ┌───────────────────────┐                   │
      │ For each CSV:         │◄───────┐          │
      └───────────┬───────────┘        │          │
                  ▼                    │          │
      ┌───────────────────────┐        │          │
      │ For each message_id:  │        │          │
      │  - Remove INBOX label │        │          │
      │  - Add claudeprocessed│        │          │
      └───────────┬───────────┘        │          │
                  ▼                    │          │
      ┌───────────────────────┐        │          │
      │ Move CSV to processed/│        │          │
      └───────────┬───────────┘        │          │
                  ▼                    │          │
      ┌───────────────────────┐        │          │
      │ Next CSV?             ├─Yes────┘          │
      └───────────┬───────────┘                   │
                  │ No                            │
                  ▼                               │
      ┌───────────────────────┐                   │
      │ Next subfolder?       ├─── Yes ───────────┘
      └───────────┬───────────┘
                  │ No
                  ▼
         ┌───────────────────────────┐
         │ Job-3-MoveToProcessedTag  │
         │ DONE                      │
         └───────────────────────────┘
```

---

## 14. Error Handling Strategy

| Scenario | Job | Handling |
|---|---|---|
| `config.yaml` missing/invalid | All | Log CRITICAL, exit with clear error message |
| `EmployeesList.md` malformed | Job-1 | Log ERROR, skip malformed rows, continue with valid ones |
| OAuth token expired | Job-1 | Auto-refresh; if refresh fails, re-run consent flow |
| Gmail API rate limit (429) | Job-1 | Exponential backoff: 1s, 2s, 4s, 8s, 16s (max 5 retries) |
| Gmail API server error (500/503) | Job-1 | Same exponential backoff as rate limit |
| No emails found for employee | Job-1 | Log INFO, skip to next employee (not an error) |
| Attachment download fails | Job-1 | Log ERROR with message ID & attachment name, skip attachment, continue |
| Folder creation fails (permissions) | All | Log CRITICAL, exit |
| Individual employee processing fails | All | Log ERROR, continue to next employee |
| Network timeout | Job-1 | Retry with backoff, log WARNING |
| PDF text extraction fails | Job-2 | Move file to `unableToParse/`, log WARNING, continue |
| OCR returns empty/garbled text | Job-2 | Move file to `unableToParse/`, log WARNING, continue |
| No hours or dates extracted | Job-2 | Move file to `unableToParse/`, log WARNING, continue |
| No matching week/month row in Excel | Job-2 | Log WARNING with date details, skip (don't crash) |
| Excel file corrupted/locked | Job-2 | Log ERROR, skip employee, continue |
| Unsupported file type in downloaded/ | Job-2 | Log WARNING, skip file, continue |
| Tesseract not installed | Job-2 | Log CRITICAL with install instructions, exit |
| Image too small (< 300px) | Job-2 | Return empty text (email signature), move to unableToParse/ |
| OCR-corrupted hours (> 168/744) | Job-2 | Reject value, try next hours pattern in chain |
| ZIP contains no supported files | Job-2 | Move zip to unableToParse/, log WARNING |
| Corrupted file inside ZIP | Job-2 | Log ERROR, move extracted file to unableToParse/, continue |
| Gmail label creation fails | Job-3 | Log CRITICAL, exit |
| Email relabel fails (single msg) | Job-3 | Log ERROR, continue to next message in CSV |
| CSV partially processed | Job-3 | Leave CSV in place (not moved to processed/), can retry |
| Gmail modify scope missing | Job-3 | API returns 403 — log CRITICAL with re-auth instructions |

**Principle**: Never crash the entire run because of one employee's failure. Log and move on.

---

## 15. Logging Strategy

### Log Levels Used

| Level | When |
|---|---|
| DEBUG | Gmail API queries, raw responses, file paths |
| INFO | Starting/completing employee, download counts, batch CSV written |
| WARNING | Rate limits hit, retries, no emails found |
| ERROR | Failed downloads, malformed employee rows, API errors |
| CRITICAL | Config missing, auth failure, permission errors |

### Log File Rotation

- **Location**: `logs/timesheets_processor.log`
- **Strategy**: `RotatingFileHandler`
- **Max size**: 5 MB per file
- **Backup count**: 7 files (total ~40 MB max)
- **Format**: `2026-02-26 10:30:00 | INFO     | Job_1_Download | Processing employee: John Smith`

### Key Events to Log

```
--- Job-1-Download ---
[INFO]  Job-1-Download started with config: START_DATE=..., END_DATE=..., BATCH_SIZE=...
[INFO]  Loaded {n} employees from {file}
[INFO]  Gmail API authenticated successfully
[INFO]  Processing employee: {name} ({email}) for project: {project}
[INFO]  Gmail query: {query}
[INFO]  Found {n} emails for {name}
[INFO]  Saved attachment: {filename} ({size} bytes) to {path}
[INFO]  Wrote download-{batch}.csv with {n} records for {name}
[INFO]  Completed employee: {name} - {emails} emails, {attachments} attachments
[INFO]  Job-1-Download completed.
[WARN]  No emails found for {name} ({email})
[ERROR] Failed to download attachment {filename} from message {id}: {error}

--- Job-2-Parse ---
[INFO]  Job-2-Parse started
[INFO]  Found {n} employee folders in {path}
[INFO]  Parsing timesheets for: {folder_name}
[INFO]  Found {n} files in {folder}/downloaded/
[INFO]  Extracted text from {filename} ({n} characters)
[INFO]  Parsed {filename}: {hours}h, {start}-{end}, period={weekly|monthly}, project={project}
[INFO]  Updated Excel {year} tab: week {start}-{end} → {hours}h
[INFO]  Saved {folder_name}.xlsx
[INFO]  Completed {folder_name}: {parsed} parsed, {failed} failed
[INFO]  Job-2-Parse completed.
[WARN]  Unable to parse {filename}: no hours found — moved to unableToParse/
[WARN]  Unable to parse {filename}: no dates found — moved to unableToParse/
[WARN]  No matching row for date {date} in {year} tab
[WARN]  Skipping small image {filename} ({w}x{h}) — likely email signature
[WARN]  OCR hours value {val} exceeds sanity limit — rejected
[ERROR] Failed to process {folder_name}: {error}

--- Job-3-MoveToProcessedTag ---
[INFO]  Job-3-MoveToProcessedTag started
[INFO]  Label "claudeprocessed" ready (id: {label_id})
[INFO]  Processing: {folder_name}
[INFO]  Found {n} download CSVs in {folder_name}
[INFO]  Moved message {id} to claudeprocessed
[INFO]  Moved {csv_name} to processed/
[INFO]  Completed {folder_name}: {n} emails moved
[INFO]  Job-3-MoveToProcessedTag completed.
[WARN]  No download CSVs found for {folder_name}
[ERROR] Failed to move message {id}: {error}
```

---

## 16. Implementation Phases

### Phase 1: Restructure to Modular Architecture (DONE)
| Step | Action |
|---|---|
| 1.1 | Create `src/common/` directory with `__init__.py` |
| 1.2 | Move existing modules from `src/` → `src/common/` |
| 1.3 | Create `src/common/__init__.py` with `init_app()` convenience function |
| 1.4 | Create `src/jobs/` directory with `__init__.py` |

### Phase 2: Job-1-Download (DONE)
| Step | Action |
|---|---|
| 2.1 | Create `src/jobs/Job_1_Download.py` — email attachment downloader |
| 2.2 | Verified working with live Gmail API |

### Phase 3: Job-2-Parse — Common Modules (DONE)
| Step | File | Action |
|---|---|---|
| 3.1 | `requirements.txt` | Add pdfplumber, pytesseract, Pillow, opencv-python, openpyxl, pymupdf |
| 3.2 | `src/common/timesheet_parser.py` | PDF text extraction (pdfplumber + PyMuPDF OCR fallback), image OCR (pytesseract + adaptive threshold), Excel parsing, ZIP extraction, 8 timesheet format parsers, OCR error tolerance |
| 3.3 | `src/common/excel_writer.py` | Workbook create/load, year tab creation with weekly/monthly date ranges, row update by date match |
| 3.4 | `src/common/__init__.py` | Add `init_app_light()` — logging + config only (no Gmail auth) for Job-2 |

### Phase 4: Job-2-Parse — Executable (DONE)
| Step | File | Action |
|---|---|---|
| 4.1 | `src/jobs/Job_2_Parse.py` | Main orchestrator: scan folders, parse files (PDF/PNG/XLSX/ZIP), update Excel, handle failures |
| 4.2 | — | `--reparse` CLI flag to re-process `unableToParse/` files |
| 4.3 | — | ZIP extraction before individual file processing |
| 4.4 | — | Verified: 361 parsed, 82 failed across 18 employees. 17/18 employees have Excel output (1 employee has only email signatures). |

### Phase 5: Job-3-MoveToProcessedTag
| Step | File | Action |
|---|---|---|
| 5.1 | `src/common/gmail_auth.py` | Update SCOPES to `gmail.modify` (delete `token.pickle` to re-auth) |
| 5.2 | `src/common/gmail_client.py` | Add `ensure_label()`, `move_to_label()` functions |
| 5.3 | `src/jobs/Job_3_MoveToProcessedTag.py` | Main orchestrator: scan for download CSVs, move emails to "claudeprocessed" label, move processed CSVs |
| 5.4 | — | Test with live Gmail API |

---

## 17. Critical Gotchas & Constraints

### Gmail API Specifics
1. **Date format in queries**: Must use `YYYY/MM/DD` (slashes), not dashes — dashes silently fail
2. **`before:` is exclusive**: To include END_DATE, query `before:{END_DATE + 1 day}`
3. **`messages.list()` returns stubs only**: Just `{id, threadId}` — must call `messages.get()` for content
4. **Attachment encoding**: Gmail uses **URL-safe base64** — must use `base64.urlsafe_b64decode()`
5. **MIME parts are nested**: Must recursively walk `multipart/*` containers to find all attachments
6. **`internalDate` is milliseconds**: Divide by 1000 for Python `datetime.fromtimestamp()`
7. **Scope is gmail.modify**: Required by Job-3 to add/remove labels. Cannot permanently delete emails. If upgrading from `gmail.readonly`, delete `token.pickle` and re-authenticate

### File System
8. **Folder name sanitization**: Strip spaces, dots, special chars — only keep `[a-zA-Z0-9_]`
9. **Attachment filename conflicts**: Prefix with `{message_id}_` to avoid collisions
10. **Path safety**: Use `Path(filename).name` to strip any path traversal attempts in filenames

### Batch Processing
11. **EMAIL_BATCH_SIZE** limits `maxResults` in `messages.list()` — controls emails per run, not a pagination page size
12. **Batch CSV numbering**: Scan existing `download-*.csv` files to determine next batch number
13. **Idempotency**: The message ID in the CSV can be used to avoid re-downloading in future enhancements

### Config Parsing
14. **Date input format**: `MM-DD-YYYY` in config → parse with `strptime('%m-%d-%Y')`
15. **Date query format**: Convert to `YYYY/MM/DD` for Gmail API query
16. **Empty string = use default**: All optional params default when value is `""` or not set

### Modular Architecture
17. **Job-1 uses `init_app()`** (logging + config + employees + Gmail auth)
18. **Job-2 uses `init_app_light()`** (logging + config only — no Gmail auth needed)
19. **Job naming convention**: `Job_{N}_{Name}.py` — underscore for Python module compatibility
20. **Each job is independently executable**: `python -m src.jobs.Job_1_Download` / `python -m src.jobs.Job_2_Parse`
21. **Common modules must remain stateless**: No global mutable state — pass config/service as arguments

### Timesheet Parsing (Job-2)
22. **Tesseract is a system dependency**: Must be installed via `brew install tesseract` before Job-2 can run
23. **File type routing**: `.pdf` → pdfplumber (+ OCR fallback), `.png/.jpg/.jpeg` → pytesseract, `.xlsx` → direct cell parsing, `.zip` → extract then parse contents
24. **PDF OCR fallback**: When pdfplumber returns empty text (image-based PDFs), PyMuPDF renders pages at 300 DPI and runs OCR
25. **OCR preprocessing**: 2x upscale + adaptive Gaussian threshold — much better than simple Otsu for colored UI screenshots
26. **Image size filter**: Images < 300px wide or tall are skipped — these are email signature logos, not timesheets
27. **OCR error tolerance**: Parser handles common misreads: `F` dropped from "From", `'` instead of `/`, `§` junk chars, `S` instead of `8`, `<` instead of `4`
28. **Hours sanity check**: Reject OCR-corrupted hours values > 168 (weekly) or > 744 (monthly)
29. **Hours search chain**: `Total Hours` → `Reported Hours` → `Totals:` → inline hours after dates → standalone `Total XX.XX`
30. **8 timesheet formats**: Oracle HCM, PeopleSoft HRCMS, MBO Partners, Unanet Time-List, Unanet EaZyTyme, calendar grid, Techno-Comp Excel, ZIP archives
31. **Weekly vs monthly detection**: If date span ≤ 7 days → weekly (Sun-Sat); if > 7 days → monthly
32. **Week alignment**: Weekly rows use Sunday as start day — extracted dates must be normalized to the containing Sunday
33. **openpyxl returns datetime, not date**: Always normalize `cell.value` with `.date()` before comparing
34. **Move, don't delete**: Unparseable files are moved to `unableToParse/`, never deleted. Use `--reparse` to re-process after parser improvements
35. **Excel filename**: `{employee_folder_name}.xlsx` — matches the folder name for consistency
36. **Year tabs created lazily**: Only created when a timesheet with that year's date is encountered
37. **Regex parsing is best-effort**: Timesheet formats vary wildly — parse what we can, move the rest to `unableToParse/`
38. **Project/client extraction**: PeopleSoft timesheets contain a category line (e.g., "VLO Category 2B", "Full Service Cat1") used as project identifier. Oracle HCM timesheets have no project field — left empty
39. **Multi-entry files**: A single file (especially Oracle HCM) may contain 20+ weekly timesheet entries — all parsed and written to Excel in one pass
40. **Post-parse file management**: Successfully parsed files are renamed with date range prefix and moved to `processed/` folder
41. **ZIP handling**: Extract supported files first, then process extracted files individually. Skip macOS metadata (`__MACOSX/`, `._` files, `~$` temp files)

### Email Processing (Job-3)
42. **Gmail "move" is relabeling**: Remove INBOX label + add "claudeprocessed" label — email is never deleted
43. **Label auto-creation**: Job-3 creates "claudeprocessed" label on first run if it doesn't exist
44. **CSV-atomic processing**: CSV is only moved to `processed/` after ALL its emails are successfully relabeled — partial failures leave CSV in place for retry
45. **Never delete emails**: Hard rule — Job-3 only adds/removes labels, never calls `messages().trash()` or `messages().delete()`

---

## Summary of Files

| # | File | Status | Purpose |
|---|---|---|---|
| 1 | `requirements.txt` | Done | Dependencies (includes pymupdf for OCR fallback) |
| 2 | `.gitignore` | Done | Ignore secrets, caches, logs |
| 3 | `config/config.yaml` | Done | Configuration template |
| 4 | `config/EmployeesList.md` | Done | Employee list |
| 5 | `src/__init__.py` | Done | Root package marker |
| 6 | `src/common/__init__.py` | Done | Module #0 — `init_app()` and `init_app_light()` |
| 7 | `src/common/logger_setup.py` | Done | Logging with rotation |
| 8 | `src/common/config_loader.py` | Done | Config reader + validation |
| 9 | `src/common/employee_parser.py` | Done | Employee list parser |
| 10 | `src/common/gmail_auth.py` | Update | OAuth2 authentication — needs `gmail.modify` scope for Job-3 |
| 11 | `src/common/gmail_client.py` | Update | Gmail search + download — add label management for Job-3 |
| 12 | `src/common/file_manager.py` | Done | Folders, sanitization, CSV |
| 13 | `src/common/timesheet_parser.py` | Done | PDF/image/Excel text extraction, OCR fallback, 8 format parsers, ZIP extraction |
| 14 | `src/common/excel_writer.py` | Done | Excel workbook, year tabs, date ranges, row updates |
| 15 | `src/jobs/__init__.py` | Done | Jobs package marker |
| 16 | `src/jobs/Job_1_Download.py` | Done | **Job-1-Download** executable |
| 17 | `src/jobs/Job_2_Parse.py` | Done | **Job-2-Parse** executable (with `--reparse` flag) |
| 18 | `src/jobs/Job_3_MoveToProcessedTag.py` | **New** | **Job-3-MoveToProcessedTag** executable |
