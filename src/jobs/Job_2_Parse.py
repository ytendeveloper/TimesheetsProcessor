"""
Job-2-Parse — Timesheet File Parser

Scans employee folders for downloaded timesheet files (PDF/PNG),
extracts structured data (dates, hours, project), and writes results
to per-employee Excel workbooks with one tab per year.

Successfully parsed files are renamed with start and end dates and
moved to a 'processed' subfolder. Files that cannot be parsed are
moved to an 'unableToParse' subfolder. Files are never deleted.

This job uses Module #0 (src.common) for initialization and the
timesheet_parser and excel_writer common modules for parsing and output.

Usage:
    python -m src.jobs.Job_2_Parse
    python -m src.jobs.Job_2_Parse --reparse   # re-process unableToParse/ files
"""

import argparse
import logging
import shutil
from pathlib import Path

from src.common import init_app_light
from src.common.timesheet_parser import (
    extract_text, parse_all_entries, parse_xlsx_timesheet,
    extract_from_zip, normalize_to_week_start,
    SUPPORTED_EXTENSIONS,
)
from src.common.excel_writer import (
    get_or_create_workbook, get_or_create_year_sheet,
    update_weekly_row, update_monthly_row, save_workbook,
)

logger = logging.getLogger("timesheets_processor")

JOB_NAME = "Job-2-Parse"


def main():
    """Entry point for Job-2-Parse."""
    parser = argparse.ArgumentParser(description="Job-2-Parse: Timesheet File Parser")
    parser.add_argument(
        "--reparse", action="store_true",
        help="Re-process files from unableToParse/ folders instead of downloaded/",
    )
    args = parser.parse_args()

    # --- Initialize (logging + config only, no Gmail auth needed) ---
    app_logger, config = init_app_light()

    app_logger.info("=" * 60)
    app_logger.info("%s started%s", JOB_NAME, " (--reparse mode)" if args.reparse else "")

    parent_dir = config.employees_parent_dir
    processing_folder = "unableToParse" if args.reparse else config.processing_folder_name

    # Scan for employee subfolders that contain the target folder
    if not parent_dir.exists():
        app_logger.critical("EMPLOYEES_PARENT_DIR_PATH does not exist: %s", parent_dir)
        return

    employee_folders = sorted([
        d for d in parent_dir.iterdir()
        if d.is_dir() and (d / processing_folder).is_dir()
    ])

    if not employee_folders:
        app_logger.warning(
            "No employee folders with '%s' subfolder found in %s",
            processing_folder, parent_dir
        )
        return

    app_logger.info(
        "Found %d employee folder(s) in %s", len(employee_folders), parent_dir
    )

    # --- Process each employee folder ---
    for folder in employee_folders:
        app_logger.info("Parsing timesheets for: %s", folder.name)

        try:
            _process_employee_folder(folder, processing_folder)
        except Exception as e:
            app_logger.error(
                "Failed to process %s: %s", folder.name, e, exc_info=True
            )
            continue

    app_logger.info("=" * 60)
    app_logger.info("%s completed", JOB_NAME)


def _process_employee_folder(employee_dir: Path, processing_folder: str):
    """
    Parse all timesheet files in an employee's downloaded folder.

    For each file:
    - Extract text (PDF or image OCR)
    - Parse structured data (dates, hours, project)
    - If unparseable → move to unableToParse/
    - If parseable → update the employee's Excel workbook

    Args:
        employee_dir: Path to the employee's folder (e.g. Employees/AjayBenadict_ZZ/).
        processing_folder: Name of the downloads subfolder (e.g. "downloaded").
    """
    download_dir = employee_dir / processing_folder

    # Extract any zip files first, then remove the zip from the processing list
    for zf in sorted(download_dir.glob("*.zip")):
        try:
            extracted = extract_from_zip(zf, download_dir)
            if extracted:
                _move_to_processed_simple(zf, employee_dir)
            else:
                _move_to_unable_to_parse(zf, employee_dir)
        except Exception as e:
            logger.error("Error extracting %s: %s", zf.name, e, exc_info=True)
            _move_to_unable_to_parse(zf, employee_dir)

    # Find all supported files in the downloaded folder (excluding zip, already handled)
    files = sorted([
        f for f in download_dir.iterdir()
        if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS
           and f.suffix.lower() != ".zip"
    ])

    if not files:
        logger.info("No supported files in %s. Skipping.", download_dir)
        return

    logger.info("Found %d file(s) in %s", len(files), download_dir)

    # Excel output path: {employee_dir}/{folder_name}.xlsx
    excel_path = employee_dir / f"{employee_dir.name}.xlsx"
    wb = get_or_create_workbook(excel_path)

    parsed_count = 0
    failed_count = 0

    for file_path in files:
        try:
            success = _parse_and_update(file_path, employee_dir, wb)
            if success:
                parsed_count += 1
            else:
                failed_count += 1
        except Exception as e:
            logger.error(
                "Error processing %s: %s", file_path.name, e, exc_info=True
            )
            _move_to_unable_to_parse(file_path, employee_dir)
            failed_count += 1

    # Save the workbook if any files were parsed
    if parsed_count > 0:
        save_workbook(wb, excel_path)

    logger.info(
        "Completed %s: %d parsed, %d failed",
        employee_dir.name, parsed_count, failed_count
    )


def _parse_and_update(file_path: Path, employee_dir: Path, wb) -> bool:
    """
    Extract text from a file, parse all timesheet entries, and update the Excel workbook.

    Handles multi-entry timesheets where a single file contains multiple
    weekly time card blocks.

    Args:
        file_path: Path to the timesheet file.
        employee_dir: Path to the employee's folder.
        wb: The openpyxl Workbook to update.

    Returns:
        True if successfully parsed and written, False if moved to unableToParse.
    """
    # Step 1: Extract and parse
    if file_path.suffix.lower() == ".xlsx":
        # xlsx timesheets are parsed directly from cell values
        file_data = parse_xlsx_timesheet(file_path)
    else:
        text = extract_text(file_path)
        if not text or not text.strip():
            logger.warning(
                "Unable to parse %s: no text extracted — moved to unableToParse/",
                file_path.name
            )
            _move_to_unable_to_parse(file_path, employee_dir)
            return False

        logger.debug("Extracted text from %s: %d characters", file_path.name, len(text))

        # Step 2: Parse all entries from the file
        file_data = parse_all_entries(text)

    if not file_data.entries:
        logger.warning(
            "Unable to parse %s: no valid entries found — moved to unableToParse/",
            file_path.name
        )
        _move_to_unable_to_parse(file_path, employee_dir)
        return False

    # Step 3: Update Excel for each entry
    updated_count = 0
    project = file_data.project
    for entry in file_data.entries:
        year = entry.start_date.year
        ws = get_or_create_year_sheet(wb, year, mode=entry.period_type)

        if entry.period_type == "weekly":
            week_start = normalize_to_week_start(entry.start_date)
            success = update_weekly_row(ws, week_start, entry.hours, project)
            if not success:
                logger.warning(
                    "No matching weekly row for date %s in %d tab",
                    week_start, year
                )
            else:
                updated_count += 1
        else:
            success = update_monthly_row(
                ws, year, entry.start_date.month, entry.hours, project
            )
            if not success:
                logger.warning(
                    "No matching monthly row for %d/%d in %d tab",
                    entry.start_date.month, year, year
                )
            else:
                updated_count += 1

    # Step 4: Rename with date range and move to processed/
    earliest_start = min(e.start_date for e in file_data.entries)
    latest_end = max(e.end_date for e in file_data.entries)
    _move_to_processed(file_path, employee_dir, earliest_start, latest_end)

    logger.info(
        "Parsed %s: %d entries found, %d rows updated (employee: %s) — moved to processed/",
        file_path.name, len(file_data.entries), updated_count,
        file_data.employee_name
    )
    return True


def _move_to_processed(file_path: Path, employee_dir: Path,
                       start_date: 'date', end_date: 'date'):
    """
    Rename a file with start and end dates and move to the processed/ folder.

    New filename format: {start_date}_{end_date}_{original_name}
    e.g., 2025-03-02_2025-08-30_AjayTimesheetsFromMarch.pdf

    Args:
        file_path: Path to the successfully parsed file.
        employee_dir: Path to the employee's folder.
        start_date: Earliest start date from parsed entries.
        end_date: Latest end date from parsed entries.
    """
    processed_dir = employee_dir / "processed"
    processed_dir.mkdir(exist_ok=True)

    new_name = f"{start_date}_{end_date}_{file_path.name}"
    dest = processed_dir / new_name

    # Avoid overwriting if a file with the same name already exists
    if dest.exists():
        stem = f"{start_date}_{end_date}_{file_path.stem}"
        suffix = file_path.suffix
        counter = 1
        while dest.exists():
            dest = processed_dir / f"{stem}_{counter}{suffix}"
            counter += 1

    shutil.move(str(file_path), str(dest))
    logger.debug("Moved %s to %s", file_path.name, dest)


def _move_to_processed_simple(file_path: Path, employee_dir: Path):
    """Move a file to processed/ without date-renaming (used for zip archives)."""
    processed_dir = employee_dir / "processed"
    processed_dir.mkdir(exist_ok=True)
    dest = processed_dir / file_path.name
    if dest.exists():
        counter = 1
        while dest.exists():
            dest = processed_dir / f"{file_path.stem}_{counter}{file_path.suffix}"
            counter += 1
    shutil.move(str(file_path), str(dest))
    logger.debug("Moved %s to %s", file_path.name, dest)


def _move_to_unable_to_parse(file_path: Path, employee_dir: Path):
    """
    Move a file to the unableToParse subfolder within the employee directory.

    Creates the folder if it does not exist. Never deletes files.

    Args:
        file_path: Path to the file to move.
        employee_dir: Path to the employee's folder.
    """
    unable_dir = employee_dir / "unableToParse"
    unable_dir.mkdir(exist_ok=True)

    dest = unable_dir / file_path.name
    # Avoid overwriting if a file with the same name already exists
    if dest.exists():
        stem = file_path.stem
        suffix = file_path.suffix
        counter = 1
        while dest.exists():
            dest = unable_dir / f"{stem}_{counter}{suffix}"
            counter += 1

    shutil.move(str(file_path), str(dest))
    logger.debug("Moved %s to %s", file_path.name, dest)


if __name__ == "__main__":
    main()
