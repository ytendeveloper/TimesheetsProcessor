"""
File and folder management for Timesheets Processor.

Handles:
- Creating employee folder structures
- Determining batch run numbers from existing CSV files
- Writing download metadata CSV files
"""

import csv
import logging
import re
from datetime import datetime
from pathlib import Path

logger = logging.getLogger("timesheets_processor")


def ensure_employee_folders(parent_dir: Path, folder_name: str,
                            processing_folder: str) -> tuple[Path, Path]:
    """
    Create the employee folder and its processing subfolder if they don't exist.

    Directory structure:
        {parent_dir}/{folder_name}/
        {parent_dir}/{folder_name}/{processing_folder}/

    Args:
        parent_dir: EMPLOYEES_PARENT_DIR_PATH — employee folders are created directly here.
        folder_name: Sanitized employee folder name (e.g. "JohnDoe_ProjectAlpha").
        processing_folder: Name of the downloads subfolder (e.g. "downloaded").

    Returns:
        Tuple of (employee_folder_path, processing_folder_path).
    """
    employee_dir = parent_dir / folder_name
    processing_dir = employee_dir / processing_folder

    employee_dir.mkdir(parents=True, exist_ok=True)
    processing_dir.mkdir(parents=True, exist_ok=True)

    logger.debug("Ensured folders: %s and %s", employee_dir, processing_dir)
    return employee_dir, processing_dir


def get_next_batch_number(employee_dir: Path) -> int:
    """
    Determine the next batch run number by scanning existing download CSV files.

    Looks for files matching 'download-{N}.csv' and returns the next number.

    Args:
        employee_dir: Path to the employee's folder.

    Returns:
        Next batch number (1 if no existing CSVs found).
    """
    existing = list(employee_dir.glob("download-*.csv"))
    if not existing:
        return 1

    # Extract numbers from filenames like download-1.csv, download-12.csv
    numbers = []
    for f in existing:
        match = re.search(r"download-(\d+)\.csv$", f.name)
        if match:
            numbers.append(int(match.group(1)))

    return max(numbers) + 1 if numbers else 1


def write_download_csv(employee_dir: Path, batch_number: int,
                       records: list[dict]) -> Path:
    """
    Write download metadata to a CSV file for this batch run.

    Creates a file named 'download-{batch_number}.csv' with columns:
    gmail_message_id, num_attachments, email_datetime

    Args:
        employee_dir: Path to the employee's folder.
        batch_number: The batch run number for this file.
        records: List of dicts, each with keys:
            - gmail_message_id (str)
            - num_attachments (int)
            - email_datetime (datetime)

    Returns:
        Path to the written CSV file.
    """
    csv_path = employee_dir / f"download-{batch_number}.csv"
    fieldnames = ["gmail_message_id", "num_attachments", "email_datetime"]

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for record in records:
            writer.writerow({
                "gmail_message_id": record["gmail_message_id"],
                "num_attachments": record["num_attachments"],
                "email_datetime": record["email_datetime"].strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
            })

    logger.info(
        "Wrote %s with %d records", csv_path.name, len(records)
    )
    return csv_path
