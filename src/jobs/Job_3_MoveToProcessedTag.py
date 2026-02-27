"""
Job-3-MoveToProcessedTag — Email Label Mover

After Job-1 downloads attachments and Job-2 parses them, this job closes
the loop: it reads the download-*.csv files to find Gmail message IDs,
moves those emails from INBOX to a "claudeprocessed" label, then moves
the CSV files into processed/ subfolders.

This job uses Module #0 (src.common) for initialization, configuration,
and Gmail API access.

Usage:
    python -m src.jobs.Job_3_MoveToProcessedTag
"""

import csv
import logging
import shutil
from pathlib import Path

from src.common import init_app
from src.common.gmail_client import ensure_label, move_to_label

logger = logging.getLogger("timesheets_processor")

JOB_NAME = "Job-3-MoveToProcessedTag"
LABEL_NAME = "claudeprocessed"


def main():
    """Entry point for Job-3-MoveToProcessedTag."""
    # --- Initialize via Module #0 ---
    app_logger, config, employees, service = init_app()

    app_logger.info("=" * 60)
    app_logger.info("%s started", JOB_NAME)

    # --- Ensure the "claudeprocessed" label exists ---
    label_id = ensure_label(service, LABEL_NAME, user_id=config.gmail_user_id)

    # --- Scan employee subfolders ---
    parent_dir = config.employees_parent_dir
    if not parent_dir.exists():
        app_logger.critical("EMPLOYEES_PARENT_DIR_PATH does not exist: %s", parent_dir)
        return

    employee_folders = sorted([
        d for d in parent_dir.iterdir()
        if d.is_dir() and not d.name.startswith(".")
    ])

    if not employee_folders:
        app_logger.warning("No employee folders found in %s", parent_dir)
        return

    app_logger.info("Found %d employee folder(s) in %s", len(employee_folders), parent_dir)

    total_messages_moved = 0
    total_csvs_moved = 0

    for folder in employee_folders:
        try:
            moved_msgs, moved_csvs = _process_employee_folder(
                service, folder, label_id, config.gmail_user_id
            )
            total_messages_moved += moved_msgs
            total_csvs_moved += moved_csvs
        except Exception as e:
            app_logger.error(
                "Failed to process %s: %s", folder.name, e, exc_info=True
            )
            continue

    app_logger.info("=" * 60)
    app_logger.info(
        "%s completed: %d emails moved to '%s', %d CSVs moved to processed/",
        JOB_NAME, total_messages_moved, LABEL_NAME, total_csvs_moved
    )


def _process_employee_folder(service, employee_dir: Path, label_id: str,
                              user_id: str) -> tuple[int, int]:
    """
    Process all download-*.csv files in an employee folder root.

    For each CSV:
    - Read gmail_message_id values
    - Move each email to the claudeprocessed label
    - If ALL messages succeed, move the CSV to processed/
    - If any fail, leave the CSV in place

    Args:
        service: Authenticated Gmail API service.
        employee_dir: Path to the employee's folder.
        label_id: Gmail label ID for "claudeprocessed".
        user_id: Gmail user ID.

    Returns:
        Tuple of (total_messages_moved, total_csvs_moved).
    """
    # Find download-*.csv files in the folder root (NOT in processed/)
    csv_files = sorted(employee_dir.glob("download-*.csv"))

    if not csv_files:
        return 0, 0

    logger.info("Found %d download CSV(s) in %s", len(csv_files), employee_dir.name)

    total_messages_moved = 0
    total_csvs_moved = 0

    for csv_path in csv_files:
        success_count, fail_count = _process_csv(
            service, csv_path, label_id, user_id
        )

        total_messages_moved += success_count

        if fail_count == 0 and success_count > 0:
            # All messages moved successfully — move CSV to processed/
            _move_csv_to_processed(csv_path, employee_dir)
            total_csvs_moved += 1
        elif fail_count > 0:
            logger.warning(
                "Leaving %s in place: %d/%d messages failed to move",
                csv_path.name, fail_count, success_count + fail_count
            )

    return total_messages_moved, total_csvs_moved


def _process_csv(service, csv_path: Path, label_id: str,
                 user_id: str) -> tuple[int, int]:
    """
    Read a download CSV and move each email to the claudeprocessed label.

    Args:
        service: Authenticated Gmail API service.
        csv_path: Path to the download-*.csv file.
        label_id: Gmail label ID.
        user_id: Gmail user ID.

    Returns:
        Tuple of (success_count, fail_count).
    """
    success_count = 0
    fail_count = 0

    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            message_id = row.get("gmail_message_id", "").strip()
            if not message_id:
                continue

            try:
                move_to_label(service, message_id, label_id, user_id=user_id)
                success_count += 1
            except Exception as e:
                logger.error(
                    "Failed to move message %s to label: %s",
                    message_id, e
                )
                fail_count += 1

    logger.info(
        "CSV %s: %d messages moved, %d failed",
        csv_path.name, success_count, fail_count
    )
    return success_count, fail_count


def _move_csv_to_processed(csv_path: Path, employee_dir: Path):
    """Move a download CSV to the processed/ subfolder."""
    processed_dir = employee_dir / "processed"
    processed_dir.mkdir(exist_ok=True)

    dest = processed_dir / csv_path.name
    if dest.exists():
        counter = 1
        while dest.exists():
            dest = processed_dir / f"{csv_path.stem}_{counter}{csv_path.suffix}"
            counter += 1

    shutil.move(str(csv_path), str(dest))
    logger.info("Moved %s to %s", csv_path.name, dest)


if __name__ == "__main__":
    main()
