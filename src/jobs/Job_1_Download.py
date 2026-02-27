"""
Job-1-Download — Email Attachment Downloader

Searches Gmail for emails with attachments from each employee in the
employees list, downloads attachments into organized per-employee folders,
and writes a metadata CSV for each batch run.

This job uses Module #0 (src.common) for initialization, configuration,
Gmail API access, and file management.

Usage:
    python -m src.jobs.Job_1_Download
"""

import logging

from src.common import init_app
from src.common.gmail_client import search_emails, get_message_details, download_attachment
from src.common.file_manager import ensure_employee_folders, get_next_batch_number, write_download_csv

logger = logging.getLogger("timesheets_processor")

JOB_NAME = "Job-1-Download"


def main():
    """Entry point for Job-1-Download."""
    # --- Initialize via Module #0 ---
    app_logger, config, employees, service = init_app()

    app_logger.info("=" * 60)
    app_logger.info("%s started", JOB_NAME)

    # --- Process each employee ---
    total_employees = len(employees)

    for idx, employee in enumerate(employees, start=1):
        app_logger.info(
            "[%d/%d] Processing employee: %s (%s) — Project: %s",
            idx, total_employees, employee.name, employee.email, employee.project
        )

        try:
            _process_employee(service, config, employee)
        except Exception as e:
            # Log error but continue to next employee — don't crash the whole run
            app_logger.error(
                "Failed to process employee %s (%s): %s",
                employee.name, employee.email, e,
                exc_info=True
            )
            continue

    app_logger.info("=" * 60)
    app_logger.info("%s completed", JOB_NAME)


def _process_employee(service, config, employee):
    """
    Process a single employee: search emails, download attachments, write CSV.

    Args:
        service: Authenticated Gmail API service.
        config: Validated Config object.
        employee: Employee dataclass instance.
    """
    # Search Gmail for emails with attachments from this employee
    message_stubs = search_emails(
        service,
        sender_email=employee.email,
        start_date=config.start_date,
        end_date=config.end_date,
        batch_size=config.email_batch_size,
        user_id=config.gmail_user_id
    )

    if not message_stubs:
        logger.info("No emails found for %s (%s). Skipping.", employee.name, employee.email)
        return

    # Create employee folder structure directly under EMPLOYEES_PARENT_DIR_PATH
    employee_dir, processing_dir = ensure_employee_folders(
        config.employees_parent_dir,
        employee.folder_name,
        config.processing_folder_name
    )

    # Determine next batch CSV number
    batch_number = get_next_batch_number(employee_dir)
    logger.info("Batch run #%d for %s", batch_number, employee.name)

    # Process each email message
    csv_records = []
    total_attachments = 0

    for msg_stub in message_stubs:
        msg_id = msg_stub["id"]

        try:
            details = get_message_details(service, msg_id, user_id=config.gmail_user_id)
        except Exception as e:
            logger.error("Failed to get details for message %s: %s", msg_id, e)
            continue

        attachment_parts = details["attachment_parts"]
        num_attachments = len(attachment_parts)

        # Download each attachment
        for part in attachment_parts:
            try:
                download_attachment(
                    service,
                    message_id=msg_id,
                    attachment_id=part["attachment_id"],
                    filename=part["filename"],
                    output_dir=processing_dir,
                    user_id=config.gmail_user_id
                )
                total_attachments += 1
            except Exception as e:
                logger.error(
                    "Failed to download attachment '%s' from message %s: %s",
                    part["filename"], msg_id, e
                )

        # Record metadata for CSV
        csv_records.append({
            "gmail_message_id": msg_id,
            "num_attachments": num_attachments,
            "email_datetime": details["email_datetime"],
        })

        logger.info(
            "Downloaded %d attachment(s) from message %s (date: %s)",
            num_attachments, msg_id, details["email_datetime"].strftime("%Y-%m-%d %H:%M")
        )

    # Write the batch CSV file
    if csv_records:
        write_download_csv(employee_dir, batch_number, csv_records)

    logger.info(
        "Completed %s: %d emails processed, %d attachments downloaded",
        employee.name, len(csv_records), total_attachments
    )


if __name__ == "__main__":
    main()
