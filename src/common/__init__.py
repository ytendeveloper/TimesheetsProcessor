"""
Module #0 — Common/Init for Timesheets Processor.

Shared foundation used by all job executables. Provides:
- Configuration loading and validation
- Logging with rotation
- Employee list parsing
- Gmail API authentication
- Gmail search and attachment download
- File and folder management

Usage from any job:
    # Jobs that need Gmail access:
    from src.common import init_app
    logger, config, employees, service = init_app()

    # Jobs that only need local file access (no Gmail auth):
    from src.common import init_app_light
    logger, config = init_app_light()
"""

import sys
import logging

from src.common.logger_setup import setup_logging
from src.common.config_loader import load_config
from src.common.employee_parser import parse_employees
from src.common.gmail_auth import authenticate

logger = logging.getLogger("timesheets_processor")


def init_app():
    """
    Standard initialization sequence for all jobs.

    Performs:
    1. Setup rotating log handlers
    2. Load and validate config/config.yaml
    3. Parse the employee list file
    4. Authenticate with Gmail API (read-only)

    Returns:
        Tuple of (logger, config, employees, gmail_service).

    Exits:
        Calls sys.exit(1) on critical config or auth failures.
    """
    # Step 1: Logging
    app_logger = setup_logging()

    # Step 2: Configuration
    try:
        config = load_config()
    except (FileNotFoundError, ValueError) as e:
        app_logger.critical("Configuration error: %s", e)
        sys.exit(1)

    app_logger.info(
        "Config loaded: START_DATE=%s, END_DATE=%s, BATCH_SIZE=%d, "
        "PROCESSING_FOLDER=%s",
        config.start_date, config.end_date,
        config.email_batch_size, config.processing_folder_name
    )

    # Step 3: Employee list
    try:
        employees = parse_employees(config.employees_list_file)
    except FileNotFoundError as e:
        app_logger.critical("Employee list error: %s", e)
        sys.exit(1)

    if not employees:
        app_logger.warning(
            "No employees found in %s. Nothing to process.",
            config.employees_list_file
        )
        sys.exit(0)

    # Step 4: Gmail authentication
    try:
        service = authenticate(config.gmail_credentials_file, config.gmail_token_file)
    except FileNotFoundError as e:
        app_logger.critical("Authentication error: %s", e)
        sys.exit(1)
    except Exception as e:
        app_logger.critical("Gmail authentication failed: %s", e)
        sys.exit(1)

    return app_logger, config, employees, service


def init_app_light():
    """
    Lightweight initialization for jobs that don't need Gmail access.

    Performs:
    1. Setup rotating log handlers
    2. Load and validate config/config.yaml

    Returns:
        Tuple of (logger, config).

    Exits:
        Calls sys.exit(1) on critical config failures.
    """
    app_logger = setup_logging()

    try:
        config = load_config()
    except (FileNotFoundError, ValueError) as e:
        app_logger.critical("Configuration error: %s", e)
        sys.exit(1)

    app_logger.info(
        "Config loaded: EMPLOYEES_PARENT_DIR=%s, PROCESSING_FOLDER=%s",
        config.employees_parent_dir, config.processing_folder_name
    )

    return app_logger, config
