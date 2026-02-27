"""
Configuration loader for Timesheets Processor.

Reads config/config.yaml, validates required fields, applies defaults,
and returns a typed Config dataclass.
"""

import os
from dataclasses import dataclass
from datetime import datetime, date
from pathlib import Path

import yaml


@dataclass
class Config:
    """Validated application configuration."""
    start_date: date
    end_date: date
    employees_parent_dir: Path   # Directory where employee folders are created directly
    employees_list_file: Path
    email_batch_size: int
    processing_folder_name: str
    gmail_credentials_file: Path
    gmail_token_file: Path
    gmail_user_id: str


def load_config(config_path: str = "config/config.yaml") -> Config:
    """
    Load and validate configuration from a YAML file.

    Args:
        config_path: Path to the YAML config file.

    Returns:
        Config dataclass with all fields validated and defaults applied.

    Raises:
        FileNotFoundError: If config file does not exist.
        ValueError: If START_DATE is missing or dates are invalid.
    """
    config_file = Path(config_path)
    if not config_file.exists():
        raise FileNotFoundError(f"Configuration file not found: {config_path}")

    with open(config_file, "r") as f:
        raw = yaml.safe_load(f) or {}

    # --- START_DATE (required, MM-DD-YYYY) ---
    start_date_str = raw.get("START_DATE", "").strip()
    if not start_date_str:
        raise ValueError("START_DATE is required in config.yaml (format: MM-DD-YYYY)")
    start_date = _parse_date(start_date_str, "START_DATE")

    # --- END_DATE (optional, defaults to today) ---
    end_date_str = raw.get("END_DATE", "").strip()
    if end_date_str:
        end_date = _parse_date(end_date_str, "END_DATE")
    else:
        end_date = date.today()

    # Validate date range
    if start_date > end_date:
        raise ValueError(
            f"START_DATE ({start_date}) cannot be after END_DATE ({end_date})"
        )

    # --- EMPLOYEES_PARENT_DIR_PATH (defaults to cwd) ---
    # Employee folders are created directly under this path
    parent_dir_str = raw.get("EMPLOYEES_PARENT_DIR_PATH", "").strip()
    employees_parent_dir = Path(parent_dir_str) if parent_dir_str else Path(os.getcwd())

    # --- EMPLOYEES_LIST_FILE (defaults to config/EmployeesList.md) ---
    emp_list_str = raw.get("EMPLOYEES_LIST_FILE", "").strip()
    employees_list_file = Path(emp_list_str) if emp_list_str else Path("config/EmployeesList.md")

    # --- EMAIL_BATCH_SIZE (defaults to 5) ---
    email_batch_size = raw.get("EMAIL_BATCH_SIZE", 5)
    if not isinstance(email_batch_size, int) or email_batch_size < 1:
        raise ValueError(f"EMAIL_BATCH_SIZE must be a positive integer, got: {email_batch_size}")

    # --- PROCESSING_FOLDER_NAME (defaults to 'downloaded') ---
    processing_folder = raw.get("PROCESSING_FOLDER_NAME", "").strip()
    if not processing_folder:
        processing_folder = "downloaded"

    # --- Gmail auth paths ---
    gmail_creds = raw.get("GMAIL_CREDENTIALS_FILE", "config/credentials.json").strip()
    gmail_token = raw.get("GMAIL_TOKEN_FILE", "token.pickle").strip()
    gmail_user_id = raw.get("GMAIL_USER_ID", "me").strip()

    return Config(
        start_date=start_date,
        end_date=end_date,
        employees_parent_dir=employees_parent_dir,
        employees_list_file=employees_list_file,
        email_batch_size=email_batch_size,
        processing_folder_name=processing_folder,
        gmail_credentials_file=Path(gmail_creds),
        gmail_token_file=Path(gmail_token),
        gmail_user_id=gmail_user_id,
    )


def _parse_date(date_str: str, field_name: str) -> date:
    """
    Parse a date string in MM-DD-YYYY format.

    Args:
        date_str: Date string to parse.
        field_name: Config field name (for error messages).

    Returns:
        date object.

    Raises:
        ValueError: If the format is invalid.
    """
    try:
        return datetime.strptime(date_str, "%m-%d-%Y").date()
    except ValueError:
        raise ValueError(
            f"{field_name} must be in MM-DD-YYYY format, got: '{date_str}'"
        )
