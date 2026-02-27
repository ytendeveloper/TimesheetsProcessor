"""
Logging setup with rotating file handler.

Configures application-wide logging to:
- Console: INFO and above
- File: DEBUG and above, rotated at 5 MB, 7 backups kept
- Log files stored in logs/ directory
"""

import logging
import logging.handlers
from pathlib import Path


def setup_logging(log_dir: str = "logs",
                  log_file: str = "timesheets_processor.log",
                  level: int = logging.DEBUG) -> logging.Logger:
    """
    Initialize and return the application logger with console + rotating file handlers.

    Args:
        log_dir: Directory to store log files (created if missing).
        log_file: Name of the log file.
        level: Root logger level.

    Returns:
        Configured logging.Logger instance.
    """
    log_path = Path(log_dir)
    log_path.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("timesheets_processor")
    logger.setLevel(level)

    # Prevent duplicate handlers on repeated calls
    if logger.handlers:
        return logger

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(module)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # Rotating file handler: 5 MB per file, keep 7 backups (~40 MB total max)
    file_handler = logging.handlers.RotatingFileHandler(
        log_path / log_file,
        maxBytes=5 * 1024 * 1024,
        backupCount=7,
        encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    # Console handler: INFO and above
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger
