"""
Employee list parser for Timesheets Processor.

Parses a Markdown file containing a table of employees with columns:
Employee Name, Employee Email, Project Name.

Generates sanitized folder names by stripping spaces, dots, and special characters.
"""

import logging
import re
from dataclasses import dataclass
from pathlib import Path

logger = logging.getLogger("timesheets_processor")


@dataclass
class Employee:
    """Represents an employee entry from the list file."""
    name: str          # Original name, e.g. "John Doe"
    email: str         # Email address, e.g. "john.doe@example.com"
    project: str       # Project name, e.g. "Project Alpha"
    folder_name: str   # Sanitized folder name, e.g. "JohnDoe_ProjectAlpha"


def sanitize_folder_name(name: str, project: str) -> str:
    """
    Create a filesystem-safe folder name from employee name and project.

    Concatenates name and project with underscore separator,
    then strips all characters except alphanumerics and underscores.

    Args:
        name: Employee name (e.g. "John Doe").
        project: Project name (e.g. "Project Alpha").

    Returns:
        Sanitized string like "JohnDoe_ProjectAlpha".
    """
    combined = f"{name}_{project}"
    # Remove everything except letters, digits, and underscores
    return re.sub(r"[^a-zA-Z0-9_]", "", combined)


def parse_employees(file_path: Path) -> list[Employee]:
    """
    Parse a Markdown file containing an employee table.

    Expects a pipe-delimited Markdown table with at least 3 columns:
    Employee Name | Employee Email | Project Name

    Skips header rows (lines containing only dashes/pipes) and blank lines.

    Args:
        file_path: Path to the Markdown employee list file.

    Returns:
        List of Employee dataclass instances.

    Raises:
        FileNotFoundError: If the file does not exist.
    """
    if not file_path.exists():
        raise FileNotFoundError(f"Employees list file not found: {file_path}")

    employees: list[Employee] = []
    content = file_path.read_text(encoding="utf-8")

    for line_num, line in enumerate(content.splitlines(), start=1):
        line = line.strip()

        # Skip empty lines, comments/headings, and separator rows (|---|---|---|)
        if not line or line.startswith("#") or re.match(r"^\|[\s\-|]+\|$", line):
            continue

        # Must be a pipe-delimited table row
        if not line.startswith("|"):
            continue

        # Split by pipe and strip whitespace from each cell
        cells = [cell.strip() for cell in line.split("|")]
        # Remove empty strings from leading/trailing pipes
        cells = [c for c in cells if c]

        if len(cells) < 3:
            logger.warning(
                "Skipping malformed row at line %d: expected at least 3 columns, got %d — '%s'",
                line_num, len(cells), line
            )
            continue

        # Skip the header row (contains column titles)
        if any(c.lower() == "employee name" for c in cells):
            continue

        # Find the email column (the one containing '@')
        email_idx = None
        for idx, c in enumerate(cells):
            if "@" in c:
                email_idx = idx
                break

        if email_idx is None:
            logger.warning(
                "Skipping row at line %d: no email address found — '%s'",
                line_num, line
            )
            continue

        # Name is the column before email, project is the column after
        if email_idx < 1 or email_idx + 1 >= len(cells):
            logger.warning(
                "Skipping row at line %d: unexpected column layout — '%s'",
                line_num, line
            )
            continue

        name = cells[email_idx - 1]
        email = cells[email_idx]
        project = cells[email_idx + 1]

        # Use empty project as-is (default to underscore placeholder)
        if not project:
            project = "_"

        folder = sanitize_folder_name(name, project)
        emp = Employee(name=name, email=email, project=project, folder_name=folder)
        employees.append(emp)
        logger.debug("Parsed employee: %s (%s) → folder: %s", name, email, folder)

    logger.info("Loaded %d employees from %s", len(employees), file_path)
    return employees
