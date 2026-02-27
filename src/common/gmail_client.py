"""
Gmail API client for Timesheets Processor.

Provides functions to:
- Search emails by sender, date range, and attachment presence
- Retrieve email metadata (message ID, date, attachment count)
- Download attachments to local filesystem

All operations are non-destructive. Emails are never deleted.
"""

import base64
import logging
import random
import time
from datetime import date, timedelta
from pathlib import Path
from typing import Any

from googleapiclient.errors import HttpError

logger = logging.getLogger("timesheets_processor")

# Retry config for rate limit / transient errors
MAX_RETRIES = 5
RETRYABLE_STATUS_CODES = (429, 500, 503)


def _execute_with_backoff(request, description: str = "API call") -> Any:
    """
    Execute a Gmail API request with exponential backoff on transient errors.

    Args:
        request: A googleapiclient HttpRequest object.
        description: Human-readable description for log messages.

    Returns:
        The API response.

    Raises:
        HttpError: If a non-retryable error occurs or retries are exhausted.
    """
    for attempt in range(MAX_RETRIES):
        try:
            return request.execute()
        except HttpError as e:
            if e.resp.status in RETRYABLE_STATUS_CODES:
                wait = (2 ** attempt) + random.random()
                logger.warning(
                    "%s hit HTTP %d, retrying in %.1fs (attempt %d/%d)",
                    description, e.resp.status, wait, attempt + 1, MAX_RETRIES
                )
                time.sleep(wait)
            else:
                raise
    raise RuntimeError(f"{description} failed after {MAX_RETRIES} retries")


def search_emails(service, sender_email: str, start_date: date,
                  end_date: date, batch_size: int,
                  user_id: str = "me") -> list[dict]:
    """
    Search Gmail for emails from a sender within a date range that have attachments.

    Args:
        service: Authenticated Gmail API service object.
        sender_email: Email address of the sender to search for.
        start_date: Start of date range (inclusive).
        end_date: End of date range (inclusive).
        batch_size: Maximum number of email stubs to return.
        user_id: Gmail user ID ('me' for authenticated user).

    Returns:
        List of message stubs: [{"id": "...", "threadId": "..."}, ...]
    """
    # Gmail query dates use YYYY/MM/DD format (slashes, not dashes)
    # 'before:' is exclusive, so add 1 day to make end_date inclusive
    after_str = start_date.strftime("%Y/%m/%d")
    before_str = (end_date + timedelta(days=1)).strftime("%Y/%m/%d")

    query = f"from:{sender_email} after:{after_str} before:{before_str} has:attachment"
    logger.info("Gmail query: %s", query)

    response = _execute_with_backoff(
        service.users().messages().list(
            userId=user_id,
            q=query,
            maxResults=batch_size
        ),
        description=f"Search emails from {sender_email}"
    )

    messages = response.get("messages", [])
    logger.info("Found %d emails from %s", len(messages), sender_email)
    return messages


def get_message_details(service, message_id: str,
                        user_id: str = "me") -> dict:
    """
    Fetch full message payload to extract metadata and attachment info.

    Args:
        service: Authenticated Gmail API service object.
        message_id: Gmail unique message ID.
        user_id: Gmail user ID.

    Returns:
        Dict with keys: message_id, email_datetime, attachment_parts, subject.
    """
    msg = _execute_with_backoff(
        service.users().messages().get(
            userId=user_id,
            id=message_id,
            format="full"
        ),
        description=f"Get message {message_id}"
    )

    # Extract date from internalDate (milliseconds since epoch)
    internal_date_ms = int(msg.get("internalDate", 0))
    from datetime import datetime, timezone
    email_datetime = datetime.fromtimestamp(
        internal_date_ms / 1000, tz=timezone.utc
    )

    # Extract subject from headers
    headers = {h["name"]: h["value"] for h in msg["payload"].get("headers", [])}
    subject = headers.get("Subject", "(no subject)")

    # Recursively find all attachment parts
    attachment_parts = _get_attachment_parts(msg["payload"])

    logger.debug(
        "Message %s: date=%s, subject='%s', attachments=%d",
        message_id, email_datetime.isoformat(), subject, len(attachment_parts)
    )

    return {
        "message_id": message_id,
        "email_datetime": email_datetime,
        "subject": subject,
        "attachment_parts": attachment_parts,
    }


def _get_attachment_parts(payload: dict) -> list[dict]:
    """
    Recursively walk the MIME parts tree to find all attachments.

    Gmail nests parts inside multipart/* containers, so we must recurse
    to find all attachments in complex email structures.

    Args:
        payload: The message payload dict from Gmail API.

    Returns:
        List of dicts with: filename, mime_type, size, attachment_id.
    """
    parts = []

    for part in payload.get("parts", []):
        filename = part.get("filename", "")
        attachment_id = part.get("body", {}).get("attachmentId")

        if filename and attachment_id:
            parts.append({
                "filename": filename,
                "mime_type": part.get("mimeType", ""),
                "size": part.get("body", {}).get("size", 0),
                "attachment_id": attachment_id,
            })

        # Recurse into nested multipart containers
        if "parts" in part:
            parts.extend(_get_attachment_parts(part))

    return parts


def download_attachment(service, message_id: str, attachment_id: str,
                        filename: str, output_dir: Path,
                        user_id: str = "me") -> Path:
    """
    Download a single attachment and save it to disk.

    Files are prefixed with the message ID to prevent filename collisions
    across different emails.

    Args:
        service: Authenticated Gmail API service object.
        message_id: Gmail message ID the attachment belongs to.
        attachment_id: Gmail attachment ID.
        filename: Original attachment filename.
        output_dir: Directory to save the file into.
        user_id: Gmail user ID.

    Returns:
        Path to the saved file.
    """
    response = _execute_with_backoff(
        service.users().messages().attachments().get(
            userId=user_id,
            messageId=message_id,
            id=attachment_id
        ),
        description=f"Download attachment '{filename}' from message {message_id}"
    )

    # Gmail uses URL-safe base64 encoding (- and _ instead of + and /)
    file_data = base64.urlsafe_b64decode(response["data"])

    # Strip path traversal from filename, prefix with message ID for uniqueness
    safe_name = Path(filename).name
    file_path = output_dir / f"{message_id}_{safe_name}"
    file_path.write_bytes(file_data)

    logger.info(
        "Saved attachment: %s (%d bytes) to %s",
        safe_name, len(file_data), file_path
    )
    return file_path


def ensure_label(service, label_name: str, user_id: str = "me") -> str:
    """
    Get or create a Gmail label by name.

    Args:
        service: Authenticated Gmail API service object.
        label_name: The label name to find or create (case-sensitive).
        user_id: Gmail user ID.

    Returns:
        The label ID string.
    """
    response = _execute_with_backoff(
        service.users().labels().list(userId=user_id),
        description="List labels"
    )

    for label in response.get("labels", []):
        if label["name"].lower() == label_name.lower():
            logger.info("Found existing label '%s' (id=%s)", label["name"], label["id"])
            return label["id"]

    # Label not found — create it
    label_body = {
        "name": label_name,
        "labelListVisibility": "labelShow",
        "messageListVisibility": "show",
    }
    created = _execute_with_backoff(
        service.users().labels().create(userId=user_id, body=label_body),
        description=f"Create label '{label_name}'"
    )

    label_id = created["id"]
    logger.info("Created new label '%s' (id=%s)", label_name, label_id)
    return label_id


def move_to_label(service, message_id: str, label_id: str,
                  user_id: str = "me") -> None:
    """
    Move a message from INBOX to the specified label.

    Removes the INBOX label and adds the target label.

    Args:
        service: Authenticated Gmail API service object.
        message_id: Gmail message ID to modify.
        label_id: Label ID to apply.
        user_id: Gmail user ID.
    """
    body = {
        "removeLabelIds": ["INBOX"],
        "addLabelIds": [label_id],
    }
    _execute_with_backoff(
        service.users().messages().modify(
            userId=user_id, id=message_id, body=body
        ),
        description=f"Move message {message_id} to label {label_id}"
    )
    logger.debug("Moved message %s to label %s", message_id, label_id)
