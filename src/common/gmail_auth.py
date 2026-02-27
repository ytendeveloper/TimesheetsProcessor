"""
Gmail API OAuth2 authentication for Timesheets Processor.

Handles the OAuth2 flow:
- Loads existing token from pickle file if available
- Refreshes expired tokens automatically
- Runs interactive browser consent flow on first use
- Saves tokens for future sessions

Scope: gmail.modify — read and modify access (move emails, manage labels). Cannot delete emails.
"""

import logging
import pickle
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

logger = logging.getLogger("timesheets_processor")

# Modify scope allows reading, labeling, and moving emails (but NOT deleting)
SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]


def authenticate(credentials_file: Path, token_file: Path):
    """
    Authenticate with Gmail API and return a service object.

    On first run, opens a browser window for user consent. Subsequent runs
    use the saved token, refreshing it automatically if expired.

    Args:
        credentials_file: Path to the Google OAuth2 credentials.json file.
        token_file: Path to save/load the pickled auth token.

    Returns:
        googleapiclient.discovery.Resource — authenticated Gmail API service.

    Raises:
        FileNotFoundError: If credentials.json is missing.
        Exception: If authentication fails after all attempts.
    """
    if not credentials_file.exists():
        raise FileNotFoundError(
            f"Gmail credentials file not found: {credentials_file}\n"
            "Download it from Google Cloud Console → APIs & Services → Credentials → "
            "OAuth 2.0 Client IDs → Download JSON, and save as config/credentials.json"
        )

    creds = None

    # Load existing token if available
    if token_file.exists():
        logger.debug("Loading existing token from %s", token_file)
        with open(token_file, "rb") as f:
            creds = pickle.load(f)

    # Refresh or run consent flow if needed
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            logger.info("Token expired, refreshing...")
            creds.refresh(Request())
            logger.info("Token refreshed successfully")
        else:
            logger.info("No valid token found, starting OAuth2 consent flow...")
            logger.info("A browser window will open for Gmail authorization")
            flow = InstalledAppFlow.from_client_secrets_file(
                str(credentials_file), SCOPES
            )
            creds = flow.run_local_server(port=0)
            logger.info("OAuth2 consent completed successfully")

        # Save the token for future runs
        with open(token_file, "wb") as f:
            pickle.dump(creds, f)
        logger.debug("Token saved to %s", token_file)

    service = build("gmail", "v1", credentials=creds)
    logger.info("Gmail API authenticated successfully (modify scope)")
    return service
