"""
Google OAuth2 authentication helper for the GDrive MCP server.

First-time setup:
1. Go to https://console.cloud.google.com/
2. Create a project, enable Google Drive API + Google Sheets API + Google Docs API
3. Create OAuth 2.0 credentials (Desktop app)
4. Download the JSON and save it as 'credentials.json' next to this file
5. Run: python auth.py
6. Browser opens, sign in, authorize — tokens are saved to 'token.json'
"""

import json
import os
from pathlib import Path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# Where to look for credentials and store tokens
CONFIG_DIR = Path(__file__).parent
CREDENTIALS_FILE = CONFIG_DIR / "credentials.json"
TOKEN_FILE = CONFIG_DIR / "token.json"

# Scopes needed — full read-write access to Drive, Docs, and Sheets
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/script.projects",
]


def get_credentials() -> Credentials:
    """Get valid Google OAuth2 credentials, refreshing or re-authenticating as needed."""
    creds = None

    # Load existing token
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    # Refresh or re-authenticate
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_FILE.exists():
                raise FileNotFoundError(
                    f"Missing {CREDENTIALS_FILE}. Download OAuth credentials from "
                    f"Google Cloud Console and save as 'credentials.json' in {CONFIG_DIR}"
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), SCOPES
            )
            creds = flow.run_local_server(port=3334)

        # Save for next time
        TOKEN_FILE.write_text(creds.to_json())

    return creds


if __name__ == "__main__":
    print("Authenticating with Google...")
    creds = get_credentials()
    print(f"✅ Authenticated! Token saved to {TOKEN_FILE}")
    print("You can now run the MCP server.")
