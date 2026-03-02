"""Lazy-initialized Google API service clients."""

from googleapiclient.discovery import build
from auth import get_credentials

_drive_service = None
_docs_service = None
_sheets_service = None
_slides_service = None
_scripts_service = None
_gmail_service = None
_calendar_service = None


def get_drive():
    global _drive_service
    if _drive_service is None:
        _drive_service = build("drive", "v3", credentials=get_credentials())
    return _drive_service


def get_docs():
    global _docs_service
    if _docs_service is None:
        _docs_service = build("docs", "v1", credentials=get_credentials())
    return _docs_service


def get_sheets():
    global _sheets_service
    if _sheets_service is None:
        _sheets_service = build("sheets", "v4", credentials=get_credentials())
    return _sheets_service


def get_slides():
    global _slides_service
    if _slides_service is None:
        _slides_service = build("slides", "v1", credentials=get_credentials())
    return _slides_service


def get_scripts():
    global _scripts_service
    if _scripts_service is None:
        _scripts_service = build("script", "v1", credentials=get_credentials())
    return _scripts_service


def get_gmail():
    global _gmail_service
    if _gmail_service is None:
        _gmail_service = build("gmail", "v1", credentials=get_credentials())
    return _gmail_service


def get_calendar():
    global _calendar_service
    if _calendar_service is None:
        _calendar_service = build("calendar", "v3", credentials=get_credentials())
    return _calendar_service
