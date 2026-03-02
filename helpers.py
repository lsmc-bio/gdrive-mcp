"""Shared helpers used across tool modules."""

import re
import sys
from typing import Optional, List

from services import get_drive


# ── MIME type labels ─────────────────────────────────────────────────────────

GOOGLE_MIME_LABELS = {
    "application/vnd.google-apps.document": "Google Doc",
    "application/vnd.google-apps.spreadsheet": "Google Sheet",
    "application/vnd.google-apps.presentation": "Google Slides",
    "application/vnd.google-apps.folder": "Folder",
    "application/vnd.google-apps.form": "Google Form",
    "application/vnd.google-apps.drawing": "Google Drawing",
}

FILE_FIELDS = "id, name, mimeType, modifiedTime, createdTime, owners, webViewLink, parents, size, shared, trashed"


# ── Drive helpers ────────────────────────────────────────────────────────────

def format_file_entry(f: dict, show_snippet: bool = False) -> str:
    """Format a single Drive file as a readable line."""
    mime = f.get("mimeType", "")
    label = GOOGLE_MIME_LABELS.get(mime, mime.split("/")[-1] if "/" in mime else "file")
    name = f.get("name", "untitled")
    fid = f.get("id", "")
    modified = f.get("modifiedTime", "")[:10]
    link = f.get("webViewLink", "")
    owners = ", ".join(o.get("displayName", "?") for o in f.get("owners", []))

    line = f"**{name}** ({label})\n"
    line += f"  ID: `{fid}`"
    if modified:
        line += f" | Modified: {modified}"
    if owners:
        line += f" | Owner: {owners}"
    if link:
        line += f"\n  Link: {link}"

    return line


def drive_query_files(query: str, max_results: int = 20, order_by: str = "modifiedTime desc") -> list:
    """Execute a Drive files.list query and return file metadata."""
    try:
        results = get_drive().files().list(
            q=query,
            pageSize=max_results,
            fields=f"files({FILE_FIELDS})",
            orderBy=order_by,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        return results.get("files", [])
    except Exception as e:
        print(f"Drive query error: {e}", file=sys.stderr)
        return []


# ── Docs helpers ─────────────────────────────────────────────────────────────

def get_doc_plain_text(doc: dict) -> str:
    """Extract plain text from a Google Docs document structure."""
    text = ""
    for element in doc.get("body", {}).get("content", []):
        if "paragraph" in element:
            for elem in element["paragraph"].get("elements", []):
                if "textRun" in elem:
                    text += elem["textRun"].get("content", "")
        elif "table" in element:
            for row in element["table"].get("tableRows", []):
                for cell in row.get("tableCells", []):
                    for content in cell.get("content", []):
                        if "paragraph" in content:
                            for elem in content["paragraph"].get("elements", []):
                                if "textRun" in elem:
                                    text += elem["textRun"].get("content", "")
    return text


def find_heading_end_index(doc: dict, heading_text: str) -> Optional[int]:
    """Find the end index of a heading element matching the given text."""
    for element in doc.get("body", {}).get("content", []):
        if "paragraph" in element:
            para = element["paragraph"]
            style = para.get("paragraphStyle", {}).get("namedStyleType", "NORMAL_TEXT")
            if style.startswith("HEADING_"):
                para_text = ""
                for elem in para.get("elements", []):
                    if "textRun" in elem:
                        para_text += elem["textRun"].get("content", "")
                if heading_text.strip().lower() in para_text.strip().lower():
                    return element.get("endIndex")
    return None


def find_heading_section_range(doc: dict, heading_text: str) -> Optional[tuple]:
    """Find the (start_index, end_index) of all content under a heading.

    Returns the range from the heading's start to the start of the next heading
    at the same or higher level, or end of document.
    """
    content = doc.get("body", {}).get("content", [])
    found_level = None
    section_start = None

    for element in content:
        if "paragraph" in element:
            para = element["paragraph"]
            style = para.get("paragraphStyle", {}).get("namedStyleType", "NORMAL_TEXT")

            if style.startswith("HEADING_"):
                level = int(style.split("_")[1])
                para_text = ""
                for elem in para.get("elements", []):
                    if "textRun" in elem:
                        para_text += elem["textRun"].get("content", "")

                if section_start is not None and level <= found_level:
                    # Hit a heading at same or higher level — section ends here
                    return (section_start, element.get("startIndex"))

                if heading_text.strip().lower() in para_text.strip().lower():
                    found_level = level
                    section_start = element.get("startIndex")

    if section_start is not None:
        # Section goes to end of document
        last = content[-1]
        return (section_start, last.get("endIndex", section_start + 1) - 1)

    return None


def find_text_indices(doc: dict, search_text: str) -> List[tuple]:
    """Find all (start_index, end_index) pairs for a text string in a doc."""
    results = []
    for element in doc.get("body", {}).get("content", []):
        if "paragraph" in element:
            for elem in element["paragraph"].get("elements", []):
                if "textRun" in elem:
                    content = elem["textRun"].get("content", "")
                    start = elem.get("startIndex", 0)
                    idx = 0
                    while True:
                        pos = content.find(search_text, idx)
                        if pos == -1:
                            break
                        abs_start = start + pos
                        abs_end = abs_start + len(search_text)
                        results.append((abs_start, abs_end))
                        idx = pos + 1
    return results


# ── Sheets helpers ───────────────────────────────────────────────────────────

def hex_to_color(hex_str: str) -> dict:
    """Convert hex color string to Google Sheets/Docs color dict."""
    hex_str = hex_str.lstrip("#")
    r = int(hex_str[0:2], 16) / 255.0
    g = int(hex_str[2:4], 16) / 255.0
    b = int(hex_str[4:6], 16) / 255.0
    return {"red": r, "green": g, "blue": b}


def a1_to_grid_range(a1_range: str, sheet_id: int) -> dict:
    """Convert A1 notation to GridRange dict."""
    match = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", a1_range.upper())
    if not match:
        match = re.match(r"([A-Z]+)(\d+)", a1_range.upper())
        if match:
            col = sum((ord(c) - ord("A") + 1) * (26 ** i) for i, c in enumerate(reversed(match.group(1)))) - 1
            row = int(match.group(2)) - 1
            return {"sheetId": sheet_id, "startRowIndex": row, "endRowIndex": row + 1, "startColumnIndex": col, "endColumnIndex": col + 1}
        raise ValueError(f"Invalid A1 range: {a1_range}")

    start_col = sum((ord(c) - ord("A") + 1) * (26 ** i) for i, c in enumerate(reversed(match.group(1)))) - 1
    start_row = int(match.group(2)) - 1
    end_col = sum((ord(c) - ord("A") + 1) * (26 ** i) for i, c in enumerate(reversed(match.group(3)))) - 1
    end_row = int(match.group(4)) - 1
    return {"sheetId": sheet_id, "startRowIndex": start_row, "endRowIndex": end_row + 1, "startColumnIndex": start_col, "endColumnIndex": end_col + 1}


def get_sheet_id(spreadsheet: dict, sheet_name: Optional[str]) -> int:
    """Get the sheetId for a named tab, or the first tab."""
    sheets = spreadsheet.get("sheets", [])
    if sheet_name:
        for s in sheets:
            if s["properties"]["title"] == sheet_name:
                return s["properties"]["sheetId"]
        raise ValueError(f"Sheet '{sheet_name}' not found.")
    return sheets[0]["properties"]["sheetId"]
