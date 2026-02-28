# gdrive-mcp Enhancement Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add Gmail (read-only), Calendar (full CRUD), enhanced Docs formatting, and expanded Apps Script management to the gdrive-mcp server, refactored into domain-specific modules.

**Architecture:** Refactor the monolithic 2927-line `server.py` into a `tools/` directory with per-domain modules (drive, docs, sheets, slides, comments, gmail, calendar, scripts). Shared services and helpers extracted to `services.py` and `helpers.py`. All modules register tools against a shared FastMCP instance.

**Tech Stack:** Python 3.10+, FastMCP (mcp[cli]>=1.0.0), google-api-python-client, Pydantic v2, Google APIs (Drive v3, Docs v1, Sheets v4, Slides v1, Gmail v1, Calendar v3, Script v1)

---

## Phase 1: Refactor Into Modules

### Task 1: Create services.py — Extract service getters

**Files:**
- Create: `services.py`

**Step 1: Create services.py with all service getters**

```python
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
```

**Step 2: Commit**

```bash
git add services.py
git commit -m "refactor: extract service getters to services.py"
```

### Task 2: Create helpers.py — Extract shared utilities

**Files:**
- Create: `helpers.py`

**Step 1: Create helpers.py with all shared helpers**

Extract from `server.py` lines 101-209. These are used across multiple tool modules:

```python
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

    Returns the range from the heading's end to the start of the next heading
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
```

**Step 2: Commit**

```bash
git add helpers.py
git commit -m "refactor: extract shared helpers to helpers.py"
```

### Task 3: Create tools/ directory and split existing tools

**Files:**
- Create: `tools/__init__.py`
- Create: `tools/drive.py` — lines 212-971 (search, read_doc, read_sheet, list_folder, recent, file_info)
- Create: `tools/docs.py` — lines 973-1453 (find_replace, insert_text, delete_text, create_doc) + lines 1591-1800 (read_section, insert_table)
- Create: `tools/sheets.py` — lines 1453-1591 (write_sheet, append_sheet) + lines 1800-2270 (format_cells, manage_sheets, create_sheet, insert_rows_cols, delete_rows_cols)
- Create: `tools/slides.py` — lines 2635-2747 (read_slides)
- Create: `tools/comments.py` — lines 2483-2575 (comments)
- Create: `tools/scripts.py` — lines 2748-2921 (run_script, create_script)
- Create: `tools/management.py` — lines 2270-2483 (move_copy, share, export) + lines 2577-2635 (versions)
- Modify: `server.py` — rewrite as thin coordinator

This is a large mechanical refactor. The pattern for each tool module file is:

```python
"""[Domain] tools for gdrive-mcp."""

from pydantic import BaseModel, Field, ConfigDict
from typing import Optional, List
from enum import Enum

from services import get_docs  # whichever service(s) needed
from helpers import find_heading_end_index, find_text_indices  # whichever helpers needed


def register(mcp):
    """Register all [domain] tools with the MCP server."""

    class SomeInput(BaseModel):
        # ... (moved from server.py)

    @mcp.tool(name="gdrive_tool_name", annotations={...})
    async def gdrive_tool_name(params: SomeInput) -> str:
        # ... (moved from server.py)
```

Each module exports a `register(mcp)` function. The `__init__.py` collects them:

```python
"""Tool modules for gdrive-mcp."""

from . import drive, docs, sheets, slides, comments, scripts, management


def register_all(mcp):
    """Register all tool modules."""
    drive.register(mcp)
    docs.register(mcp)
    sheets.register(mcp)
    slides.register(mcp)
    comments.register(mcp)
    scripts.register(mcp)
    management.register(mcp)
```

**Step 1: Create tools/__init__.py**

Create the file with the `register_all` function above.

**Step 2: Create tools/drive.py**

Move from `server.py`:
- `FileType` enum (line 214)
- `FILE_TYPE_QUERIES` dict (line 223)
- `SearchInput` + `gdrive_search` (lines 234-391)
- `ReadDocInput` + `gdrive_read_doc` (lines 395-589)
- `ReadSheetInput` + `gdrive_read_sheet` (lines 593-717)
- `ListFolderInput` + `gdrive_list_folder` (lines 721-806)
- `RecentFilesInput` + `gdrive_recent` (lines 810-881)
- `FileInfoInput` + `gdrive_file_info` (lines 885-971)

Imports needed: `get_drive, get_docs, get_sheets` from services, `format_file_entry, drive_query_files, GOOGLE_MIME_LABELS, FILE_FIELDS` from helpers.

**Step 3: Create tools/docs.py**

Move from `server.py`:
- `InsertLocation` enum (line 1058)
- `FindReplaceInput` + `gdrive_find_replace` (lines 975-1054)
- `InsertTextInput` + `gdrive_insert_text` (lines 1065-1242)
- `DeleteTextInput` + `gdrive_delete_text` (lines 1246-1359)
- `CreateDocInput` + `gdrive_create_doc` (lines 1363-1451)
- `ReadSectionInput` + `gdrive_read_section` (lines 1593-1674)
- `InsertTableInput` + `gdrive_insert_table` (lines 1678-1798)

Imports needed: `get_docs` from services, `find_heading_end_index, find_text_indices, get_doc_plain_text` from helpers. Also `markdown` library and `MediaInMemoryUpload` from googleapiclient.

**Step 4: Create tools/sheets.py**

Move from `server.py`:
- `WriteSheetInput` + `gdrive_write_sheet` (lines 1455-1519)
- `AppendSheetInput` + `gdrive_append_sheet` (lines 1523-1589)
- `FormatCellsInput` + `gdrive_format_cells` (lines 1802-2010)
- `SheetAction` enum + `ManageSheetsInput` + `gdrive_manage_sheets` (lines 2014-2095)
- `CreateSheetInput` + `gdrive_create_sheet` (lines 2099-2158)
- `InsertRowsColsInput` + `gdrive_insert_rows_cols` (lines 2162-2212)
- `DeleteRowsColsInput` + `gdrive_delete_rows_cols` (lines 2216-2268)

Imports needed: `get_sheets` from services, `hex_to_color, a1_to_grid_range, get_sheet_id` from helpers.

**Step 5: Create tools/slides.py**

Move from `server.py`:
- `ReadSlidesInput` + `gdrive_read_slides` (lines 2637-2746)

Imports needed: `get_slides` from services.

**Step 6: Create tools/comments.py**

Move from `server.py`:
- `CommentsInput` + `gdrive_comments` (lines 2485-2575)

Imports needed: `get_drive` from services.

**Step 7: Create tools/management.py**

Move from `server.py`:
- `MoveCopyInput` + `gdrive_move_copy` (lines 2272-2337)
- `ShareInput` + `gdrive_share` (lines 2341-2401)
- `ExportInput` + `gdrive_export` (lines 2405-2481)
- `VersionsInput` + `gdrive_versions` (lines 2579-2633)

Imports needed: `get_drive` from services, `format_file_entry` from helpers.

**Step 8: Create tools/scripts.py**

Move from `server.py`:
- `RunScriptInput` + `gdrive_run_script` (lines 2750-2834)
- `CreateScriptInput` + `gdrive_create_script` (lines 2838-2921)

Imports needed: `get_scripts` from services.

**Step 9: Rewrite server.py as thin coordinator**

```python
"""
Google Drive MCP Server for Claude

Run with:  python server.py
Or via Claude Code config pointing to this file.
"""

from mcp.server.fastmcp import FastMCP
from tools import register_all

mcp = FastMCP("gdrive_mcp")
register_all(mcp)

if __name__ == "__main__":
    mcp.run()
```

**Step 10: Smoke test — verify MCP starts without errors**

Run: `cd /home/ageller/gdrive-mcp && timeout 5 python server.py 2>&1 || true`

Expected: Server starts without import errors. It will block on stdin (MCP protocol), so timeout is fine.

**Step 11: Commit**

```bash
git add tools/ server.py
git commit -m "refactor: split server.py into domain-specific tool modules

Reorganize 2927-line monolith into tools/ directory:
- tools/drive.py: search, read_doc, read_sheet, list_folder, recent, file_info
- tools/docs.py: find_replace, insert_text, delete_text, create_doc, read_section, insert_table
- tools/sheets.py: write_sheet, append_sheet, format_cells, manage_sheets, create_sheet, rows/cols
- tools/slides.py: read_slides
- tools/comments.py: comments
- tools/management.py: move_copy, share, export, versions
- tools/scripts.py: run_script, create_script
- services.py: lazy-initialized Google API clients
- helpers.py: shared utilities

All 27 existing tools preserved with identical behavior."
```

---

## Phase 2: Enhanced Docs Formatting

### Task 4: Extend gdrive_insert_text with rich formatting

**Files:**
- Modify: `tools/docs.py` (InsertTextInput class and gdrive_insert_text function)

**Step 1: Add new fields to InsertTextInput**

Add these fields after the existing `heading_level` field:

```python
    underline: bool = Field(
        default=False,
        description="If true, the inserted text will be underlined.",
    )
    strikethrough: bool = Field(
        default=False,
        description="If true, the inserted text will have strikethrough.",
    )
    font_family: Optional[str] = Field(
        default=None,
        description="Font family for inserted text (e.g., 'Arial', 'Times New Roman', 'Courier New').",
    )
    font_size: Optional[int] = Field(
        default=None,
        description="Font size in points for inserted text.",
        ge=1,
        le=400,
    )
    text_color: Optional[str] = Field(
        default=None,
        description="Text color as hex (e.g., '#FF0000' for red).",
    )
    bg_color: Optional[str] = Field(
        default=None,
        description="Background/highlight color as hex (e.g., '#FFFF00' for yellow).",
    )
    link_url: Optional[str] = Field(
        default=None,
        description="URL to make the inserted text a hyperlink.",
    )
```

**Step 2: Update the text styling section of gdrive_insert_text**

Replace the existing bold/italic style block (the `if params.bold or params.italic:` section) with:

```python
    # Apply text styling
    style = {}
    fields = []
    if params.bold:
        style["bold"] = True
        fields.append("bold")
    if params.italic:
        style["italic"] = True
        fields.append("italic")
    if params.underline:
        style["underline"] = True
        fields.append("underline")
    if params.strikethrough:
        style["strikethrough"] = True
        fields.append("strikethrough")
    if params.font_family:
        style["weightedFontFamily"] = {"fontFamily": params.font_family}
        fields.append("weightedFontFamily")
    if params.font_size:
        style["fontSize"] = {"magnitude": params.font_size, "unit": "PT"}
        fields.append("fontSize")
    if params.text_color:
        style["foregroundColor"] = {"color": {"rgbColor": hex_to_color(params.text_color)}}
        fields.append("foregroundColor")
    if params.bg_color:
        style["backgroundColor"] = {"color": {"rgbColor": hex_to_color(params.bg_color)}}
        fields.append("backgroundColor")
    if params.link_url:
        style["link"] = {"url": params.link_url}
        fields.append("link")

    if style:
        requests.append({
            "updateTextStyle": {
                "range": {
                    "startIndex": insert_index,
                    "endIndex": insert_index + len(text_to_insert),
                },
                "textStyle": style,
                "fields": ",".join(fields),
            }
        })
```

Add `hex_to_color` to the imports from helpers.

Also update the formatting summary at the end to include the new options:

```python
        formatting = []
        if params.bold:
            formatting.append("bold")
        if params.italic:
            formatting.append("italic")
        if params.underline:
            formatting.append("underline")
        if params.strikethrough:
            formatting.append("strikethrough")
        if params.font_family:
            formatting.append(f"font: {params.font_family}")
        if params.font_size:
            formatting.append(f"{params.font_size}pt")
        if params.text_color:
            formatting.append(f"color: {params.text_color}")
        if params.bg_color:
            formatting.append(f"bg: {params.bg_color}")
        if params.link_url:
            formatting.append("linked")
        if params.heading_level:
            formatting.append(f"heading {params.heading_level}")
        fmt_str = f" ({', '.join(formatting)})" if formatting else ""
```

Update the docstring to mention the new formatting options.

**Step 3: Commit**

```bash
git add tools/docs.py
git commit -m "feat: extend gdrive_insert_text with rich formatting

Add underline, strikethrough, font_family, font_size, text_color,
bg_color, and link_url options to text insertion."
```

### Task 5: Create gdrive_format_text tool

**Files:**
- Modify: `tools/docs.py` (add new tool at the end)

**Step 1: Add the FormatTextTarget enum and FormatTextInput model**

```python
class FormatTextTarget(str, Enum):
    MATCH = "match"
    RANGE = "range"
    HEADING = "heading"


class FormatTextInput(BaseModel):
    """Input for formatting existing text in a Google Doc."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(..., description="The Google Drive file ID of the document.", min_length=1)
    target: FormatTextTarget = Field(
        ...,
        description="How to select text: 'match' (find by text), 'range' (by character index), 'heading' (all content under a heading).",
    )

    # Match mode
    match_text: Optional[str] = Field(
        default=None,
        description="Text to find in the document (required for target='match').",
    )
    match_occurrence: Optional[int] = Field(
        default=None,
        description="Which occurrence to format: None = all, 1 = first, -1 = last, N = Nth occurrence.",
    )

    # Range mode
    start_index: Optional[int] = Field(default=None, description="Start character index (required for target='range').", ge=1)
    end_index: Optional[int] = Field(default=None, description="End character index (required for target='range').", ge=2)

    # Heading mode
    heading_text: Optional[str] = Field(
        default=None,
        description="Heading text to find (case-insensitive partial match, required for target='heading').",
    )

    # Text formatting (all optional)
    bold: Optional[bool] = Field(default=None, description="Set bold on/off.")
    italic: Optional[bool] = Field(default=None, description="Set italic on/off.")
    underline: Optional[bool] = Field(default=None, description="Set underline on/off.")
    strikethrough: Optional[bool] = Field(default=None, description="Set strikethrough on/off.")
    font_family: Optional[str] = Field(default=None, description="Font family (e.g., 'Arial', 'Georgia').")
    font_size: Optional[int] = Field(default=None, description="Font size in points.", ge=1, le=400)
    text_color: Optional[str] = Field(default=None, description="Text color as hex (e.g., '#FF0000').")
    bg_color: Optional[str] = Field(default=None, description="Background/highlight color as hex.")
    link_url: Optional[str] = Field(default=None, description="URL to set as hyperlink.")
    remove_link: Optional[bool] = Field(default=None, description="If true, remove existing hyperlink from the text.")

    # Paragraph formatting (all optional)
    named_style: Optional[str] = Field(
        default=None,
        description="Paragraph style: 'HEADING_1' through 'HEADING_6', 'NORMAL_TEXT', 'TITLE', 'SUBTITLE'.",
    )
    alignment: Optional[str] = Field(
        default=None,
        description="Paragraph alignment: 'START', 'CENTER', 'END', 'JUSTIFIED'.",
    )
    line_spacing: Optional[float] = Field(
        default=None,
        description="Line spacing multiplier (1.0 = single, 1.5, 2.0 = double).",
    )
    indent_start: Optional[float] = Field(
        default=None,
        description="Left indent in points.",
        ge=0,
    )
```

**Step 2: Add the gdrive_format_text tool function**

```python
@mcp.tool(
    name="gdrive_format_text",
    annotations={
        "title": "Format Text in Google Doc",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def gdrive_format_text(params: FormatTextInput) -> str:
    """Apply formatting to existing text in a Google Doc.

    Three targeting modes:
    - 'match': Find text by string match, format specific or all occurrences.
    - 'range': Format text at specific character index range.
    - 'heading': Format all content under a heading (until next heading of same or higher level).

    Supports both text-level formatting (bold, italic, underline, font, colors, links)
    and paragraph-level formatting (heading styles, alignment, spacing, indentation).

    Args:
        params: File ID, target selection, and formatting options.

    Returns:
        Confirmation of formatting applied.
    """
    try:
        doc = get_docs().documents().get(documentId=params.file_id).execute()
    except Exception as e:
        return f"Error reading document: {e}"

    # Determine target ranges
    ranges = []

    if params.target == FormatTextTarget.MATCH:
        if not params.match_text:
            return "Error: match_text is required when target='match'."
        all_matches = find_text_indices(doc, params.match_text)
        if not all_matches:
            return f"Error: Text '{params.match_text}' not found in document."

        if params.match_occurrence is None:
            ranges = all_matches
        elif params.match_occurrence == -1:
            ranges = [all_matches[-1]]
        elif 1 <= params.match_occurrence <= len(all_matches):
            ranges = [all_matches[params.match_occurrence - 1]]
        else:
            return f"Error: Occurrence {params.match_occurrence} out of range (found {len(all_matches)} matches)."

    elif params.target == FormatTextTarget.RANGE:
        if params.start_index is None or params.end_index is None:
            return "Error: start_index and end_index are required when target='range'."
        ranges = [(params.start_index, params.end_index)]

    elif params.target == FormatTextTarget.HEADING:
        if not params.heading_text:
            return "Error: heading_text is required when target='heading'."
        section = find_heading_section_range(doc, params.heading_text)
        if section is None:
            return f"Error: Heading '{params.heading_text}' not found."
        ranges = [section]

    if not ranges:
        return "Error: No text ranges identified for formatting."

    # Build requests
    requests = []
    applied = []

    # Text style
    text_style = {}
    text_fields = []

    if params.bold is not None:
        text_style["bold"] = params.bold
        text_fields.append("bold")
    if params.italic is not None:
        text_style["italic"] = params.italic
        text_fields.append("italic")
    if params.underline is not None:
        text_style["underline"] = params.underline
        text_fields.append("underline")
    if params.strikethrough is not None:
        text_style["strikethrough"] = params.strikethrough
        text_fields.append("strikethrough")
    if params.font_family:
        text_style["weightedFontFamily"] = {"fontFamily": params.font_family}
        text_fields.append("weightedFontFamily")
    if params.font_size:
        text_style["fontSize"] = {"magnitude": params.font_size, "unit": "PT"}
        text_fields.append("fontSize")
    if params.text_color:
        text_style["foregroundColor"] = {"color": {"rgbColor": hex_to_color(params.text_color)}}
        text_fields.append("foregroundColor")
    if params.bg_color:
        text_style["backgroundColor"] = {"color": {"rgbColor": hex_to_color(params.bg_color)}}
        text_fields.append("backgroundColor")
    if params.link_url:
        text_style["link"] = {"url": params.link_url}
        text_fields.append("link")
    if params.remove_link:
        text_style["link"] = {}
        text_fields.append("link")

    if text_style:
        for start, end in ranges:
            requests.append({
                "updateTextStyle": {
                    "range": {"startIndex": start, "endIndex": end},
                    "textStyle": text_style,
                    "fields": ",".join(text_fields),
                }
            })
        applied.append("text formatting")

    # Paragraph style
    para_style = {}
    para_fields = []

    if params.named_style:
        para_style["namedStyleType"] = params.named_style
        para_fields.append("namedStyleType")
    if params.alignment:
        para_style["alignment"] = params.alignment
        para_fields.append("alignment")
    if params.line_spacing:
        para_style["lineSpacing"] = params.line_spacing * 100  # API uses percentage
        para_fields.append("lineSpacing")
    if params.indent_start is not None:
        para_style["indentStart"] = {"magnitude": params.indent_start, "unit": "PT"}
        para_fields.append("indentStart")

    if para_style:
        for start, end in ranges:
            requests.append({
                "updateParagraphStyle": {
                    "range": {"startIndex": start, "endIndex": end},
                    "paragraphStyle": para_style,
                    "fields": ",".join(para_fields),
                }
            })
        applied.append("paragraph formatting")

    if not requests:
        return "No formatting options specified."

    try:
        get_docs().documents().batchUpdate(
            documentId=params.file_id,
            body={"requests": requests},
        ).execute()

        target_desc = {
            FormatTextTarget.MATCH: f"'{params.match_text}' ({len(ranges)} occurrence{'s' if len(ranges) != 1 else ''})",
            FormatTextTarget.RANGE: f"index {params.start_index}-{params.end_index}",
            FormatTextTarget.HEADING: f"section under '{params.heading_text}'",
        }

        return f"Applied {', '.join(applied)} to {target_desc[params.target]}."

    except Exception as e:
        return f"Error applying formatting: {e}"
```

Note: Add `find_heading_section_range` to imports from helpers. This is a new helper added in Task 2.

**Step 3: Commit**

```bash
git add tools/docs.py
git commit -m "feat: add gdrive_format_text tool for rich text formatting

Three targeting modes (match, range, heading) with full text styling
(bold, italic, underline, strikethrough, font, colors, links) and
paragraph formatting (heading styles, alignment, spacing, indentation)."
```

---

## Phase 3: Gmail Integration

### Task 6: Add Gmail OAuth scope

**Files:**
- Modify: `auth.py`

**Step 1: Add Gmail readonly scope**

Add to the SCOPES list in auth.py:

```python
    "https://www.googleapis.com/auth/gmail.readonly",
```

Update the comment on line 25 to mention Gmail.

**Step 2: Commit**

```bash
git add auth.py
git commit -m "feat: add Gmail readonly OAuth scope"
```

### Task 7: Create tools/gmail.py with all 5 Gmail tools

**Files:**
- Create: `tools/gmail.py`
- Modify: `tools/__init__.py` (add gmail import and registration)

**Step 1: Create tools/gmail.py**

```python
"""Gmail tools for gdrive-mcp (read-only)."""

import base64
import email
from typing import Optional, List
from enum import Enum

from pydantic import BaseModel, Field, ConfigDict

from services import get_gmail


def register(mcp):
    """Register all Gmail tools with the MCP server."""

    # ── gmail_search ─────────────────────────────────────────────────────────

    class GmailSearchInput(BaseModel):
        """Input for searching Gmail messages."""
        model_config = ConfigDict(str_strip_whitespace=True)

        query: str = Field(
            ...,
            description=(
                "Gmail search query. Supports operators: from:, to:, subject:, is:unread, "
                "is:starred, label:, has:attachment, before:, after:, in:inbox, in:sent, "
                "newer_than:2d, older_than:1y, filename:pdf. "
                "Example: 'from:boss@company.com subject:Q4 after:2025/01/01'"
            ),
            min_length=1,
        )
        max_results: int = Field(
            default=20,
            description="Maximum number of results to return.",
            ge=1,
            le=100,
        )
        include_body: bool = Field(
            default=False,
            description="If true, include a body preview for each result. More tokens but saves a second call.",
        )
        body_length: int = Field(
            default=500,
            description="Maximum characters of body to include per message when include_body=True.",
            ge=100,
            le=5000,
        )

    @mcp.tool(
        name="gmail_search",
        annotations={"title": "Search Gmail", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gmail_search(params: GmailSearchInput) -> str:
        """Search Gmail messages with optional inline body previews.

        Uses full Gmail search syntax. Returns metadata and snippets for all results,
        with optional body previews to avoid needing a second call for triage.

        Args:
            params: Search query, result limit, and body preview options.

        Returns:
            Markdown-formatted search results with message metadata.
        """
        try:
            svc = get_gmail()
            result = svc.users().messages().list(
                userId="me",
                q=params.query,
                maxResults=params.max_results,
            ).execute()

            messages = result.get("messages", [])
            if not messages:
                return f"No messages found for query: `{params.query}`"

            output = f"## Gmail Search Results ({len(messages)} messages)\n\n"

            for msg_stub in messages:
                msg = svc.users().messages().get(
                    userId="me",
                    id=msg_stub["id"],
                    format="full" if params.include_body else "metadata",
                    metadataHeaders=["From", "To", "Subject", "Date"],
                ).execute()

                headers = {h["name"]: h["value"] for h in msg.get("payload", {}).get("headers", [])}
                labels = msg.get("labelIds", [])
                snippet = msg.get("snippet", "")

                output += f"**{headers.get('Subject', '(no subject)')}**\n"
                output += f"  From: {headers.get('From', '?')} | Date: {headers.get('Date', '?')}\n"
                output += f"  To: {headers.get('To', '?')}\n"
                output += f"  Labels: {', '.join(labels)}\n"
                output += f"  Message ID: `{msg['id']}` | Thread ID: `{msg['threadId']}`\n"
                if snippet:
                    output += f"  Snippet: {snippet}\n"

                if params.include_body:
                    body = _extract_body(msg.get("payload", {}))
                    if body:
                        preview = body[:params.body_length]
                        if len(body) > params.body_length:
                            preview += "..."
                        output += f"  Body preview:\n  > {preview}\n"

                output += "\n"

            if result.get("nextPageToken"):
                output += f"_More results available. Refine your query or increase max_results._\n"

            return output

        except Exception as e:
            return f"Error searching Gmail: {e}"

    # ── gmail_read ───────────────────────────────────────────────────────────

    class GmailReadFormat(str, Enum):
        FULL = "full"
        METADATA = "metadata"
        SUMMARY = "summary"

    class GmailReadInput(BaseModel):
        """Input for reading a single Gmail message."""
        model_config = ConfigDict(str_strip_whitespace=True)

        message_id: str = Field(..., description="Gmail message ID (from search results).", min_length=1)
        format: GmailReadFormat = Field(
            default=GmailReadFormat.FULL,
            description="'full' (complete body), 'metadata' (headers only), 'summary' (first N chars of body).",
        )
        max_body_length: Optional[int] = Field(
            default=None,
            description="Truncate body at this many characters. Useful for long emails/newsletters.",
            ge=100,
        )
        include_headers: bool = Field(default=True, description="Include From, To, Cc, Date, Subject headers.")
        include_attachments: bool = Field(
            default=False,
            description="List attachment filenames and sizes (does not download content).",
        )

    @mcp.tool(
        name="gmail_read",
        annotations={"title": "Read Gmail Message", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gmail_read(params: GmailReadInput) -> str:
        """Read a single Gmail message with granularity control.

        Supports full content, metadata-only, or summary modes. Body truncation
        prevents overwhelming context with large emails.

        Args:
            params: Message ID, format, and display options.

        Returns:
            Formatted email content.
        """
        try:
            svc = get_gmail()
            api_format = "metadata" if params.format == GmailReadFormat.METADATA else "full"
            msg = svc.users().messages().get(
                userId="me",
                id=params.message_id,
                format=api_format,
            ).execute()

            output = f"## Email\n\n"

            if params.include_headers:
                headers = {h["name"]: h["value"] for h in msg.get("payload", {}).get("headers", [])}
                output += f"**Subject:** {headers.get('Subject', '(no subject)')}\n"
                output += f"**From:** {headers.get('From', '?')}\n"
                output += f"**To:** {headers.get('To', '?')}\n"
                if headers.get("Cc"):
                    output += f"**Cc:** {headers['Cc']}\n"
                output += f"**Date:** {headers.get('Date', '?')}\n"
                output += f"**Message-ID:** `{msg['id']}` | **Thread:** `{msg['threadId']}`\n"
                labels = msg.get("labelIds", [])
                if labels:
                    output += f"**Labels:** {', '.join(labels)}\n"
                output += "\n---\n\n"

            if params.format != GmailReadFormat.METADATA:
                body = _extract_body(msg.get("payload", {}))
                if body:
                    max_len = params.max_body_length
                    if params.format == GmailReadFormat.SUMMARY and max_len is None:
                        max_len = 1000
                    if max_len and len(body) > max_len:
                        body = body[:max_len] + f"\n\n_... truncated ({len(body)} total chars)_"
                    output += body + "\n"
                else:
                    output += "_No text body found._\n"

            if params.include_attachments:
                attachments = _list_attachments(msg.get("payload", {}))
                if attachments:
                    output += "\n**Attachments:**\n"
                    for att in attachments:
                        output += f"- {att['filename']} ({att['mimeType']}, {att['size']} bytes)\n"

            return output

        except Exception as e:
            return f"Error reading message: {e}"

    # ── gmail_read_thread ────────────────────────────────────────────────────

    class GmailReadThreadInput(BaseModel):
        """Input for reading an entire email thread."""
        model_config = ConfigDict(str_strip_whitespace=True)

        thread_id: str = Field(..., description="Gmail thread ID.", min_length=1)
        max_messages: int = Field(default=10, description="Maximum messages to include from the thread.", ge=1, le=100)
        max_body_length: Optional[int] = Field(
            default=None,
            description="Truncate each message body at this many characters.",
            ge=100,
        )
        newest_first: bool = Field(default=True, description="If true, show newest messages first.")

    @mcp.tool(
        name="gmail_read_thread",
        annotations={"title": "Read Gmail Thread", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gmail_read_thread(params: GmailReadThreadInput) -> str:
        """Read an entire email conversation thread in one call.

        Returns all messages in the thread with optional body truncation.

        Args:
            params: Thread ID, message limit, and display options.

        Returns:
            Full conversation thread formatted as markdown.
        """
        try:
            svc = get_gmail()
            thread = svc.users().threads().get(
                userId="me",
                id=params.thread_id,
                format="full",
            ).execute()

            messages = thread.get("messages", [])
            if not messages:
                return "Thread is empty."

            if params.newest_first:
                messages = list(reversed(messages))

            messages = messages[:params.max_messages]

            # Get subject from first message
            first_headers = {h["name"]: h["value"] for h in messages[0].get("payload", {}).get("headers", [])}
            subject = first_headers.get("Subject", "(no subject)")

            output = f"## Thread: {subject} ({len(messages)} message{'s' if len(messages) != 1 else ''})\n\n"

            for i, msg in enumerate(messages, 1):
                headers = {h["name"]: h["value"] for h in msg.get("payload", {}).get("headers", [])}
                output += f"### Message {i}\n"
                output += f"**From:** {headers.get('From', '?')} | **Date:** {headers.get('Date', '?')}\n\n"

                body = _extract_body(msg.get("payload", {}))
                if body:
                    if params.max_body_length and len(body) > params.max_body_length:
                        body = body[:params.max_body_length] + f"\n_... truncated_"
                    output += body + "\n"
                else:
                    output += "_No text body._\n"

                output += "\n---\n\n"

            return output

        except Exception as e:
            return f"Error reading thread: {e}"

    # ── gmail_read_batch ─────────────────────────────────────────────────────

    class GmailReadBatchInput(BaseModel):
        """Input for reading multiple Gmail messages at once."""
        model_config = ConfigDict(str_strip_whitespace=True)

        message_ids: List[str] = Field(
            ...,
            description="List of Gmail message IDs to read (max 50).",
            min_length=1,
            max_length=50,
        )
        format: GmailReadFormat = Field(
            default=GmailReadFormat.FULL,
            description="'full', 'metadata', or 'summary'.",
        )
        max_body_length: Optional[int] = Field(
            default=None,
            description="Truncate each message body at this many characters.",
            ge=100,
        )

    @mcp.tool(
        name="gmail_read_batch",
        annotations={"title": "Read Gmail Messages (Batch)", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gmail_read_batch(params: GmailReadBatchInput) -> str:
        """Read multiple Gmail messages in a single call.

        More efficient than individual reads. Supports up to 50 messages
        with the same format and truncation controls as gmail_read.

        Args:
            params: List of message IDs, format, and truncation options.

        Returns:
            All requested messages formatted as markdown.
        """
        try:
            svc = get_gmail()
            api_format = "metadata" if params.format == GmailReadFormat.METADATA else "full"

            output = f"## Batch Read ({len(params.message_ids)} messages)\n\n"

            for mid in params.message_ids:
                try:
                    msg = svc.users().messages().get(
                        userId="me",
                        id=mid,
                        format=api_format,
                    ).execute()

                    headers = {h["name"]: h["value"] for h in msg.get("payload", {}).get("headers", [])}
                    output += f"### {headers.get('Subject', '(no subject)')}\n"
                    output += f"**From:** {headers.get('From', '?')} | **Date:** {headers.get('Date', '?')}\n"
                    output += f"**ID:** `{mid}` | **Thread:** `{msg['threadId']}`\n\n"

                    if params.format != GmailReadFormat.METADATA:
                        body = _extract_body(msg.get("payload", {}))
                        if body:
                            max_len = params.max_body_length
                            if params.format == GmailReadFormat.SUMMARY and max_len is None:
                                max_len = 1000
                            if max_len and len(body) > max_len:
                                body = body[:max_len] + "\n_... truncated_"
                            output += body + "\n"

                    output += "\n---\n\n"

                except Exception as e:
                    output += f"### Error reading `{mid}`\n{e}\n\n---\n\n"

            return output

        except Exception as e:
            return f"Error in batch read: {e}"

    # ── gmail_labels ─────────────────────────────────────────────────────────

    class GmailLabelsInput(BaseModel):
        """Input for listing Gmail labels."""
        model_config = ConfigDict(str_strip_whitespace=True)

    @mcp.tool(
        name="gmail_labels",
        annotations={"title": "List Gmail Labels", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gmail_labels(params: GmailLabelsInput) -> str:
        """List all Gmail labels with message and unread counts.

        Returns:
            All user labels with name, ID, and counts.
        """
        try:
            svc = get_gmail()
            result = svc.users().labels().list(userId="me").execute()
            labels = result.get("labels", [])

            if not labels:
                return "No labels found."

            # Get detailed info for each label
            user_labels = []
            system_labels = []

            for lbl in labels:
                detail = svc.users().labels().get(userId="me", id=lbl["id"]).execute()
                entry = {
                    "name": detail.get("name", "?"),
                    "id": detail.get("id", "?"),
                    "type": detail.get("type", "user"),
                    "total": detail.get("messagesTotal", 0),
                    "unread": detail.get("messagesUnread", 0),
                }
                if entry["type"] == "system":
                    system_labels.append(entry)
                else:
                    user_labels.append(entry)

            output = "## Gmail Labels\n\n"

            if system_labels:
                output += "### System Labels\n\n"
                output += "| Label | Total | Unread |\n|---|---|---|\n"
                for lbl in sorted(system_labels, key=lambda x: x["name"]):
                    output += f"| {lbl['name']} | {lbl['total']} | {lbl['unread']} |\n"
                output += "\n"

            if user_labels:
                output += "### User Labels\n\n"
                output += "| Label | ID | Total | Unread |\n|---|---|---|---|\n"
                for lbl in sorted(user_labels, key=lambda x: x["name"]):
                    output += f"| {lbl['name']} | `{lbl['id']}` | {lbl['total']} | {lbl['unread']} |\n"

            return output

        except Exception as e:
            return f"Error listing labels: {e}"

    # ── Gmail helper functions ───────────────────────────────────────────────

    def _extract_body(payload: dict) -> str:
        """Extract plain text body from Gmail message payload, handling MIME parts."""
        # Direct body (simple messages)
        if payload.get("body", {}).get("data"):
            return base64.urlsafe_b64decode(payload["body"]["data"]).decode("utf-8", errors="replace")

        # Multipart — find text/plain first, fall back to text/html
        parts = payload.get("parts", [])
        plain_body = None
        html_body = None

        for part in parts:
            mime = part.get("mimeType", "")
            if mime == "text/plain" and part.get("body", {}).get("data"):
                plain_body = base64.urlsafe_b64decode(part["body"]["data"]).decode("utf-8", errors="replace")
            elif mime == "text/html" and part.get("body", {}).get("data"):
                html_body = base64.urlsafe_b64decode(part["body"]["data"]).decode("utf-8", errors="replace")
            elif mime.startswith("multipart/"):
                # Recurse into nested multipart
                nested = _extract_body(part)
                if nested:
                    if plain_body is None:
                        plain_body = nested

        if plain_body:
            return plain_body
        if html_body:
            # Strip HTML tags for a rough plain text version
            import re
            text = re.sub(r'<br\s*/?>', '\n', html_body)
            text = re.sub(r'<[^>]+>', '', text)
            text = re.sub(r'&nbsp;', ' ', text)
            text = re.sub(r'&amp;', '&', text)
            text = re.sub(r'&lt;', '<', text)
            text = re.sub(r'&gt;', '>', text)
            text = re.sub(r'\n{3,}', '\n\n', text)
            return text.strip()

        return ""

    def _list_attachments(payload: dict) -> list:
        """List attachments in a message payload (name, type, size only)."""
        attachments = []
        parts = payload.get("parts", [])
        for part in parts:
            filename = part.get("filename")
            if filename:
                attachments.append({
                    "filename": filename,
                    "mimeType": part.get("mimeType", "unknown"),
                    "size": part.get("body", {}).get("size", 0),
                })
            if part.get("parts"):
                attachments.extend(_list_attachments(part))
        return attachments
```

**Step 2: Update tools/__init__.py**

Add `gmail` to the imports and `register_all`.

**Step 3: Smoke test**

Run: `cd /home/ageller/gdrive-mcp && timeout 5 python server.py 2>&1 || true`

Expected: No import errors.

**Step 4: Commit**

```bash
git add tools/gmail.py tools/__init__.py
git commit -m "feat: add Gmail tools (search, read, read_thread, read_batch, labels)

Read-only Gmail integration with:
- Smart search with optional inline body previews
- Granular message reading (full/metadata/summary with truncation)
- Thread-level conversation reading
- Batch message reading (up to 50)
- Label listing with message counts"
```

---

## Phase 4: Calendar Integration

### Task 8: Add Calendar OAuth scope

**Files:**
- Modify: `auth.py`

**Step 1: Add Calendar scope**

Add to SCOPES list:

```python
    "https://www.googleapis.com/auth/calendar",
```

**Step 2: Commit**

```bash
git add auth.py
git commit -m "feat: add Google Calendar OAuth scope"
```

### Task 9: Create tools/calendar.py with all 7 Calendar tools

**Files:**
- Create: `tools/calendar.py`
- Modify: `tools/__init__.py`

**Step 1: Create tools/calendar.py**

```python
"""Google Calendar tools for gdrive-mcp."""

from datetime import datetime, timezone
from typing import Optional, List

from pydantic import BaseModel, Field, ConfigDict

from services import get_calendar


def _parse_datetime(dt_str: str) -> str:
    """Parse a flexible datetime string to RFC3339.

    Accepts: RFC3339 ('2025-03-15T14:00:00Z'), date ('2025-03-15'),
    or date with simple time ('2025-03-15 2:30 PM' — treated as UTC).
    Returns RFC3339 string.
    """
    if not dt_str:
        return dt_str

    # Already RFC3339
    if "T" in dt_str:
        return dt_str

    # Just a date
    if len(dt_str) == 10:
        return dt_str

    # Try parsing "YYYY-MM-DD HH:MM" or "YYYY-MM-DD H:MM PM"
    for fmt in ["%Y-%m-%d %I:%M %p", "%Y-%m-%d %H:%M", "%Y-%m-%d %I:%M%p", "%Y-%m-%d %H:%M:%S"]:
        try:
            dt = datetime.strptime(dt_str.strip(), fmt)
            return dt.strftime("%Y-%m-%dT%H:%M:%S")
        except ValueError:
            continue

    return dt_str  # Return as-is; API will validate


def _build_recurrence(recurrence: str, count: Optional[int], until: Optional[str]) -> list:
    """Build RRULE list from simple string or raw RRULE."""
    if recurrence.startswith("RRULE:"):
        return [recurrence]

    freq_map = {"daily": "DAILY", "weekly": "WEEKLY", "monthly": "MONTHLY", "yearly": "YEARLY"}
    freq = freq_map.get(recurrence.lower())
    if not freq:
        return [recurrence]  # Assume raw RRULE without prefix

    rule = f"RRULE:FREQ={freq}"
    if count:
        rule += f";COUNT={count}"
    if until:
        until_clean = until.replace("-", "")
        if len(until_clean) == 8:
            until_clean += "T235959Z"
        rule += f";UNTIL={until_clean}"

    return [rule]


def _format_event(event: dict, detailed: bool = False) -> str:
    """Format a calendar event as markdown."""
    summary = event.get("summary", "(no title)")
    start = event.get("start", {})
    end = event.get("end", {})
    start_str = start.get("dateTime", start.get("date", "?"))
    end_str = end.get("dateTime", end.get("date", "?"))

    output = f"**{summary}**\n"
    output += f"  When: {start_str} → {end_str}\n"
    output += f"  Event ID: `{event.get('id', '?')}`\n"

    if event.get("location"):
        output += f"  Location: {event['location']}\n"
    if event.get("htmlLink"):
        output += f"  Link: {event['htmlLink']}\n"

    status = event.get("status", "")
    if status and status != "confirmed":
        output += f"  Status: {status}\n"

    if event.get("recurringEventId"):
        output += f"  Recurring (series ID: `{event['recurringEventId']}`)\n"

    organizer = event.get("organizer", {})
    if organizer:
        output += f"  Organizer: {organizer.get('displayName', organizer.get('email', '?'))}\n"

    if detailed:
        if event.get("description"):
            output += f"  Description: {event['description']}\n"
        attendees = event.get("attendees", [])
        if attendees:
            output += f"  Attendees ({len(attendees)}):\n"
            for a in attendees:
                name = a.get("displayName", a.get("email", "?"))
                resp = a.get("responseStatus", "?")
                opt = " (optional)" if a.get("optional") else ""
                output += f"    - {name} [{resp}]{opt}\n"

        hangout = event.get("hangoutLink") or event.get("conferenceData", {}).get("entryPoints", [{}])[0].get("uri")
        if hangout:
            output += f"  Meet: {hangout}\n"

        if event.get("recurrence"):
            for rule in event["recurrence"]:
                output += f"  Recurrence: {rule}\n"

        reminders = event.get("reminders", {})
        if reminders.get("overrides"):
            for r in reminders["overrides"]:
                output += f"  Reminder: {r['method']} {r['minutes']}min before\n"

        if event.get("attachments"):
            output += "  Attachments:\n"
            for att in event["attachments"]:
                output += f"    - {att.get('title', 'untitled')} ({att.get('mimeType', '?')})\n"

    return output


def register(mcp):
    """Register all Calendar tools with the MCP server."""

    # ── gcal_list_events ─────────────────────────────────────────────────────

    class CalListEventsInput(BaseModel):
        """Input for listing calendar events."""
        model_config = ConfigDict(str_strip_whitespace=True)

        calendar_id: str = Field(default="primary", description="Calendar ID. Use 'primary' for your main calendar.")
        time_min: Optional[str] = Field(default=None, description="Start of time range (RFC3339, date, or 'YYYY-MM-DD HH:MM'). Defaults to now.")
        time_max: Optional[str] = Field(default=None, description="End of time range (same formats).")
        query: Optional[str] = Field(default=None, description="Keyword search across event title, description, and location.")
        max_results: int = Field(default=25, description="Maximum events to return.", ge=1, le=250)
        include_recurring: bool = Field(default=True, description="If true, expand recurring events into individual instances.")
        detailed: bool = Field(default=False, description="If true, include attendees, description, attachments, and Meet links.")
        status_filter: Optional[str] = Field(
            default=None,
            description="Filter by YOUR response status: 'accepted', 'tentative', 'declined', 'needsAction'.",
        )

    @mcp.tool(
        name="gcal_list_events",
        annotations={"title": "List Calendar Events", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gcal_list_events(params: CalListEventsInput) -> str:
        """List events from a Google Calendar with flexible filtering.

        Supports time range, keyword search, recurring event expansion,
        and response status filtering.

        Args:
            params: Calendar ID, time range, search query, and display options.

        Returns:
            Markdown-formatted event listing.
        """
        try:
            svc = get_calendar()
            kwargs = {
                "calendarId": params.calendar_id,
                "maxResults": params.max_results,
                "singleEvents": params.include_recurring,
                "orderBy": "startTime" if params.include_recurring else "updated",
            }

            if params.time_min:
                parsed = _parse_datetime(params.time_min)
                if "T" not in parsed:
                    parsed += "T00:00:00Z"
                if not parsed.endswith("Z") and "+" not in parsed and "-" not in parsed[10:]:
                    parsed += "Z"
                kwargs["timeMin"] = parsed
            else:
                kwargs["timeMin"] = datetime.now(timezone.utc).isoformat()

            if params.time_max:
                parsed = _parse_datetime(params.time_max)
                if "T" not in parsed:
                    parsed += "T23:59:59Z"
                if not parsed.endswith("Z") and "+" not in parsed and "-" not in parsed[10:]:
                    parsed += "Z"
                kwargs["timeMax"] = parsed

            if params.query:
                kwargs["q"] = params.query

            result = svc.events().list(**kwargs).execute()
            events = result.get("items", [])

            if params.status_filter:
                filtered = []
                for ev in events:
                    for att in ev.get("attendees", []):
                        if att.get("self") and att.get("responseStatus") == params.status_filter:
                            filtered.append(ev)
                            break
                events = filtered

            if not events:
                return "No events found."

            output = f"## Calendar Events ({len(events)})\n\n"
            for ev in events:
                output += _format_event(ev, detailed=params.detailed) + "\n"

            return output

        except Exception as e:
            return f"Error listing events: {e}"

    # ── gcal_get_event ───────────────────────────────────────────────────────

    class CalGetEventInput(BaseModel):
        """Input for getting a single calendar event."""
        model_config = ConfigDict(str_strip_whitespace=True)

        event_id: str = Field(..., description="Event ID.", min_length=1)
        calendar_id: str = Field(default="primary", description="Calendar ID.")

    @mcp.tool(
        name="gcal_get_event",
        annotations={"title": "Get Calendar Event", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gcal_get_event(params: CalGetEventInput) -> str:
        """Get full details for a single calendar event.

        Args:
            params: Event ID and calendar ID.

        Returns:
            Detailed event information.
        """
        try:
            event = get_calendar().events().get(
                calendarId=params.calendar_id,
                eventId=params.event_id,
            ).execute()
            return _format_event(event, detailed=True)
        except Exception as e:
            return f"Error getting event: {e}"

    # ── gcal_create_event ────────────────────────────────────────────────────

    class CalCreateEventInput(BaseModel):
        """Input for creating a calendar event."""
        model_config = ConfigDict(str_strip_whitespace=True)

        summary: str = Field(..., description="Event title.", min_length=1)
        start: str = Field(..., description="Start time (RFC3339, date 'YYYY-MM-DD', or 'YYYY-MM-DD HH:MM').")
        end: str = Field(..., description="End time (same formats as start).")
        calendar_id: str = Field(default="primary", description="Calendar ID.")
        timezone: Optional[str] = Field(default=None, description="IANA timezone (e.g., 'America/New_York'). Uses calendar default if omitted.")
        description: Optional[str] = Field(default=None, description="Event description.")
        location: Optional[str] = Field(default=None, description="Event location.")
        attendees: Optional[List[str]] = Field(default=None, description="Attendee email addresses.")
        recurrence: Optional[str] = Field(
            default=None,
            description="Recurrence: 'daily', 'weekly', 'monthly', 'yearly', or raw RRULE string.",
        )
        recurrence_count: Optional[int] = Field(default=None, description="Number of times to repeat.", ge=1)
        recurrence_until: Optional[str] = Field(default=None, description="Repeat until this date (YYYY-MM-DD).")
        reminders: Optional[List[dict]] = Field(
            default=None,
            description="Reminder list: [{'method': 'popup', 'minutes': 15}]. Method: 'popup' or 'email'.",
        )
        add_google_meet: bool = Field(default=False, description="If true, add a Google Meet video conference.")
        visibility: Optional[str] = Field(default=None, description="'default', 'public', or 'private'.")
        transparency: Optional[str] = Field(default=None, description="'opaque' (busy) or 'transparent' (free).")
        guests_can_modify: Optional[bool] = Field(default=None, description="Allow guests to modify the event.")
        send_notifications: str = Field(default="all", description="'all', 'none', or 'externalOnly'.")

    @mcp.tool(
        name="gcal_create_event",
        annotations={"title": "Create Calendar Event", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
    )
    async def gcal_create_event(params: CalCreateEventInput) -> str:
        """Create a new Google Calendar event with full options.

        Supports recurrence (simple strings like 'weekly' auto-convert to RRULE),
        Google Meet, attendees (just email strings), reminders, and visibility settings.

        Args:
            params: Event details.

        Returns:
            Confirmation with event ID and link.
        """
        try:
            start_parsed = _parse_datetime(params.start)
            end_parsed = _parse_datetime(params.end)

            is_all_day = len(start_parsed) == 10

            event_body = {"summary": params.summary}

            if is_all_day:
                event_body["start"] = {"date": start_parsed}
                event_body["end"] = {"date": end_parsed}
            else:
                start_obj = {"dateTime": start_parsed}
                end_obj = {"dateTime": end_parsed}
                if params.timezone:
                    start_obj["timeZone"] = params.timezone
                    end_obj["timeZone"] = params.timezone
                event_body["start"] = start_obj
                event_body["end"] = end_obj

            if params.description:
                event_body["description"] = params.description
            if params.location:
                event_body["location"] = params.location
            if params.attendees:
                event_body["attendees"] = [{"email": e} for e in params.attendees]
            if params.recurrence:
                event_body["recurrence"] = _build_recurrence(
                    params.recurrence, params.recurrence_count, params.recurrence_until
                )
            if params.reminders:
                event_body["reminders"] = {"useDefault": False, "overrides": params.reminders}
            if params.visibility:
                event_body["visibility"] = params.visibility
            if params.transparency:
                event_body["transparency"] = params.transparency
            if params.guests_can_modify is not None:
                event_body["guestsCanModify"] = params.guests_can_modify

            kwargs = {
                "calendarId": params.calendar_id,
                "body": event_body,
                "sendUpdates": params.send_notifications,
            }

            if params.add_google_meet:
                event_body["conferenceData"] = {
                    "createRequest": {
                        "conferenceSolutionKey": {"type": "hangoutsMeet"},
                        "requestId": f"gdrive-mcp-{datetime.now().strftime('%Y%m%d%H%M%S')}",
                    }
                }
                kwargs["conferenceDataVersion"] = 1

            event = get_calendar().events().insert(**kwargs).execute()

            output = f"Event created: **{event.get('summary')}**\n\n"
            output += f"Event ID: `{event['id']}`\n"
            output += f"Link: {event.get('htmlLink', 'N/A')}\n"

            hangout = event.get("hangoutLink")
            if hangout:
                output += f"Google Meet: {hangout}\n"

            return output

        except Exception as e:
            return f"Error creating event: {e}"

    # ── gcal_update_event ────────────────────────────────────────────────────

    class CalUpdateEventInput(BaseModel):
        """Input for updating a calendar event."""
        model_config = ConfigDict(str_strip_whitespace=True)

        event_id: str = Field(..., description="Event ID to update.", min_length=1)
        calendar_id: str = Field(default="primary", description="Calendar ID.")
        summary: Optional[str] = Field(default=None, description="New title.")
        start: Optional[str] = Field(default=None, description="New start time.")
        end: Optional[str] = Field(default=None, description="New end time.")
        timezone: Optional[str] = Field(default=None, description="IANA timezone.")
        description: Optional[str] = Field(default=None, description="New description.")
        location: Optional[str] = Field(default=None, description="New location.")
        add_attendees: Optional[List[str]] = Field(default=None, description="Email addresses to ADD to the event.")
        remove_attendees: Optional[List[str]] = Field(default=None, description="Email addresses to REMOVE from the event.")
        recurrence: Optional[str] = Field(default=None, description="New recurrence rule.")
        reminders: Optional[List[dict]] = Field(default=None, description="New reminders.")
        add_google_meet: Optional[bool] = Field(default=None, description="Add (true) or keep (null) Google Meet.")
        visibility: Optional[str] = Field(default=None, description="'default', 'public', or 'private'.")
        transparency: Optional[str] = Field(default=None, description="'opaque' or 'transparent'.")
        send_notifications: str = Field(default="all", description="'all', 'none', or 'externalOnly'.")

    @mcp.tool(
        name="gcal_update_event",
        annotations={"title": "Update Calendar Event", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gcal_update_event(params: CalUpdateEventInput) -> str:
        """Update an existing calendar event with partial changes.

        Use add_attendees/remove_attendees to modify the guest list without
        having to re-submit the entire list. Only include fields you want to change.

        Args:
            params: Event ID and fields to update.

        Returns:
            Confirmation with updated event details.
        """
        try:
            svc = get_calendar()

            # Fetch existing event
            event = svc.events().get(
                calendarId=params.calendar_id,
                eventId=params.event_id,
            ).execute()

            # Apply changes
            updated = []

            if params.summary is not None:
                event["summary"] = params.summary
                updated.append("title")
            if params.description is not None:
                event["description"] = params.description
                updated.append("description")
            if params.location is not None:
                event["location"] = params.location
                updated.append("location")
            if params.visibility:
                event["visibility"] = params.visibility
                updated.append("visibility")
            if params.transparency:
                event["transparency"] = params.transparency
                updated.append("transparency")

            if params.start or params.end:
                if params.start:
                    parsed = _parse_datetime(params.start)
                    is_all_day = len(parsed) == 10
                    if is_all_day:
                        event["start"] = {"date": parsed}
                    else:
                        obj = {"dateTime": parsed}
                        if params.timezone:
                            obj["timeZone"] = params.timezone
                        event["start"] = obj
                if params.end:
                    parsed = _parse_datetime(params.end)
                    is_all_day = len(parsed) == 10
                    if is_all_day:
                        event["end"] = {"date": parsed}
                    else:
                        obj = {"dateTime": parsed}
                        if params.timezone:
                            obj["timeZone"] = params.timezone
                        event["end"] = obj
                updated.append("time")

            # Attendee management — add/remove without full list
            if params.add_attendees or params.remove_attendees:
                current = event.get("attendees", [])
                current_emails = {a["email"] for a in current}

                if params.remove_attendees:
                    remove_set = set(params.remove_attendees)
                    current = [a for a in current if a["email"] not in remove_set]

                if params.add_attendees:
                    for email_addr in params.add_attendees:
                        if email_addr not in current_emails:
                            current.append({"email": email_addr})

                event["attendees"] = current
                updated.append("attendees")

            if params.recurrence:
                event["recurrence"] = _build_recurrence(params.recurrence, None, None)
                updated.append("recurrence")

            if params.reminders:
                event["reminders"] = {"useDefault": False, "overrides": params.reminders}
                updated.append("reminders")

            kwargs = {
                "calendarId": params.calendar_id,
                "eventId": params.event_id,
                "body": event,
                "sendUpdates": params.send_notifications,
            }

            if params.add_google_meet:
                event["conferenceData"] = {
                    "createRequest": {
                        "conferenceSolutionKey": {"type": "hangoutsMeet"},
                        "requestId": f"gdrive-mcp-update-{datetime.now().strftime('%Y%m%d%H%M%S')}",
                    }
                }
                kwargs["conferenceDataVersion"] = 1
                updated.append("Google Meet")

            result = svc.events().update(**kwargs).execute()

            return f"Updated event **{result.get('summary')}**: {', '.join(updated)}.\nLink: {result.get('htmlLink', 'N/A')}"

        except Exception as e:
            return f"Error updating event: {e}"

    # ── gcal_delete_event ────────────────────────────────────────────────────

    class CalDeleteEventInput(BaseModel):
        """Input for deleting a calendar event."""
        model_config = ConfigDict(str_strip_whitespace=True)

        event_id: str = Field(..., description="Event ID to delete.", min_length=1)
        calendar_id: str = Field(default="primary", description="Calendar ID.")
        send_notifications: str = Field(default="all", description="'all', 'none', or 'externalOnly'.")

    @mcp.tool(
        name="gcal_delete_event",
        annotations={"title": "Delete Calendar Event", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": True, "openWorldHint": True},
    )
    async def gcal_delete_event(params: CalDeleteEventInput) -> str:
        """Delete a calendar event.

        Args:
            params: Event ID and notification preference.

        Returns:
            Confirmation of deletion.
        """
        try:
            get_calendar().events().delete(
                calendarId=params.calendar_id,
                eventId=params.event_id,
                sendUpdates=params.send_notifications,
            ).execute()
            return f"Deleted event `{params.event_id}`."
        except Exception as e:
            return f"Error deleting event: {e}"

    # ── gcal_free_busy ───────────────────────────────────────────────────────

    class CalFreeBusyInput(BaseModel):
        """Input for checking free/busy status."""
        model_config = ConfigDict(str_strip_whitespace=True)

        calendars: List[str] = Field(default=["primary"], description="Calendar IDs to check.")
        time_min: str = Field(..., description="Start of time range.")
        time_max: str = Field(..., description="End of time range.")
        timezone: Optional[str] = Field(default=None, description="IANA timezone.")

    @mcp.tool(
        name="gcal_free_busy",
        annotations={"title": "Check Free/Busy", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gcal_free_busy(params: CalFreeBusyInput) -> str:
        """Check free/busy status for one or more calendars.

        Args:
            params: Calendar IDs and time range.

        Returns:
            Busy time blocks for each calendar.
        """
        try:
            time_min = _parse_datetime(params.time_min)
            time_max = _parse_datetime(params.time_max)

            if "T" not in time_min:
                time_min += "T00:00:00Z"
            if "T" not in time_max:
                time_max += "T23:59:59Z"

            for s in [time_min, time_max]:
                if not s.endswith("Z") and "+" not in s and "-" not in s[10:]:
                    s += "Z"

            body = {
                "timeMin": time_min,
                "timeMax": time_max,
                "items": [{"id": c} for c in params.calendars],
            }
            if params.timezone:
                body["timeZone"] = params.timezone

            result = get_calendar().freebusy().query(body=body).execute()

            output = "## Free/Busy\n\n"
            for cal_id, cal_data in result.get("calendars", {}).items():
                busy = cal_data.get("busy", [])
                output += f"**{cal_id}**\n"
                if not busy:
                    output += "  Free during this period.\n"
                else:
                    for block in busy:
                        output += f"  Busy: {block['start']} → {block['end']}\n"
                output += "\n"

            return output

        except Exception as e:
            return f"Error checking free/busy: {e}"

    # ── gcal_list_calendars ──────────────────────────────────────────────────

    class CalListCalendarsInput(BaseModel):
        """Input for listing calendars."""
        model_config = ConfigDict(str_strip_whitespace=True)

    @mcp.tool(
        name="gcal_list_calendars",
        annotations={"title": "List Calendars", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gcal_list_calendars(params: CalListCalendarsInput) -> str:
        """List all available Google Calendars.

        Returns:
            All calendars with ID, name, access role, and timezone.
        """
        try:
            result = get_calendar().calendarList().list().execute()
            calendars = result.get("items", [])

            if not calendars:
                return "No calendars found."

            output = "## Your Calendars\n\n"
            output += "| Calendar | ID | Role | Timezone |\n|---|---|---|---|\n"
            for cal in calendars:
                name = cal.get("summary", "?")
                primary = " (primary)" if cal.get("primary") else ""
                cal_id = cal.get("id", "?")
                role = cal.get("accessRole", "?")
                tz = cal.get("timeZone", "?")
                output += f"| {name}{primary} | `{cal_id}` | {role} | {tz} |\n"

            return output

        except Exception as e:
            return f"Error listing calendars: {e}"
```

**Step 2: Update tools/__init__.py to include calendar**

**Step 3: Smoke test**

**Step 4: Commit**

```bash
git add tools/calendar.py tools/__init__.py
git commit -m "feat: add Calendar tools (list, get, create, update, delete, free_busy, list_calendars)

Full Google Calendar integration with:
- Flexible event listing with status filtering and keyword search
- Single event detail view
- Event creation with recurrence, Google Meet, attendees, reminders
- Partial event updates with add/remove attendees (no full-list re-submission)
- Event deletion
- Free/busy availability checking
- Calendar listing"
```

---

## Phase 5: Enhanced Apps Script

### Task 10: Add Apps Script deployment scope

**Files:**
- Modify: `auth.py`

**Step 1: Add deployment scope**

Add to SCOPES:

```python
    "https://www.googleapis.com/auth/script.deployments",
```

**Step 2: Commit**

```bash
git add auth.py
git commit -m "feat: add Apps Script deployments OAuth scope"
```

### Task 11: Add 4 new Apps Script tools

**Files:**
- Modify: `tools/scripts.py`

**Step 1: Add gdrive_list_scripts**

```python
    class ListScriptsInput(BaseModel):
        """Input for listing Apps Script projects."""
        model_config = ConfigDict(str_strip_whitespace=True)

        max_results: int = Field(default=20, description="Maximum projects to return.", ge=1, le=100)
        bound_to: Optional[str] = Field(
            default=None,
            description="File ID to filter scripts bound to a specific doc/sheet/slides.",
        )
        query: Optional[str] = Field(default=None, description="Search script projects by name.")

    @mcp.tool(
        name="gdrive_list_scripts",
        annotations={"title": "List Apps Script Projects", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gdrive_list_scripts(params: ListScriptsInput) -> str:
        """List Apps Script projects, optionally filtered by bound document or name.

        Args:
            params: Filters and result limit.

        Returns:
            List of script projects with IDs and links.
        """
        try:
            q = "mimeType='application/vnd.google-apps.script'"
            if params.query:
                q += f" and name contains '{params.query}'"
            if params.bound_to:
                q += f" and '{params.bound_to}' in parents"

            files = drive_query_files(q, max_results=params.max_results)

            if not files:
                return "No Apps Script projects found."

            output = f"## Apps Script Projects ({len(files)})\n\n"
            for f in files:
                name = f.get("name", "untitled")
                fid = f.get("id", "?")
                modified = f.get("modifiedTime", "")[:10]
                link = f"https://script.google.com/d/{fid}/edit"
                output += f"**{name}**\n  ID: `{fid}` | Modified: {modified}\n  Editor: {link}\n\n"

            return output

        except Exception as e:
            return f"Error listing scripts: {e}"
```

Add import: `from helpers import drive_query_files`

**Step 2: Add gdrive_get_script**

```python
    class GetScriptInput(BaseModel):
        """Input for reading an Apps Script project."""
        model_config = ConfigDict(str_strip_whitespace=True)

        script_id: str = Field(..., description="Apps Script project ID.", min_length=1)
        file_name: Optional[str] = Field(
            default=None,
            description="Specific file to read (e.g., 'Code.gs'). If omitted, returns all files.",
        )
        include_manifest: bool = Field(default=True, description="Include appsscript.json manifest.")

    @mcp.tool(
        name="gdrive_get_script",
        annotations={"title": "Read Apps Script Project", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gdrive_get_script(params: GetScriptInput) -> str:
        """Read the contents of an Apps Script project.

        Can read a specific file or all files. Includes the manifest by default.

        Args:
            params: Script ID and optional file filter.

        Returns:
            Script file contents.
        """
        try:
            content = get_scripts().projects().getContent(scriptId=params.script_id).execute()
            files = content.get("files", [])

            if params.file_name:
                files = [f for f in files if f.get("name") == params.file_name.replace(".gs", "").replace(".html", "").replace(".json", "")]
                if not files:
                    return f"File '{params.file_name}' not found in project."

            if not params.include_manifest:
                files = [f for f in files if f.get("name") != "appsscript"]

            if not files:
                return "No files in project."

            output = f"## Apps Script Project `{params.script_id}`\n\n"
            for f in files:
                name = f.get("name", "?")
                ftype = f.get("type", "?")
                ext = {"SERVER_JS": ".gs", "HTML": ".html", "JSON": ".json"}.get(ftype, "")
                source = f.get("source", "")
                output += f"### {name}{ext} ({ftype})\n\n```{'javascript' if ftype == 'SERVER_JS' else 'html' if ftype == 'HTML' else 'json'}\n{source}\n```\n\n"

            return output

        except Exception as e:
            return f"Error reading script: {e}"
```

**Step 3: Add gdrive_update_script**

```python
    class ScriptFileInput(BaseModel):
        name: str = Field(..., description="File name (e.g., 'Code.gs', 'utils.gs', 'sidebar.html').")
        source: str = Field(..., description="File source code.")
        type: Optional[str] = Field(
            default=None,
            description="File type: 'SERVER_JS', 'HTML', or 'JSON'. Auto-detected from extension if omitted.",
        )

    class UpdateScriptInput(BaseModel):
        """Input for updating an Apps Script project."""
        model_config = ConfigDict(str_strip_whitespace=True)

        script_id: str = Field(..., description="Apps Script project ID.", min_length=1)
        files: List[ScriptFileInput] = Field(..., description="Files to create or update.", min_length=1)
        merge: bool = Field(
            default=True,
            description="If true (default), only update specified files — keep existing files untouched. If false, replace ALL files.",
        )

    @mcp.tool(
        name="gdrive_update_script",
        annotations={"title": "Update Apps Script Project", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gdrive_update_script(params: UpdateScriptInput) -> str:
        """Update files in an Apps Script project.

        Default merge mode only updates the files you specify — existing files
        are preserved. Set merge=False to replace all files.

        Args:
            params: Script ID, files to update, and merge mode.

        Returns:
            Confirmation of files updated.
        """
        try:
            svc = get_scripts()

            # Auto-detect file types from extensions
            type_map = {".gs": "SERVER_JS", ".js": "SERVER_JS", ".html": "HTML", ".json": "JSON"}

            new_files = []
            for f in params.files:
                ftype = f.type
                if not ftype:
                    for ext, t in type_map.items():
                        if f.name.endswith(ext):
                            ftype = t
                            break
                    if not ftype:
                        ftype = "SERVER_JS"

                # Strip extension for API (it uses bare names)
                name = f.name
                for ext in type_map:
                    if name.endswith(ext):
                        name = name[:-len(ext)]
                        break

                new_files.append({"name": name, "type": ftype, "source": f.source})

            if params.merge:
                # Fetch existing files and merge
                existing = svc.projects().getContent(scriptId=params.script_id).execute()
                existing_files = existing.get("files", [])

                new_names = {f["name"] for f in new_files}
                merged = [f for f in existing_files if f["name"] not in new_names]
                merged.extend(new_files)
                final_files = merged
            else:
                final_files = new_files

            svc.projects().updateContent(
                scriptId=params.script_id,
                body={"files": final_files},
            ).execute()

            names = [f["name"] for f in new_files]
            mode = "merged into" if params.merge else "replaced all files in"
            return f"Updated {len(new_files)} file(s) ({', '.join(names)}) — {mode} project `{params.script_id}`."

        except Exception as e:
            return f"Error updating script: {e}"
```

**Step 4: Add gdrive_deploy_script**

```python
    class DeployScriptInput(BaseModel):
        """Input for managing Apps Script deployments."""
        model_config = ConfigDict(str_strip_whitespace=True)

        script_id: str = Field(..., description="Apps Script project ID.", min_length=1)
        action: str = Field(
            ...,
            description="Action: 'create', 'list', 'update', or 'delete'.",
        )
        deployment_id: Optional[str] = Field(default=None, description="Deployment ID (required for 'update' and 'delete').")
        description: Optional[str] = Field(default=None, description="Deployment description.")
        version: Optional[int] = Field(default=None, description="Script version number. If omitted, uses latest.")

    @mcp.tool(
        name="gdrive_deploy_script",
        annotations={"title": "Deploy Apps Script", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
    )
    async def gdrive_deploy_script(params: DeployScriptInput) -> str:
        """Create, list, update, or delete Apps Script deployments.

        Args:
            params: Script ID, action, and deployment details.

        Returns:
            Deployment details or listing.
        """
        try:
            svc = get_scripts()

            if params.action == "list":
                result = svc.projects().deployments().list(scriptId=params.script_id).execute()
                deployments = result.get("deployments", [])
                if not deployments:
                    return "No deployments found."

                output = f"## Deployments ({len(deployments)})\n\n"
                for d in deployments:
                    did = d.get("deploymentId", "?")
                    config = d.get("deploymentConfig", {})
                    output += f"**{config.get('description', 'No description')}**\n"
                    output += f"  ID: `{did}`\n"
                    output += f"  Version: {config.get('versionNumber', 'HEAD')}\n"
                    output += f"  Type: {config.get('scriptId', 'API_EXECUTABLE')}\n\n"
                return output

            elif params.action == "create":
                # Create a version first if needed
                if params.version is None:
                    ver = svc.projects().versions().create(
                        scriptId=params.script_id,
                        body={"description": params.description or "Deployed via gdrive-mcp"},
                    ).execute()
                    version_number = ver.get("versionNumber")
                else:
                    version_number = params.version

                config = {"versionNumber": version_number}
                if params.description:
                    config["description"] = params.description

                deployment = svc.projects().deployments().create(
                    scriptId=params.script_id,
                    body={"deploymentConfig": config},
                ).execute()

                did = deployment.get("deploymentId", "?")
                return f"Created deployment `{did}` (version {version_number})."

            elif params.action == "update":
                if not params.deployment_id:
                    return "Error: deployment_id is required for 'update'."

                config = {}
                if params.version:
                    config["versionNumber"] = params.version
                if params.description:
                    config["description"] = params.description

                svc.projects().deployments().update(
                    scriptId=params.script_id,
                    deploymentId=params.deployment_id,
                    body={"deploymentConfig": config},
                ).execute()

                return f"Updated deployment `{params.deployment_id}`."

            elif params.action == "delete":
                if not params.deployment_id:
                    return "Error: deployment_id is required for 'delete'."

                svc.projects().deployments().delete(
                    scriptId=params.script_id,
                    deploymentId=params.deployment_id,
                ).execute()

                return f"Deleted deployment `{params.deployment_id}`."

            else:
                return "Error: action must be 'create', 'list', 'update', or 'delete'."

        except Exception as e:
            return f"Error managing deployment: {e}"
```

**Step 5: Smoke test**

**Step 6: Commit**

```bash
git add tools/scripts.py
git commit -m "feat: add Apps Script management tools (list, get, update, deploy)

Complete Apps Script lifecycle:
- List projects with filtering by bound document and name search
- Read project contents (single file or all files, with manifest)
- Update files with merge mode (preserves existing files by default)
- Full deployment management (create, list, update, delete)"
```

---

## Phase 6: Final Verification

### Task 12: Re-auth and full smoke test

**Step 1: Delete old token to force re-auth with new scopes**

```bash
rm /home/ageller/gdrive-mcp/token.json
```

**Step 2: Re-authenticate**

```bash
cd /home/ageller/gdrive-mcp && python auth.py
```

This opens the browser for OAuth consent with the 3 new scopes (Gmail readonly, Calendar, Script deployments).

**Step 3: Verify server starts cleanly**

```bash
cd /home/ageller/gdrive-mcp && timeout 5 python server.py 2>&1 || true
```

Expected: No import errors, server starts and blocks on stdin.

**Step 4: Verify tool count**

Add a quick check script:

```bash
cd /home/ageller/gdrive-mcp && python -c "
from mcp.server.fastmcp import FastMCP
from tools import register_all
mcp = FastMCP('test')
register_all(mcp)
tools = mcp.list_tools()
print(f'Total tools registered: {len(tools)}')
for t in sorted(tools, key=lambda x: x.name):
    print(f'  - {t.name}')
"
```

Expected: 46 tools listed (27 original + 1 format_text + 5 gmail + 7 calendar + 4 scripts + 2 modified-in-place).

**Step 5: Final commit**

```bash
git add -A
git commit -m "chore: verify all 46 tools register and server starts cleanly"
```

---

## Appendix: File Map Reference

| Module | Tools | Line Budget (est.) |
|---|---|---|
| `services.py` | — | ~60 |
| `helpers.py` | — | ~180 |
| `tools/drive.py` | gdrive_search, gdrive_read_doc, gdrive_read_sheet, gdrive_list_folder, gdrive_recent, gdrive_file_info | ~760 |
| `tools/docs.py` | gdrive_find_replace, gdrive_insert_text, gdrive_delete_text, gdrive_create_doc, gdrive_read_section, gdrive_insert_table, **gdrive_format_text** | ~900 |
| `tools/sheets.py` | gdrive_write_sheet, gdrive_append_sheet, gdrive_format_cells, gdrive_manage_sheets, gdrive_create_sheet, gdrive_insert_rows_cols, gdrive_delete_rows_cols | ~820 |
| `tools/slides.py` | gdrive_read_slides | ~110 |
| `tools/comments.py` | gdrive_comments | ~90 |
| `tools/management.py` | gdrive_move_copy, gdrive_share, gdrive_export, gdrive_versions | ~360 |
| `tools/gmail.py` | gmail_search, gmail_read, gmail_read_thread, gmail_read_batch, gmail_labels | ~450 |
| `tools/calendar.py` | gcal_list_events, gcal_get_event, gcal_create_event, gcal_update_event, gcal_delete_event, gcal_free_busy, gcal_list_calendars | ~550 |
| `tools/scripts.py` | gdrive_run_script, gdrive_create_script, **gdrive_list_scripts, gdrive_get_script, gdrive_update_script, gdrive_deploy_script** | ~400 |
| `server.py` | — (coordinator) | ~15 |
| **Total** | **46 tools** | **~4700** |
