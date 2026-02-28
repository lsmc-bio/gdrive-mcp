"""
Google Drive MCP Server for Claude

A local MCP server that gives Claude full read-write access to Google Drive,
Docs, Sheets, Slides, and Apps Script:

READ:
- Smart search across files (by name, content, type, date, owner)
- Read Google Docs as clean markdown (full doc or specific section by heading)
- Read Google Sheets as markdown tables (with range/tab selection)
- Read Google Slides as markdown (text, tables, speaker notes)
- Browse folder structures and file metadata
- File version history

WRITE - Docs:
- Find-and-replace text, insert formatted text, delete text
- Insert tables with cell values
- Create new Google Docs from markdown

WRITE - Sheets:
- Write values, append rows, insert/delete rows and columns
- Format cells (bold, colors, borders, number formats, alignment, merge)
- Create spreadsheets, manage tabs (create/duplicate/delete/rename)

DRIVE:
- Move/copy files between folders
- Share files (email or link), manage permissions
- Export to PDF/DOCX/XLSX/CSV/PPTX/TXT/HTML
- Add/list/resolve comments

APPS SCRIPT:
- Create Apps Script projects (standalone or bound to files)
- Execute Apps Script functions for advanced automation

Run with:  python server.py
Or via Claude Code config pointing to this file.
"""

import json
import re
import sys
from datetime import datetime, timedelta, timezone
from typing import Optional, List
from enum import Enum

from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel, Field, ConfigDict

from googleapiclient.discovery import build
from googleapiclient.http import MediaInMemoryUpload
from auth import get_credentials

# ── Initialize MCP Server ────────────────────────────────────────────────────

mcp = FastMCP("gdrive_mcp")

# ── Google API clients (lazy-initialized) ─────────────────────────────────────

_drive_service = None
_docs_service = None
_sheets_service = None
_slides_service = None
_scripts_service = None


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


# ── Shared helpers ────────────────────────────────────────────────────────────

GOOGLE_MIME_LABELS = {
    "application/vnd.google-apps.document": "Google Doc",
    "application/vnd.google-apps.spreadsheet": "Google Sheet",
    "application/vnd.google-apps.presentation": "Google Slides",
    "application/vnd.google-apps.folder": "Folder",
    "application/vnd.google-apps.form": "Google Form",
    "application/vnd.google-apps.drawing": "Google Drawing",
}

FILE_FIELDS = "id, name, mimeType, modifiedTime, createdTime, owners, webViewLink, parents, size, shared, trashed"


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


# ── Helpers for Doc Write Tools ──────────────────────────────────────────────

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


# ── Tool: Search ──────────────────────────────────────────────────────────────

class FileType(str, Enum):
    ANY = "any"
    DOCUMENT = "document"
    SPREADSHEET = "spreadsheet"
    PRESENTATION = "presentation"
    FOLDER = "folder"
    PDF = "pdf"
    IMAGE = "image"

FILE_TYPE_QUERIES = {
    FileType.ANY: "",
    FileType.DOCUMENT: "mimeType = 'application/vnd.google-apps.document'",
    FileType.SPREADSHEET: "mimeType = 'application/vnd.google-apps.spreadsheet'",
    FileType.PRESENTATION: "mimeType = 'application/vnd.google-apps.presentation'",
    FileType.FOLDER: "mimeType = 'application/vnd.google-apps.folder'",
    FileType.PDF: "mimeType = 'application/pdf'",
    FileType.IMAGE: "(mimeType contains 'image/')",
}


class SearchInput(BaseModel):
    """Input for searching Google Drive files."""
    model_config = ConfigDict(str_strip_whitespace=True)

    query: str = Field(
        ...,
        description="Search text. Searches file names and content. Use short keywords for best results.",
        min_length=1,
        max_length=500,
    )
    file_type: FileType = Field(
        default=FileType.ANY,
        description="Filter by file type: any, document, spreadsheet, presentation, folder, pdf, image",
    )
    search_content: bool = Field(
        default=True,
        description="If true, searches inside file content (fullText). If false, only searches file names.",
    )
    modified_after: Optional[str] = Field(
        default=None,
        description="Only files modified after this date. ISO format (YYYY-MM-DD) or relative like '7d', '30d', '6m', '1y'.",
    )
    owner_email: Optional[str] = Field(
        default=None,
        description="Filter to files owned by this email address.",
    )
    in_folder: Optional[str] = Field(
        default=None,
        description="Google Drive folder ID to search within (direct children only).",
    )
    shared_with_me: Optional[bool] = Field(
        default=None,
        description="If true, only files shared with you. If false, only your own files.",
    )
    max_results: int = Field(
        default=20,
        description="Maximum number of results to return.",
        ge=1,
        le=100,
    )


def parse_relative_date(date_str: str) -> Optional[str]:
    """Parse relative date strings like '7d', '30d', '6m', '1y' into ISO format."""
    match = re.match(r"^(\d+)([dDmMyY])$", date_str.strip())
    if match:
        num = int(match.group(1))
        unit = match.group(2).lower()
        now = datetime.now(timezone.utc)
        if unit == "d":
            dt = now - timedelta(days=num)
        elif unit == "m":
            dt = now - timedelta(days=num * 30)
        elif unit == "y":
            dt = now - timedelta(days=num * 365)
        else:
            return None
        return dt.strftime("%Y-%m-%dT%H:%M:%S")
    # Try parsing as ISO date
    try:
        dt = datetime.fromisoformat(date_str)
        return dt.strftime("%Y-%m-%dT%H:%M:%S")
    except ValueError:
        return None


@mcp.tool(
    name="gdrive_search",
    annotations={
        "title": "Search Google Drive",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def gdrive_search(params: SearchInput) -> str:
    """Search for files in Google Drive by name, content, type, date, owner, and folder.

    Supports filtering by file type (docs, sheets, slides, PDFs, images, folders),
    date ranges (absolute or relative like '7d' for last 7 days), owner email,
    folder location, and shared status.

    For best results, use short specific keywords rather than long phrases.
    Use search_content=false to search only file names (faster, more precise).

    Returns a list of matching files with IDs, links, types, and modification dates.
    Use the file ID with gdrive_read_doc or gdrive_read_sheet to read contents.

    Args:
        params: Search parameters including query text and filters.

    Returns:
        Markdown-formatted list of matching files with metadata.
    """
    # Build query parts
    parts = []

    # Text search
    query_text = params.query
    if params.search_content:
        parts.append(f"fullText contains '{query_text}'")
    else:
        parts.append(f"name contains '{query_text}'")

    # File type filter
    type_q = FILE_TYPE_QUERIES.get(params.file_type, "")
    if type_q:
        parts.append(type_q)

    # Date filter
    if params.modified_after:
        iso_date = parse_relative_date(params.modified_after)
        if iso_date:
            parts.append(f"modifiedTime > '{iso_date}'")

    # Owner filter
    if params.owner_email:
        parts.append(f"'{params.owner_email}' in owners")

    # Folder filter
    if params.in_folder:
        parts.append(f"'{params.in_folder}' in parents")

    # Shared filter
    if params.shared_with_me is True:
        parts.append("sharedWithMe = true")
    elif params.shared_with_me is False:
        parts.append("'me' in owners")

    # Always exclude trashed
    parts.append("trashed = false")

    query = " and ".join(parts)

    # Use relevance ordering for fullText, date ordering otherwise
    order = "relevance desc" if params.search_content else "modifiedTime desc"

    try:
        files = drive_query_files(query, params.max_results, order)
    except Exception as e:
        # If fullText search fails (e.g., special chars), fall back to name search
        if params.search_content:
            parts[0] = f"name contains '{query_text}'"
            query = " and ".join(parts)
            files = drive_query_files(query, params.max_results, "modifiedTime desc")
        else:
            return f"Error searching Drive: {e}"

    if not files:
        return f"No files found matching '{params.query}' with the given filters.\n\nTips:\n- Try shorter/different keywords\n- Set search_content=false to search only file names\n- Remove date or type filters to broaden the search"

    output = f"## Search Results ({len(files)} files)\n\n"
    for i, f in enumerate(files, 1):
        output += f"{i}. {format_file_entry(f)}\n\n"

    return output


# ── Tool: Read Google Doc ─────────────────────────────────────────────────────

class ReadDocInput(BaseModel):
    """Input for reading a Google Doc."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(
        ...,
        description="The Google Drive file ID of the document to read. Get this from search results.",
        min_length=1,
    )
    max_length: Optional[int] = Field(
        default=None,
        description="Maximum character length to return. None = full document. Use for large docs.",
        ge=100,
    )


def doc_element_to_markdown(element: dict, lists_state: dict) -> str:
    """Convert a Google Docs structural element to markdown."""
    output = ""

    if "paragraph" in element:
        para = element["paragraph"]
        style = para.get("paragraphStyle", {}).get("namedStyleType", "NORMAL_TEXT")
        bullet = para.get("bullet")

        # Extract text with inline formatting
        text_parts = []
        for elem in para.get("elements", []):
            if "textRun" in elem:
                run = elem["textRun"]
                content = run.get("content", "")
                ts = run.get("textStyle", {})

                # Apply inline formatting
                if ts.get("bold") and content.strip():
                    content = f"**{content.strip()}** "
                if ts.get("italic") and content.strip():
                    content = f"*{content.strip()}* "
                if ts.get("strikethrough") and content.strip():
                    content = f"~~{content.strip()}~~ "
                if ts.get("link", {}).get("url"):
                    url = ts["link"]["url"]
                    content = f"[{content.strip()}]({url}) "
                if ts.get("weightedFontFamily", {}).get("fontFamily") in ("Courier New", "Consolas", "monospace"):
                    if content.strip():
                        content = f"`{content.strip()}` "

                text_parts.append(content)

        text = "".join(text_parts).rstrip("\n")

        # Handle headings
        if style == "HEADING_1":
            output = f"# {text.strip()}\n\n"
        elif style == "HEADING_2":
            output = f"## {text.strip()}\n\n"
        elif style == "HEADING_3":
            output = f"### {text.strip()}\n\n"
        elif style == "HEADING_4":
            output = f"#### {text.strip()}\n\n"
        elif style == "HEADING_5":
            output = f"##### {text.strip()}\n\n"
        elif style == "HEADING_6":
            output = f"###### {text.strip()}\n\n"
        elif bullet:
            # Bullet/numbered list
            nesting = bullet.get("nestingLevel", 0)
            indent = "  " * nesting
            list_id = bullet.get("listId", "")

            # Check if ordered or unordered
            if list_id not in lists_state:
                lists_state[list_id] = {}
            if nesting not in lists_state[list_id]:
                lists_state[list_id][nesting] = 0
            lists_state[list_id][nesting] += 1

            # We'll use "- " for all lists (Google Docs API doesn't cleanly expose ordered vs unordered)
            output = f"{indent}- {text.strip()}\n"
        elif text.strip():
            output = f"{text.strip()}\n\n"
        else:
            output = "\n"

    elif "table" in element:
        table = element["table"]
        rows = table.get("tableRows", [])
        if rows:
            md_rows = []
            for row in rows:
                cells = []
                for cell in row.get("tableCells", []):
                    cell_text = ""
                    for content in cell.get("content", []):
                        cell_text += doc_element_to_markdown(content, lists_state).strip()
                    cells.append(cell_text.replace("|", "\\|").replace("\n", " "))
                md_rows.append("| " + " | ".join(cells) + " |")

            # Insert header separator after first row
            if len(md_rows) > 0:
                num_cols = len(rows[0].get("tableCells", []))
                separator = "| " + " | ".join(["---"] * num_cols) + " |"
                md_rows.insert(1, separator)

            output = "\n".join(md_rows) + "\n\n"

    return output


@mcp.tool(
    name="gdrive_read_doc",
    annotations={
        "title": "Read Google Doc",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def gdrive_read_doc(params: ReadDocInput) -> str:
    """Read the full content of a Google Doc and return it as clean markdown.

    Preserves headings, bold/italic formatting, links, lists, tables, and code.
    Use gdrive_search first to find the file ID.

    For non-Google-Doc files (PDFs, Word docs, etc.), exports as plain text.

    Args:
        params: File ID and optional max length.

    Returns:
        The document content as markdown text.
    """
    file_id = params.file_id

    # First check what type of file this is
    try:
        file_meta = get_drive().files().get(
            fileId=file_id, fields="id, name, mimeType, webViewLink",
            supportsAllDrives=True,
        ).execute()
    except Exception as e:
        return f"Error: Could not find file with ID `{file_id}`. Check the ID is correct.\n\nError: {e}"

    mime = file_meta.get("mimeType", "")
    name = file_meta.get("name", "untitled")

    # Google Doc — use Docs API for structured content
    if mime == "application/vnd.google-apps.document":
        try:
            doc = get_docs().documents().get(documentId=file_id).execute()
            title = doc.get("title", name)

            markdown = f"# {title}\n\n"
            lists_state: dict = {}

            for element in doc.get("body", {}).get("content", []):
                markdown += doc_element_to_markdown(element, lists_state)

            # Clean up excessive newlines
            markdown = re.sub(r"\n{3,}", "\n\n", markdown)

            if params.max_length and len(markdown) > params.max_length:
                markdown = markdown[: params.max_length] + f"\n\n... (truncated at {params.max_length} chars, full doc is {len(markdown)} chars)"

            return markdown

        except Exception as e:
            return f"Error reading Google Doc: {e}"

    # Non-Google-Doc files — try exporting as text
    try:
        # For Google Workspace files, export
        if mime.startswith("application/vnd.google-apps."):
            content = get_drive().files().export(
                fileId=file_id, mimeType="text/plain",
                supportsAllDrives=True,
            ).execute()
            text = content.decode("utf-8") if isinstance(content, bytes) else str(content)
        else:
            # For regular files, download content
            content = get_drive().files().get_media(fileId=file_id, supportsAllDrives=True).execute()
            text = content.decode("utf-8") if isinstance(content, bytes) else str(content)

        header = f"# {name}\n\n"
        text = header + text

        if params.max_length and len(text) > params.max_length:
            text = text[: params.max_length] + "\n\n... (truncated)"

        return text

    except Exception as e:
        return f"Error reading file: {e}\n\nFile: {name} ({mime})\nLink: {file_meta.get('webViewLink', 'N/A')}"


# ── Tool: Read Google Sheet ───────────────────────────────────────────────────

class ReadSheetInput(BaseModel):
    """Input for reading a Google Sheet."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(
        ...,
        description="The Google Drive file ID of the spreadsheet to read.",
        min_length=1,
    )
    sheet_name: Optional[str] = Field(
        default=None,
        description="Name of the specific sheet/tab to read. If omitted, reads the first sheet.",
    )
    range: Optional[str] = Field(
        default=None,
        description="Cell range in A1 notation (e.g., 'A1:D50', 'A:F'). If omitted, reads all data.",
    )
    max_rows: int = Field(
        default=200,
        description="Maximum number of rows to return.",
        ge=1,
        le=5000,
    )
    list_sheets_only: bool = Field(
        default=False,
        description="If true, only returns the list of sheet/tab names without reading data.",
    )


@mcp.tool(
    name="gdrive_read_sheet",
    annotations={
        "title": "Read Google Sheet",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def gdrive_read_sheet(params: ReadSheetInput) -> str:
    """Read data from a Google Sheet and return it as a markdown table.

    Can read specific sheets/tabs, specific cell ranges, and limits rows for large sheets.
    Use list_sheets_only=true to see available tabs before reading.

    Args:
        params: File ID, optional sheet name, range, and row limit.

    Returns:
        Sheet data as a markdown table, or list of sheet names.
    """
    file_id = params.file_id

    try:
        # Get spreadsheet metadata (sheet names, etc.)
        spreadsheet = get_sheets().spreadsheets().get(
            spreadsheetId=file_id
        ).execute()
    except Exception as e:
        return f"Error: Could not open spreadsheet `{file_id}`. Check the ID is correct.\n\nError: {e}"

    title = spreadsheet.get("properties", {}).get("title", "Untitled")
    sheets = spreadsheet.get("sheets", [])
    sheet_names = [s["properties"]["title"] for s in sheets]

    # If just listing sheets
    if params.list_sheets_only:
        output = f"## {title}\n\nSheets/tabs in this spreadsheet:\n\n"
        for i, name in enumerate(sheet_names, 1):
            row_count = sheets[i - 1]["properties"].get("gridProperties", {}).get("rowCount", "?")
            col_count = sheets[i - 1]["properties"].get("gridProperties", {}).get("columnCount", "?")
            output += f"{i}. **{name}** ({row_count} rows x {col_count} cols)\n"
        return output

    # Determine which sheet to read
    target_sheet = params.sheet_name or sheet_names[0]
    if target_sheet not in sheet_names:
        return f"Sheet '{target_sheet}' not found. Available sheets: {', '.join(sheet_names)}"

    # Build range string
    if params.range:
        range_str = f"'{target_sheet}'!{params.range}"
    else:
        range_str = f"'{target_sheet}'"

    try:
        result = get_sheets().spreadsheets().values().get(
            spreadsheetId=file_id,
            range=range_str,
        ).execute()
    except Exception as e:
        return f"Error reading sheet data: {e}"

    rows = result.get("values", [])
    if not rows:
        return f"Sheet '{target_sheet}' is empty (no data found)."

    # Truncate if needed
    total_rows = len(rows)
    if total_rows > params.max_rows:
        rows = rows[: params.max_rows]

    # Build markdown table
    output = f"## {title} -- {target_sheet}\n\n"

    # Normalize column widths (some rows may have fewer columns)
    max_cols = max(len(row) for row in rows)
    normalized = [row + [""] * (max_cols - len(row)) for row in rows]

    # Header row
    header = normalized[0]
    output += "| " + " | ".join(str(c).replace("|", "\\|") for c in header) + " |\n"
    output += "| " + " | ".join(["---"] * max_cols) + " |\n"

    # Data rows
    for row in normalized[1:]:
        output += "| " + " | ".join(str(c).replace("|", "\\|") for c in row) + " |\n"

    if total_rows > params.max_rows:
        output += f"\n*Showing {params.max_rows} of {total_rows} rows. Use max_rows or range to adjust.*\n"

    output += f"\n*{total_rows} rows x {max_cols} columns*"

    return output


# ── Tool: List Folder ─────────────────────────────────────────────────────────

class ListFolderInput(BaseModel):
    """Input for listing files in a Google Drive folder."""
    model_config = ConfigDict(str_strip_whitespace=True)

    folder_id: Optional[str] = Field(
        default=None,
        description="Folder ID to list. Use 'root' or omit for your Drive root. Get folder IDs from search results.",
    )
    file_type: FileType = Field(
        default=FileType.ANY,
        description="Filter by file type.",
    )
    max_results: int = Field(
        default=50,
        description="Maximum number of files to return.",
        ge=1,
        le=100,
    )


@mcp.tool(
    name="gdrive_list_folder",
    annotations={
        "title": "List Google Drive Folder",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def gdrive_list_folder(params: ListFolderInput) -> str:
    """List files and subfolders in a Google Drive folder.

    Omit folder_id to list your Drive root. Use file_type to filter results.
    Returns files sorted by type (folders first) then by modification date.

    Args:
        params: Folder ID, optional type filter, and max results.

    Returns:
        Markdown list of files in the folder with types, dates, and IDs.
    """
    folder_id = params.folder_id or "root"

    parts = [f"'{folder_id}' in parents", "trashed = false"]

    type_q = FILE_TYPE_QUERIES.get(params.file_type, "")
    if type_q:
        parts.append(type_q)

    query = " and ".join(parts)

    try:
        files = drive_query_files(query, params.max_results, "folder, modifiedTime desc")
    except Exception as e:
        return f"Error listing folder: {e}"

    if not files:
        return "Folder is empty or not found."

    # Get folder name
    folder_name = "My Drive"
    if folder_id != "root":
        try:
            meta = get_drive().files().get(fileId=folder_id, fields="name", supportsAllDrives=True).execute()
            folder_name = meta.get("name", folder_id)
        except Exception:
            folder_name = folder_id

    output = f"## {folder_name}\n\n"

    folders = [f for f in files if f.get("mimeType") == "application/vnd.google-apps.folder"]
    other = [f for f in files if f.get("mimeType") != "application/vnd.google-apps.folder"]

    if folders:
        output += "### Folders\n\n"
        for f in folders:
            output += f"**{f['name']}** -- ID: `{f['id']}`\n\n"

    if other:
        output += "### Files\n\n"
        for i, f in enumerate(other, 1):
            output += f"{i}. {format_file_entry(f)}\n\n"

    return output


# ── Tool: Recent Files ────────────────────────────────────────────────────────

class RecentFilesInput(BaseModel):
    """Input for listing recently modified files."""
    model_config = ConfigDict(str_strip_whitespace=True)

    days: int = Field(
        default=7,
        description="Number of days to look back.",
        ge=1,
        le=365,
    )
    file_type: FileType = Field(
        default=FileType.ANY,
        description="Filter by file type.",
    )
    owned_by_me: bool = Field(
        default=False,
        description="If true, only files you own. If false, includes shared files.",
    )
    max_results: int = Field(
        default=20,
        description="Maximum number of files to return.",
        ge=1,
        le=100,
    )


@mcp.tool(
    name="gdrive_recent",
    annotations={
        "title": "Recent Google Drive Files",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def gdrive_recent(params: RecentFilesInput) -> str:
    """List recently modified files in Google Drive.

    Great for finding files you or others worked on recently.
    Can filter by file type and ownership.

    Args:
        params: Time window, type filter, ownership filter, and max results.

    Returns:
        Markdown list of recently modified files sorted by modification date.
    """
    cutoff = datetime.now(timezone.utc) - timedelta(days=params.days)
    cutoff_str = cutoff.strftime("%Y-%m-%dT%H:%M:%S")

    parts = [f"modifiedTime > '{cutoff_str}'", "trashed = false"]

    type_q = FILE_TYPE_QUERIES.get(params.file_type, "")
    if type_q:
        parts.append(type_q)

    if params.owned_by_me:
        parts.append("'me' in owners")

    query = " and ".join(parts)
    files = drive_query_files(query, params.max_results, "modifiedTime desc")

    if not files:
        return f"No files modified in the last {params.days} days matching your filters."

    output = f"## Recently Modified Files (last {params.days} days)\n\n"
    for i, f in enumerate(files, 1):
        output += f"{i}. {format_file_entry(f)}\n\n"

    return output


# ── Tool: File Info ───────────────────────────────────────────────────────────

class FileInfoInput(BaseModel):
    """Input for getting detailed file metadata."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(
        ...,
        description="The Google Drive file ID.",
        min_length=1,
    )


@mcp.tool(
    name="gdrive_file_info",
    annotations={
        "title": "Google Drive File Info",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def gdrive_file_info(params: FileInfoInput) -> str:
    """Get detailed metadata about a Google Drive file.

    Returns name, type, size, owners, sharing status, dates, parent folder,
    and direct link. Useful for understanding a file before reading it.

    Args:
        params: The file ID to look up.

    Returns:
        Markdown-formatted file metadata.
    """
    try:
        f = get_drive().files().get(
            fileId=params.file_id,
            fields="id, name, mimeType, modifiedTime, createdTime, owners, webViewLink, "
                   "parents, size, shared, sharingUser, lastModifyingUser, description, "
                   "starred, capabilities",
            supportsAllDrives=True,
        ).execute()
    except Exception as e:
        return f"Error: Could not find file `{params.file_id}`.\n\nError: {e}"

    mime = f.get("mimeType", "")
    label = GOOGLE_MIME_LABELS.get(mime, mime)

    output = f"## {f.get('name', 'Untitled')}\n\n"
    output += f"**Type:** {label}\n"
    output += f"**ID:** `{f.get('id')}`\n"
    if f.get("webViewLink"):
        output += f"**Link:** {f['webViewLink']}\n"
    if f.get("size"):
        size_mb = int(f["size"]) / (1024 * 1024)
        output += f"**Size:** {size_mb:.2f} MB\n"
    output += f"**Created:** {f.get('createdTime', 'N/A')}\n"
    output += f"**Modified:** {f.get('modifiedTime', 'N/A')}\n"

    owners = f.get("owners", [])
    if owners:
        output += f"**Owner:** {', '.join(o.get('displayName', o.get('emailAddress', '?')) for o in owners)}\n"

    last_mod = f.get("lastModifyingUser", {})
    if last_mod:
        output += f"**Last modified by:** {last_mod.get('displayName', last_mod.get('emailAddress', '?'))}\n"

    output += f"**Shared:** {'Yes' if f.get('shared') else 'No'}\n"
    output += f"**Starred:** {'Yes' if f.get('starred') else 'No'}\n"

    if f.get("description"):
        output += f"\n**Description:** {f['description']}\n"

    # Get parent folder name
    parents = f.get("parents", [])
    if parents:
        try:
            parent = get_drive().files().get(fileId=parents[0], fields="name, id", supportsAllDrives=True).execute()
            output += f"\n**In folder:** {parent.get('name', 'Unknown')} (`{parent.get('id')}`)\n"
        except Exception:
            pass

    return output


# ══════════════════════════════════════════════════════════════════════════════
# WRITE TOOLS — Google Docs
# ══════════════════════════════════════════════════════════════════════════════

# ── Tool: Find & Replace ─────────────────────────────────────────────────────

class FindReplaceInput(BaseModel):
    """Input for find-and-replace in a Google Doc."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(
        ...,
        description="The Google Drive file ID of the document to edit.",
        min_length=1,
    )
    find_text: str = Field(
        ...,
        description="The text to find in the document.",
        min_length=1,
    )
    replace_text: str = Field(
        ...,
        description="The text to replace it with. Use empty string to delete matches.",
    )
    match_case: bool = Field(
        default=True,
        description="If true (default), search is case-sensitive.",
    )


@mcp.tool(
    name="gdrive_find_replace",
    annotations={
        "title": "Find & Replace in Google Doc",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def gdrive_find_replace(params: FindReplaceInput) -> str:
    """Find and replace text across an entire Google Doc.

    This is the primary editing tool — no character indices needed. Just provide
    the text to find and what to replace it with. Works across the entire document
    including headers, footers, and footnotes.

    Use replace_text="" to delete all occurrences of the found text.

    Args:
        params: File ID, find text, replace text, and case sensitivity.

    Returns:
        Summary of replacements made.
    """
    try:
        result = get_docs().documents().batchUpdate(
            documentId=params.file_id,
            body={
                "requests": [
                    {
                        "replaceAllText": {
                            "containsText": {
                                "text": params.find_text,
                                "matchCase": params.match_case,
                            },
                            "replaceText": params.replace_text,
                        }
                    }
                ]
            },
        ).execute()

        # Extract replacement count from response
        replies = result.get("replies", [{}])
        count = replies[0].get("replaceAllText", {}).get("occurrencesChanged", 0)

        if count == 0:
            return f"No occurrences of '{params.find_text}' found in the document."

        action = "deleted" if params.replace_text == "" else f"replaced with '{params.replace_text}'"
        return f"Replaced {count} occurrence(s) of '{params.find_text}' -- {action}."

    except Exception as e:
        return f"Error performing find & replace: {e}"


# ── Tool: Insert Text ────────────────────────────────────────────────────────

class InsertLocation(str, Enum):
    AT_END = "at_end"
    AT_START = "at_start"
    AFTER_HEADING = "after_heading"
    AT_INDEX = "at_index"


class InsertTextInput(BaseModel):
    """Input for inserting text into a Google Doc."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(
        ...,
        description="The Google Drive file ID of the document to edit.",
        min_length=1,
    )
    text: str = Field(
        ...,
        description="The text to insert. Use \\n for newlines.",
        min_length=1,
    )
    location: InsertLocation = Field(
        default=InsertLocation.AT_END,
        description="Where to insert: 'at_end', 'at_start', 'after_heading', or 'at_index'.",
    )
    heading_text: Optional[str] = Field(
        default=None,
        description="Required when location='after_heading'. The heading text to insert after (case-insensitive partial match).",
    )
    index: Optional[int] = Field(
        default=None,
        description="Required when location='at_index'. The character index to insert at.",
        ge=1,
    )
    bold: bool = Field(
        default=False,
        description="If true, the inserted text will be bold.",
    )
    italic: bool = Field(
        default=False,
        description="If true, the inserted text will be italic.",
    )
    heading_level: Optional[int] = Field(
        default=None,
        description="If set (1-6), format the inserted text as a heading at this level.",
        ge=1,
        le=6,
    )


@mcp.tool(
    name="gdrive_insert_text",
    annotations={
        "title": "Insert Text in Google Doc",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def gdrive_insert_text(params: InsertTextInput) -> str:
    """Insert text at a specific location in a Google Doc.

    Can insert at the end, start, after a specific heading, or at a character index.
    Supports bold, italic, and heading-level formatting on inserted text.

    For 'after_heading', provide the heading text — it does a case-insensitive
    partial match against all headings in the doc.

    Args:
        params: File ID, text, location, and optional formatting.

    Returns:
        Confirmation of the insertion.
    """
    file_id = params.file_id

    # Read document to determine insertion index
    try:
        doc = get_docs().documents().get(documentId=file_id).execute()
    except Exception as e:
        return f"Error reading document: {e}"

    # Determine insertion index
    if params.location == InsertLocation.AT_END:
        body = doc.get("body", {})
        content = body.get("content", [])
        if content:
            insert_index = content[-1].get("endIndex", 1) - 1
        else:
            insert_index = 1
    elif params.location == InsertLocation.AT_START:
        insert_index = 1
    elif params.location == InsertLocation.AFTER_HEADING:
        if not params.heading_text:
            return "Error: heading_text is required when location='after_heading'."
        end_idx = find_heading_end_index(doc, params.heading_text)
        if end_idx is None:
            return f"Error: No heading matching '{params.heading_text}' found in the document."
        insert_index = end_idx
    elif params.location == InsertLocation.AT_INDEX:
        if params.index is None:
            return "Error: index is required when location='at_index'."
        insert_index = params.index
    else:
        return f"Error: Unknown location '{params.location}'."

    # Build requests
    requests = []

    # Insert the text
    text_to_insert = params.text
    if not text_to_insert.endswith("\n"):
        text_to_insert += "\n"

    requests.append({
        "insertText": {
            "location": {"index": insert_index},
            "text": text_to_insert,
        }
    })

    # Apply text styling (bold/italic)
    if params.bold or params.italic:
        style = {}
        fields = []
        if params.bold:
            style["bold"] = True
            fields.append("bold")
        if params.italic:
            style["italic"] = True
            fields.append("italic")

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

    # Apply heading style
    if params.heading_level:
        heading_style = f"HEADING_{params.heading_level}"
        requests.append({
            "updateParagraphStyle": {
                "range": {
                    "startIndex": insert_index,
                    "endIndex": insert_index + len(text_to_insert),
                },
                "paragraphStyle": {"namedStyleType": heading_style},
                "fields": "namedStyleType",
            }
        })

    try:
        get_docs().documents().batchUpdate(
            documentId=file_id,
            body={"requests": requests},
        ).execute()

        location_desc = {
            InsertLocation.AT_END: "at the end",
            InsertLocation.AT_START: "at the start",
            InsertLocation.AFTER_HEADING: f"after heading '{params.heading_text}'",
            InsertLocation.AT_INDEX: f"at index {params.index}",
        }

        formatting = []
        if params.bold:
            formatting.append("bold")
        if params.italic:
            formatting.append("italic")
        if params.heading_level:
            formatting.append(f"heading {params.heading_level}")
        fmt_str = f" ({', '.join(formatting)})" if formatting else ""

        return f"Inserted {len(params.text)} chars {location_desc[params.location]}{fmt_str}."

    except Exception as e:
        return f"Error inserting text: {e}"


# ── Tool: Delete Text ────────────────────────────────────────────────────────

class DeleteTextInput(BaseModel):
    """Input for deleting text from a Google Doc."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(
        ...,
        description="The Google Drive file ID of the document to edit.",
        min_length=1,
    )
    text: Optional[str] = Field(
        default=None,
        description="Text string to find and delete. Deletes all occurrences. Use this OR start_index/end_index, not both.",
    )
    start_index: Optional[int] = Field(
        default=None,
        description="Start character index for range deletion. Use with end_index.",
        ge=1,
    )
    end_index: Optional[int] = Field(
        default=None,
        description="End character index for range deletion (exclusive). Use with start_index.",
        ge=2,
    )


@mcp.tool(
    name="gdrive_delete_text",
    annotations={
        "title": "Delete Text from Google Doc",
        "readOnlyHint": False,
        "destructiveHint": True,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def gdrive_delete_text(params: DeleteTextInput) -> str:
    """Delete text from a Google Doc by matching text or by index range.

    Two modes:
    1. By text match: Provide 'text' to find and delete all occurrences.
    2. By index range: Provide 'start_index' and 'end_index' for precise deletion.

    For text match mode, use gdrive_read_doc first to see the document content
    and identify what to delete.

    Args:
        params: File ID and either text to match or index range.

    Returns:
        Summary of what was deleted.
    """
    file_id = params.file_id

    if params.text:
        # Text-match mode: find all occurrences and delete them
        try:
            doc = get_docs().documents().get(documentId=file_id).execute()
        except Exception as e:
            return f"Error reading document: {e}"

        indices = find_text_indices(doc, params.text)
        if not indices:
            return f"Text '{params.text}' not found in the document."

        # Delete in reverse order to preserve indices
        requests = []
        for start, end in reversed(indices):
            requests.append({
                "deleteContentRange": {
                    "range": {
                        "startIndex": start,
                        "endIndex": end,
                    }
                }
            })

        try:
            get_docs().documents().batchUpdate(
                documentId=file_id,
                body={"requests": requests},
            ).execute()
            return f"Deleted {len(indices)} occurrence(s) of '{params.text}'."
        except Exception as e:
            return f"Error deleting text: {e}"

    elif params.start_index is not None and params.end_index is not None:
        # Index range mode
        if params.end_index <= params.start_index:
            return "Error: end_index must be greater than start_index."

        try:
            get_docs().documents().batchUpdate(
                documentId=file_id,
                body={
                    "requests": [
                        {
                            "deleteContentRange": {
                                "range": {
                                    "startIndex": params.start_index,
                                    "endIndex": params.end_index,
                                }
                            }
                        }
                    ]
                },
            ).execute()
            chars = params.end_index - params.start_index
            return f"Deleted {chars} characters (index {params.start_index} to {params.end_index})."
        except Exception as e:
            return f"Error deleting text range: {e}"

    else:
        return "Error: Provide either 'text' (to delete by match) or both 'start_index' and 'end_index' (to delete by range)."


# ── Tool: Create Doc ─────────────────────────────────────────────────────────

class CreateDocInput(BaseModel):
    """Input for creating a new Google Doc."""
    model_config = ConfigDict(str_strip_whitespace=True)

    title: str = Field(
        ...,
        description="Title of the new Google Doc.",
        min_length=1,
        max_length=500,
    )
    content: str = Field(
        ...,
        description="Content for the document in markdown format. Headings (#, ##, ###), bold (**), italic (*), and lists (- ) are supported.",
        min_length=1,
    )
    folder_id: Optional[str] = Field(
        default=None,
        description="Google Drive folder ID to create the doc in. If omitted, creates in Drive root.",
    )


@mcp.tool(
    name="gdrive_create_doc",
    annotations={
        "title": "Create Google Doc",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def gdrive_create_doc(params: CreateDocInput) -> str:
    """Create a new Google Doc from markdown content.

    Converts markdown to HTML and uploads as a native Google Doc with formatting
    preserved (headings, bold, italic, lists, links).

    Args:
        params: Title, markdown content, and optional folder ID.

    Returns:
        The new document's ID and link.
    """
    import markdown as md

    # Convert markdown to HTML
    html_content = md.markdown(
        params.content,
        extensions=["tables", "fenced_code"],
    )

    # Wrap in basic HTML structure
    html = f"<html><body><h1>{params.title}</h1>{html_content}</body></html>"
    html_bytes = html.encode("utf-8")

    # Build file metadata
    file_metadata = {
        "name": params.title,
        "mimeType": "application/vnd.google-apps.document",
    }
    if params.folder_id:
        file_metadata["parents"] = [params.folder_id]

    try:
        media = MediaInMemoryUpload(
            html_bytes,
            mimetype="text/html",
            resumable=False,
        )
        created = get_drive().files().create(
            body=file_metadata,
            media_body=media,
            fields="id, name, webViewLink",
            supportsAllDrives=True,
        ).execute()

        doc_id = created.get("id")
        link = created.get("webViewLink", "")
        name = created.get("name", params.title)

        return f"Created Google Doc: **{name}**\n\nID: `{doc_id}`\nLink: {link}"

    except Exception as e:
        return f"Error creating document: {e}"


# ══════════════════════════════════════════════════════════════════════════════
# WRITE TOOLS — Google Sheets
# ══════════════════════════════════════════════════════════════════════════════

# ── Tool: Write Sheet ────────────────────────────────────────────────────────

class WriteSheetInput(BaseModel):
    """Input for writing values to a Google Sheet."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(
        ...,
        description="The Google Drive file ID of the spreadsheet.",
        min_length=1,
    )
    range: str = Field(
        ...,
        description="Cell range in A1 notation (e.g., 'A1:D10', 'Sheet2!B5:C8'). Include sheet name with ! prefix for non-first sheets.",
        min_length=1,
    )
    values: List[List[str]] = Field(
        ...,
        description="2D array of values to write. Each inner list is a row. Example: [['Name', 'Age'], ['Alice', '30'], ['Bob', '25']]",
        min_length=1,
    )


@mcp.tool(
    name="gdrive_write_sheet",
    annotations={
        "title": "Write to Google Sheet",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def gdrive_write_sheet(params: WriteSheetInput) -> str:
    """Write values to specific cells in a Google Sheet.

    Overwrites the target range with the provided values. Use A1 notation for
    the range (e.g., 'A1:C3'). Include sheet name for non-first tabs
    (e.g., 'Sheet2!A1:C3').

    Values are processed with USER_ENTERED mode, so formulas (=SUM(...)),
    numbers, dates, and text are handled intelligently.

    Args:
        params: File ID, target range, and 2D array of values.

    Returns:
        Summary of cells updated.
    """
    try:
        result = get_sheets().spreadsheets().values().update(
            spreadsheetId=params.file_id,
            range=params.range,
            valueInputOption="USER_ENTERED",
            body={"values": params.values},
        ).execute()

        updated_range = result.get("updatedRange", params.range)
        updated_rows = result.get("updatedRows", 0)
        updated_cols = result.get("updatedColumns", 0)
        updated_cells = result.get("updatedCells", 0)

        return f"Updated {updated_cells} cells ({updated_rows} rows x {updated_cols} cols) in range `{updated_range}`."

    except Exception as e:
        return f"Error writing to sheet: {e}"


# ── Tool: Append Sheet ───────────────────────────────────────────────────────

class AppendSheetInput(BaseModel):
    """Input for appending rows to a Google Sheet."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(
        ...,
        description="The Google Drive file ID of the spreadsheet.",
        min_length=1,
    )
    sheet_name: Optional[str] = Field(
        default=None,
        description="Name of the sheet/tab to append to. If omitted, appends to the first sheet.",
    )
    values: List[List[str]] = Field(
        ...,
        description="2D array of rows to append. Each inner list is a row. Example: [['Alice', '30'], ['Bob', '25']]",
        min_length=1,
    )


@mcp.tool(
    name="gdrive_append_sheet",
    annotations={
        "title": "Append Rows to Google Sheet",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    },
)
async def gdrive_append_sheet(params: AppendSheetInput) -> str:
    """Append rows to the end of a Google Sheet.

    Automatically finds the last row with data and appends below it.
    Values are processed with USER_ENTERED mode.

    Args:
        params: File ID, optional sheet name, and rows to append.

    Returns:
        Summary of rows appended.
    """
    # Build range — sheet name prefix if provided
    range_str = f"'{params.sheet_name}'!A:A" if params.sheet_name else "A:A"

    try:
        result = get_sheets().spreadsheets().values().append(
            spreadsheetId=params.file_id,
            range=range_str,
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={"values": params.values},
        ).execute()

        updates = result.get("updates", {})
        updated_range = updates.get("updatedRange", "")
        updated_rows = updates.get("updatedRows", len(params.values))

        return f"Appended {updated_rows} row(s) to `{updated_range}`."

    except Exception as e:
        return f"Error appending to sheet: {e}"


# ══════════════════════════════════════════════════════════════════════════════
# READ TOOLS — Extended
# ══════════════════════════════════════════════════════════════════════════════

# ── Tool: Read Section ────────────────────────────────────────────────────────

class ReadSectionInput(BaseModel):
    """Input for reading a section of a Google Doc."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(..., description="The Google Drive file ID.", min_length=1)
    heading: str = Field(
        ...,
        description="The heading text to read under (case-insensitive partial match). Returns all content from this heading until the next heading of equal or higher level.",
        min_length=1,
    )


@mcp.tool(
    name="gdrive_read_section",
    annotations={"title": "Read Section of Google Doc", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def gdrive_read_section(params: ReadSectionInput) -> str:
    """Read a specific section of a Google Doc by heading name.

    Returns all content from the matched heading until the next heading of
    equal or higher level (or end of document). Case-insensitive partial match.

    Args:
        params: File ID and heading text to match.

    Returns:
        The section content as markdown.
    """
    try:
        doc = get_docs().documents().get(documentId=params.file_id).execute()
    except Exception as e:
        return f"Error reading document: {e}"

    elements = doc.get("body", {}).get("content", [])
    title = doc.get("title", "Untitled")
    lists_state: dict = {}

    # Find the target heading
    target_level = None
    capturing = False
    result_md = ""

    for element in elements:
        if "paragraph" in element:
            para = element["paragraph"]
            style = para.get("paragraphStyle", {}).get("namedStyleType", "NORMAL_TEXT")

            if style.startswith("HEADING_"):
                level = int(style.split("_")[1])
                para_text = ""
                for elem in para.get("elements", []):
                    if "textRun" in elem:
                        para_text += elem["textRun"].get("content", "")

                if not capturing and params.heading.strip().lower() in para_text.strip().lower():
                    capturing = True
                    target_level = level
                    result_md += doc_element_to_markdown(element, lists_state)
                    continue
                elif capturing and level <= target_level:
                    break

        if capturing:
            result_md += doc_element_to_markdown(element, lists_state)

    if not result_md:
        # List available headings for the user
        headings = []
        for element in elements:
            if "paragraph" in element:
                style = element["paragraph"].get("paragraphStyle", {}).get("namedStyleType", "NORMAL_TEXT")
                if style.startswith("HEADING_"):
                    text = ""
                    for elem in element["paragraph"].get("elements", []):
                        if "textRun" in elem:
                            text += elem["textRun"].get("content", "")
                    headings.append(f"- {text.strip()}")
        heading_list = "\n".join(headings) if headings else "No headings found."
        return f"No section matching '{params.heading}' found in '{title}'.\n\nAvailable headings:\n{heading_list}"

    return re.sub(r"\n{3,}", "\n\n", result_md)


# ── Tool: Insert Table ───────────────────────────────────────────────────────

class InsertTableInput(BaseModel):
    """Input for inserting a table into a Google Doc."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(..., description="The Google Drive file ID.", min_length=1)
    rows: List[List[str]] = Field(
        ...,
        description="2D array of cell values. First row is treated as header. Example: [['Name', 'Age'], ['Alice', '30']]",
        min_length=1,
    )
    after_heading: Optional[str] = Field(
        default=None,
        description="Insert table after this heading (case-insensitive partial match). If omitted, appends to end of document.",
    )


@mcp.tool(
    name="gdrive_insert_table",
    annotations={"title": "Insert Table in Google Doc", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
)
async def gdrive_insert_table(params: InsertTableInput) -> str:
    """Insert a table into a Google Doc with cell values.

    Creates a table with the given rows/columns and populates all cells.
    First row is the header. Optionally insert after a specific heading.

    Args:
        params: File ID, 2D array of cell values, optional heading.

    Returns:
        Confirmation with table dimensions.
    """
    try:
        doc = get_docs().documents().get(documentId=params.file_id).execute()
    except Exception as e:
        return f"Error reading document: {e}"

    # Determine insertion index
    if params.after_heading:
        insert_index = find_heading_end_index(doc, params.after_heading)
        if insert_index is None:
            return f"Error: No heading matching '{params.after_heading}' found."
    else:
        content = doc.get("body", {}).get("content", [])
        insert_index = content[-1].get("endIndex", 1) - 1 if content else 1

    num_rows = len(params.rows)
    num_cols = max(len(row) for row in params.rows) if params.rows else 1

    # Normalize rows to same column count
    normalized = [row + [""] * (num_cols - len(row)) for row in params.rows]

    requests = []

    # 1. Insert the table structure
    requests.append({
        "insertTable": {
            "rows": num_rows,
            "columns": num_cols,
            "location": {"index": insert_index},
        }
    })

    try:
        # Insert the empty table first
        get_docs().documents().batchUpdate(
            documentId=params.file_id,
            body={"requests": requests},
        ).execute()

        # Re-read doc to get table cell indices
        doc = get_docs().documents().get(documentId=params.file_id).execute()

        # Find the table we just inserted (scan for table elements)
        table_element = None
        for element in doc.get("body", {}).get("content", []):
            if "table" in element:
                start = element.get("startIndex", 0)
                if start >= insert_index:
                    table_element = element
                    break

        if not table_element:
            return "Table created but could not locate it to populate cells."

        # Populate cells in reverse order to preserve indices
        cell_requests = []
        table_rows = table_element["table"].get("tableRows", [])
        for r_idx in range(len(table_rows) - 1, -1, -1):
            cells = table_rows[r_idx].get("tableCells", [])
            for c_idx in range(len(cells) - 1, -1, -1):
                if r_idx < len(normalized) and c_idx < len(normalized[r_idx]):
                    cell_text = normalized[r_idx][c_idx]
                    if cell_text:
                        # Each cell has content -> paragraph -> insert at paragraph start
                        cell_content = cells[c_idx].get("content", [])
                        if cell_content:
                            cell_start = cell_content[0].get("startIndex", 0)
                            cell_requests.append({
                                "insertText": {
                                    "location": {"index": cell_start},
                                    "text": cell_text,
                                }
                            })

        if cell_requests:
            get_docs().documents().batchUpdate(
                documentId=params.file_id,
                body={"requests": cell_requests},
            ).execute()

        location = f"after heading '{params.after_heading}'" if params.after_heading else "at end of document"
        return f"Inserted {num_rows}x{num_cols} table {location} with {sum(1 for r in normalized for c in r if c)} populated cells."

    except Exception as e:
        return f"Error inserting table: {e}"


# ══════════════════════════════════════════════════════════════════════════════
# WRITE TOOLS — Google Sheets (Extended)
# ══════════════════════════════════════════════════════════════════════════════

# ── Tool: Format Cells ───────────────────────────────────────────────────────

class FormatCellsInput(BaseModel):
    """Input for formatting cells in a Google Sheet."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(..., description="The Google Drive file ID of the spreadsheet.", min_length=1)
    sheet_name: Optional[str] = Field(default=None, description="Sheet/tab name. If omitted, uses first sheet.")
    range: str = Field(
        ...,
        description="Cell range in A1 notation (e.g., 'A1:D1', 'B2:E10'). Do NOT include sheet name here.",
        min_length=1,
    )
    bold: Optional[bool] = Field(default=None, description="Set bold on/off.")
    italic: Optional[bool] = Field(default=None, description="Set italic on/off.")
    font_size: Optional[int] = Field(default=None, description="Font size in points.", ge=1, le=400)
    font_color: Optional[str] = Field(default=None, description="Font color as hex (e.g., '#FF0000' for red).")
    bg_color: Optional[str] = Field(default=None, description="Background color as hex (e.g., '#FFFF00' for yellow).")
    number_format: Optional[str] = Field(
        default=None,
        description="Number format pattern (e.g., '#,##0.00', '0%', '$#,##0', 'yyyy-mm-dd').",
    )
    horizontal_align: Optional[str] = Field(
        default=None,
        description="Horizontal alignment: 'LEFT', 'CENTER', or 'RIGHT'.",
    )
    borders: Optional[str] = Field(
        default=None,
        description="Border style: 'all' (all borders), 'outline' (outer only), 'none' (remove).",
    )
    merge: Optional[bool] = Field(default=None, description="If true, merge the range into one cell. If false, unmerge.")
    wrap: Optional[str] = Field(default=None, description="Text wrap: 'WRAP', 'CLIP', or 'OVERFLOW'.")


def hex_to_color(hex_str: str) -> dict:
    """Convert hex color string to Google Sheets color dict."""
    hex_str = hex_str.lstrip("#")
    r = int(hex_str[0:2], 16) / 255.0
    g = int(hex_str[2:4], 16) / 255.0
    b = int(hex_str[4:6], 16) / 255.0
    return {"red": r, "green": g, "blue": b}


def a1_to_grid_range(a1_range: str, sheet_id: int) -> dict:
    """Convert A1 notation to GridRange dict."""
    match = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", a1_range.upper())
    if not match:
        # Try single cell like A1
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


@mcp.tool(
    name="gdrive_format_cells",
    annotations={"title": "Format Cells in Google Sheet", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def gdrive_format_cells(params: FormatCellsInput) -> str:
    """Apply formatting to cells in a Google Sheet.

    Supports bold, italic, font size, font/background colors, number formats,
    alignment, borders, merge/unmerge, and text wrapping. All formatting
    options are optional — only set the ones you want to change.

    Args:
        params: File ID, range, and formatting options.

    Returns:
        Confirmation of formatting applied.
    """
    try:
        spreadsheet = get_sheets().spreadsheets().get(spreadsheetId=params.file_id).execute()
        sheet_id = get_sheet_id(spreadsheet, params.sheet_name)
    except Exception as e:
        return f"Error: {e}"

    try:
        grid_range = a1_to_grid_range(params.range, sheet_id)
    except ValueError as e:
        return f"Error: {e}"

    requests = []
    applied = []

    # Build cell format
    cell_format = {}
    fields = []

    if params.bold is not None or params.italic is not None or params.font_size or params.font_color:
        text_format = {}
        tf_fields = []
        if params.bold is not None:
            text_format["bold"] = params.bold
            tf_fields.append("bold")
        if params.italic is not None:
            text_format["italic"] = params.italic
            tf_fields.append("italic")
        if params.font_size:
            text_format["fontSize"] = params.font_size
            tf_fields.append("fontSize")
        if params.font_color:
            text_format["foregroundColorStyle"] = {"rgbColor": hex_to_color(params.font_color)}
            tf_fields.append("foregroundColorStyle")
        cell_format["textFormat"] = text_format
        fields.extend(f"userEnteredFormat.textFormat.{f}" for f in tf_fields)
        applied.append("text formatting")

    if params.bg_color:
        cell_format["backgroundColorStyle"] = {"rgbColor": hex_to_color(params.bg_color)}
        fields.append("userEnteredFormat.backgroundColorStyle")
        applied.append(f"background color {params.bg_color}")

    if params.number_format:
        cell_format["numberFormat"] = {"type": "NUMBER", "pattern": params.number_format}
        fields.append("userEnteredFormat.numberFormat")
        applied.append(f"number format '{params.number_format}'")

    if params.horizontal_align:
        cell_format["horizontalAlignment"] = params.horizontal_align.upper()
        fields.append("userEnteredFormat.horizontalAlignment")
        applied.append(f"alignment {params.horizontal_align}")

    if params.wrap:
        cell_format["wrapStrategy"] = params.wrap.upper()
        fields.append("userEnteredFormat.wrapStrategy")
        applied.append(f"wrap {params.wrap}")

    if cell_format:
        requests.append({
            "repeatCell": {
                "range": grid_range,
                "cell": {"userEnteredFormat": cell_format},
                "fields": ",".join(fields),
            }
        })

    # Borders
    if params.borders:
        border_style = {"style": "SOLID", "width": 1, "colorStyle": {"rgbColor": {"red": 0, "green": 0, "blue": 0}}}
        no_border = {"style": "NONE"}

        if params.borders == "all":
            requests.append({
                "updateBorders": {
                    "range": grid_range,
                    "top": border_style, "bottom": border_style,
                    "left": border_style, "right": border_style,
                    "innerHorizontal": border_style, "innerVertical": border_style,
                }
            })
            applied.append("borders (all)")
        elif params.borders == "outline":
            requests.append({
                "updateBorders": {
                    "range": grid_range,
                    "top": border_style, "bottom": border_style,
                    "left": border_style, "right": border_style,
                }
            })
            applied.append("borders (outline)")
        elif params.borders == "none":
            requests.append({
                "updateBorders": {
                    "range": grid_range,
                    "top": no_border, "bottom": no_border,
                    "left": no_border, "right": no_border,
                    "innerHorizontal": no_border, "innerVertical": no_border,
                }
            })
            applied.append("borders removed")

    # Merge
    if params.merge is True:
        requests.append({"mergeCells": {"range": grid_range, "mergeType": "MERGE_ALL"}})
        applied.append("merged")
    elif params.merge is False:
        requests.append({"unmergeCells": {"range": grid_range}})
        applied.append("unmerged")

    if not requests:
        return "No formatting options specified. Set at least one formatting parameter."

    try:
        get_sheets().spreadsheets().batchUpdate(
            spreadsheetId=params.file_id,
            body={"requests": requests},
        ).execute()
        return f"Applied formatting to `{params.range}`: {', '.join(applied)}."
    except Exception as e:
        return f"Error applying formatting: {e}"


# ── Tool: Manage Sheets (tabs) ──────────────────────────────────────────────

class SheetAction(str, Enum):
    CREATE = "create"
    DUPLICATE = "duplicate"
    DELETE = "delete"
    RENAME = "rename"


class ManageSheetsInput(BaseModel):
    """Input for managing sheets/tabs in a spreadsheet."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(..., description="The Google Drive file ID of the spreadsheet.", min_length=1)
    action: SheetAction = Field(..., description="Action: 'create', 'duplicate', 'delete', or 'rename'.")
    sheet_name: Optional[str] = Field(default=None, description="Name of existing sheet (for duplicate/delete/rename).")
    new_name: Optional[str] = Field(default=None, description="New name (for create or rename).")


@mcp.tool(
    name="gdrive_manage_sheets",
    annotations={"title": "Manage Sheet Tabs", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
)
async def gdrive_manage_sheets(params: ManageSheetsInput) -> str:
    """Create, duplicate, delete, or rename sheets/tabs in a Google spreadsheet.

    Args:
        params: File ID, action, and relevant names.

    Returns:
        Confirmation of the action.
    """
    try:
        spreadsheet = get_sheets().spreadsheets().get(spreadsheetId=params.file_id).execute()
    except Exception as e:
        return f"Error opening spreadsheet: {e}"

    sheets = spreadsheet.get("sheets", [])
    sheet_map = {s["properties"]["title"]: s["properties"]["sheetId"] for s in sheets}

    requests = []

    if params.action == SheetAction.CREATE:
        name = params.new_name or "New Sheet"
        requests.append({"addSheet": {"properties": {"title": name}}})
    elif params.action == SheetAction.DUPLICATE:
        if not params.sheet_name or params.sheet_name not in sheet_map:
            return f"Sheet '{params.sheet_name}' not found. Available: {', '.join(sheet_map.keys())}"
        new_name = params.new_name or f"{params.sheet_name} (Copy)"
        requests.append({"duplicateSheet": {
            "sourceSheetId": sheet_map[params.sheet_name],
            "newSheetName": new_name,
        }})
    elif params.action == SheetAction.DELETE:
        if not params.sheet_name or params.sheet_name not in sheet_map:
            return f"Sheet '{params.sheet_name}' not found. Available: {', '.join(sheet_map.keys())}"
        if len(sheets) <= 1:
            return "Cannot delete the only sheet in a spreadsheet."
        requests.append({"deleteSheet": {"sheetId": sheet_map[params.sheet_name]}})
    elif params.action == SheetAction.RENAME:
        if not params.sheet_name or params.sheet_name not in sheet_map:
            return f"Sheet '{params.sheet_name}' not found. Available: {', '.join(sheet_map.keys())}"
        if not params.new_name:
            return "Error: new_name is required for rename."
        requests.append({"updateSheetProperties": {
            "properties": {"sheetId": sheet_map[params.sheet_name], "title": params.new_name},
            "fields": "title",
        }})

    try:
        get_sheets().spreadsheets().batchUpdate(
            spreadsheetId=params.file_id,
            body={"requests": requests},
        ).execute()
        action_desc = {
            SheetAction.CREATE: f"Created sheet '{params.new_name or 'New Sheet'}'",
            SheetAction.DUPLICATE: f"Duplicated '{params.sheet_name}' as '{params.new_name or params.sheet_name + ' (Copy)'}'",
            SheetAction.DELETE: f"Deleted sheet '{params.sheet_name}'",
            SheetAction.RENAME: f"Renamed '{params.sheet_name}' to '{params.new_name}'",
        }
        return f"{action_desc[params.action]}."
    except Exception as e:
        return f"Error: {e}"


# ── Tool: Create Spreadsheet ─────────────────────────────────────────────────

class CreateSheetInput(BaseModel):
    """Input for creating a new Google Sheet."""
    model_config = ConfigDict(str_strip_whitespace=True)

    title: str = Field(..., description="Title of the new spreadsheet.", min_length=1, max_length=500)
    sheet_names: Optional[List[str]] = Field(
        default=None,
        description="Names for sheets/tabs. If omitted, creates one default 'Sheet1'.",
    )
    headers: Optional[List[str]] = Field(
        default=None,
        description="Header row for the first sheet. Example: ['Name', 'Email', 'Score']",
    )
    folder_id: Optional[str] = Field(default=None, description="Folder ID to create in. Omit for Drive root.")


@mcp.tool(
    name="gdrive_create_sheet",
    annotations={"title": "Create Google Sheet", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
)
async def gdrive_create_sheet(params: CreateSheetInput) -> str:
    """Create a new Google Spreadsheet with optional tabs and header row.

    Args:
        params: Title, optional sheet names, headers, and folder.

    Returns:
        The new spreadsheet's ID and link.
    """
    body = {"properties": {"title": params.title}}
    if params.sheet_names:
        body["sheets"] = [{"properties": {"title": name}} for name in params.sheet_names]

    try:
        spreadsheet = get_sheets().spreadsheets().create(body=body).execute()
        file_id = spreadsheet["spreadsheetId"]
        link = spreadsheet.get("spreadsheetUrl", "")

        # Move to folder if specified
        if params.folder_id:
            get_drive().files().update(
                fileId=file_id,
                addParents=params.folder_id,
                removeParents="root",
                fields="id",
            ).execute()

        # Add headers if specified
        if params.headers:
            get_sheets().spreadsheets().values().update(
                spreadsheetId=file_id,
                range="A1",
                valueInputOption="USER_ENTERED",
                body={"values": [params.headers]},
            ).execute()

        return f"Created spreadsheet: **{params.title}**\n\nID: `{file_id}`\nLink: {link}"
    except Exception as e:
        return f"Error creating spreadsheet: {e}"


# ── Tool: Insert Rows/Columns ────────────────────────────────────────────────

class InsertRowsColsInput(BaseModel):
    """Input for inserting rows or columns in a Google Sheet."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(..., description="The Google Drive file ID of the spreadsheet.", min_length=1)
    sheet_name: Optional[str] = Field(default=None, description="Sheet/tab name. If omitted, uses first sheet.")
    dimension: str = Field(..., description="'ROWS' or 'COLUMNS'.")
    start_index: int = Field(..., description="0-based index where to start inserting.", ge=0)
    count: int = Field(default=1, description="Number of rows/columns to insert.", ge=1, le=1000)
    inherit_before: bool = Field(default=False, description="If true, inherit formatting from the row/column before the insertion point.")


@mcp.tool(
    name="gdrive_insert_rows_cols",
    annotations={"title": "Insert Rows/Columns in Sheet", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
)
async def gdrive_insert_rows_cols(params: InsertRowsColsInput) -> str:
    """Insert empty rows or columns into a Google Sheet at a specific position.

    Args:
        params: File ID, dimension (ROWS/COLUMNS), start index, and count.

    Returns:
        Confirmation of insertion.
    """
    try:
        spreadsheet = get_sheets().spreadsheets().get(spreadsheetId=params.file_id).execute()
        sheet_id = get_sheet_id(spreadsheet, params.sheet_name)
    except Exception as e:
        return f"Error: {e}"

    try:
        get_sheets().spreadsheets().batchUpdate(
            spreadsheetId=params.file_id,
            body={"requests": [{
                "insertDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": params.dimension.upper(),
                        "startIndex": params.start_index,
                        "endIndex": params.start_index + params.count,
                    },
                    "inheritFromBefore": params.inherit_before,
                }
            }]},
        ).execute()
        dim = "row(s)" if params.dimension.upper() == "ROWS" else "column(s)"
        return f"Inserted {params.count} {dim} at index {params.start_index}."
    except Exception as e:
        return f"Error inserting: {e}"


# ── Tool: Delete Rows/Columns ────────────────────────────────────────────────

class DeleteRowsColsInput(BaseModel):
    """Input for deleting rows or columns in a Google Sheet."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(..., description="The Google Drive file ID of the spreadsheet.", min_length=1)
    sheet_name: Optional[str] = Field(default=None, description="Sheet/tab name. If omitted, uses first sheet.")
    dimension: str = Field(..., description="'ROWS' or 'COLUMNS'.")
    start_index: int = Field(..., description="0-based index of first row/column to delete.", ge=0)
    count: int = Field(default=1, description="Number of rows/columns to delete.", ge=1, le=1000)


@mcp.tool(
    name="gdrive_delete_rows_cols",
    annotations={"title": "Delete Rows/Columns in Sheet", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": True},
)
async def gdrive_delete_rows_cols(params: DeleteRowsColsInput) -> str:
    """Delete rows or columns from a Google Sheet.

    Args:
        params: File ID, dimension (ROWS/COLUMNS), start index, and count.

    Returns:
        Confirmation of deletion.
    """
    try:
        spreadsheet = get_sheets().spreadsheets().get(spreadsheetId=params.file_id).execute()
        sheet_id = get_sheet_id(spreadsheet, params.sheet_name)
    except Exception as e:
        return f"Error: {e}"

    try:
        get_sheets().spreadsheets().batchUpdate(
            spreadsheetId=params.file_id,
            body={"requests": [{
                "deleteDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": params.dimension.upper(),
                        "startIndex": params.start_index,
                        "endIndex": params.start_index + params.count,
                    },
                }
            }]},
        ).execute()
        dim = "row(s)" if params.dimension.upper() == "ROWS" else "column(s)"
        return f"Deleted {params.count} {dim} starting at index {params.start_index}."
    except Exception as e:
        return f"Error deleting: {e}"


# ══════════════════════════════════════════════════════════════════════════════
# DRIVE TOOLS — File Management
# ══════════════════════════════════════════════════════════════════════════════

# ── Tool: Move / Copy ────────────────────────────────────────────────────────

class MoveCopyInput(BaseModel):
    """Input for moving or copying a file."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(..., description="The Google Drive file ID.", min_length=1)
    action: str = Field(..., description="'move' or 'copy'.")
    destination_folder_id: str = Field(..., description="Destination folder ID.", min_length=1)
    new_name: Optional[str] = Field(default=None, description="New name for the file (optional, mainly for copy).")


@mcp.tool(
    name="gdrive_move_copy",
    annotations={"title": "Move/Copy File", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
)
async def gdrive_move_copy(params: MoveCopyInput) -> str:
    """Move or copy a Google Drive file to a different folder.

    Args:
        params: File ID, action (move/copy), destination folder, optional new name.

    Returns:
        Confirmation with file ID and link.
    """
    try:
        if params.action == "copy":
            body = {}
            if params.new_name:
                body["name"] = params.new_name
            body["parents"] = [params.destination_folder_id]

            copied = get_drive().files().copy(
                fileId=params.file_id,
                body=body,
                fields="id, name, webViewLink",
                supportsAllDrives=True,
            ).execute()
            return f"Copied as **{copied['name']}**\n\nID: `{copied['id']}`\nLink: {copied.get('webViewLink', 'N/A')}"

        elif params.action == "move":
            # Get current parents
            file_meta = get_drive().files().get(
                fileId=params.file_id,
                fields="parents, name",
                supportsAllDrives=True,
            ).execute()
            current_parents = ",".join(file_meta.get("parents", []))

            update_body = {}
            if params.new_name:
                update_body["name"] = params.new_name

            result = get_drive().files().update(
                fileId=params.file_id,
                addParents=params.destination_folder_id,
                removeParents=current_parents,
                body=update_body,
                fields="id, name, webViewLink",
                supportsAllDrives=True,
            ).execute()
            return f"Moved **{result['name']}** to new folder.\n\nID: `{result['id']}`\nLink: {result.get('webViewLink', 'N/A')}"

        else:
            return "Error: action must be 'move' or 'copy'."
    except Exception as e:
        return f"Error: {e}"


# ── Tool: Share / Permissions ────────────────────────────────────────────────

class ShareInput(BaseModel):
    """Input for sharing a file."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(..., description="The Google Drive file ID.", min_length=1)
    email: Optional[str] = Field(default=None, description="Email address to share with. Omit for 'anyone with link'.")
    role: str = Field(
        default="reader",
        description="Permission role: 'reader' (view), 'commenter', 'writer' (edit).",
    )
    anyone_with_link: bool = Field(default=False, description="If true, makes the file accessible to anyone with the link.")
    notify: bool = Field(default=True, description="If true, sends notification email to the recipient.")
    message: Optional[str] = Field(default=None, description="Optional message to include in the notification email.")


@mcp.tool(
    name="gdrive_share",
    annotations={"title": "Share Google Drive File", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def gdrive_share(params: ShareInput) -> str:
    """Share a Google Drive file with specific people or make it link-accessible.

    Args:
        params: File ID, email or anyone_with_link, role, and notification options.

    Returns:
        Confirmation with sharing details.
    """
    try:
        if params.anyone_with_link:
            permission = {"type": "anyone", "role": params.role}
            get_drive().permissions().create(
                fileId=params.file_id,
                body=permission,
                fields="id",
                supportsAllDrives=True,
            ).execute()

            # Get the shareable link
            file_meta = get_drive().files().get(
                fileId=params.file_id, fields="webViewLink", supportsAllDrives=True
            ).execute()
            return f"File is now accessible to anyone with the link ({params.role}).\n\nLink: {file_meta.get('webViewLink', 'N/A')}"

        elif params.email:
            permission = {"type": "user", "role": params.role, "emailAddress": params.email}
            get_drive().permissions().create(
                fileId=params.file_id,
                body=permission,
                sendNotificationEmail=params.notify,
                emailMessage=params.message,
                fields="id",
                supportsAllDrives=True,
            ).execute()
            return f"Shared with {params.email} as {params.role}."

        else:
            return "Error: Provide either 'email' or set 'anyone_with_link=true'."
    except Exception as e:
        return f"Error sharing file: {e}"


# ── Tool: Export ─────────────────────────────────────────────────────────────

class ExportInput(BaseModel):
    """Input for exporting a Google file."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(..., description="The Google Drive file ID.", min_length=1)
    format: str = Field(
        default="pdf",
        description="Export format: 'pdf', 'docx', 'xlsx', 'csv', 'pptx', 'txt', 'html', 'tsv'.",
    )


EXPORT_MIME_MAP = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "csv": "text/csv",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "txt": "text/plain",
    "html": "text/html",
    "tsv": "text/tab-separated-values",
}


@mcp.tool(
    name="gdrive_export",
    annotations={"title": "Export Google Drive File", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def gdrive_export(params: ExportInput) -> str:
    """Export a Google Workspace file (Doc, Sheet, Slides) to another format.

    Downloads the exported content and returns it. For binary formats (pdf, docx, xlsx, pptx),
    saves to a temp file and returns the path. For text formats (csv, txt, html, tsv),
    returns the content directly.

    Args:
        params: File ID and export format.

    Returns:
        File content (text formats) or path to downloaded file (binary formats).
    """
    import tempfile
    import os

    fmt = params.format.lower()
    mime = EXPORT_MIME_MAP.get(fmt)
    if not mime:
        return f"Unsupported format '{fmt}'. Supported: {', '.join(EXPORT_MIME_MAP.keys())}"

    try:
        file_meta = get_drive().files().get(
            fileId=params.file_id, fields="name, mimeType", supportsAllDrives=True
        ).execute()
        name = file_meta.get("name", "export")

        content = get_drive().files().export(
            fileId=params.file_id, mimeType=mime
        ).execute()

        # Text formats — return content directly
        if fmt in ("csv", "txt", "html", "tsv"):
            text = content.decode("utf-8") if isinstance(content, bytes) else str(content)
            if len(text) > 50000:
                text = text[:50000] + "\n\n... (truncated at 50000 chars)"
            return f"## {name} (exported as {fmt})\n\n{text}"

        # Binary formats — save to temp file
        ext = fmt
        tmp_path = os.path.join(tempfile.gettempdir(), f"{name}.{ext}")
        with open(tmp_path, "wb") as f:
            f.write(content if isinstance(content, bytes) else content.encode("utf-8"))

        size_kb = os.path.getsize(tmp_path) / 1024
        return f"Exported **{name}** as {fmt} ({size_kb:.1f} KB)\n\nSaved to: `{tmp_path}`"

    except Exception as e:
        return f"Error exporting file: {e}"


# ── Tool: Comments ───────────────────────────────────────────────────────────

class CommentsInput(BaseModel):
    """Input for working with file comments."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(..., description="The Google Drive file ID.", min_length=1)
    action: str = Field(
        default="list",
        description="Action: 'list' (read comments), 'add' (add a comment), 'resolve' (resolve a comment).",
    )
    content: Optional[str] = Field(default=None, description="Comment text (required for 'add').")
    comment_id: Optional[str] = Field(default=None, description="Comment ID (required for 'resolve').")
    quoted_text: Optional[str] = Field(
        default=None,
        description="Text in the document to anchor the comment to (optional, for 'add').",
    )


@mcp.tool(
    name="gdrive_comments",
    annotations={"title": "Google Drive Comments", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
)
async def gdrive_comments(params: CommentsInput) -> str:
    """List, add, or resolve comments on a Google Drive file.

    Args:
        params: File ID, action, and relevant content/IDs.

    Returns:
        Comment listing or confirmation.
    """
    try:
        if params.action == "list":
            result = get_drive().comments().list(
                fileId=params.file_id,
                fields="comments(id, content, author(displayName), createdTime, resolved, quotedFileContent, replies(content, author(displayName), createdTime))",
                includeDeleted=False,
            ).execute()
            comments = result.get("comments", [])
            if not comments:
                return "No comments on this file."

            output = f"## Comments ({len(comments)})\n\n"
            for c in comments:
                status = " [RESOLVED]" if c.get("resolved") else ""
                author = c.get("author", {}).get("displayName", "Unknown")
                date = c.get("createdTime", "")[:10]
                quoted = c.get("quotedFileContent", {}).get("value", "")
                output += f"**{author}** ({date}){status} — ID: `{c['id']}`\n"
                if quoted:
                    output += f"> {quoted}\n"
                output += f"{c.get('content', '')}\n"

                # Show replies
                for r in c.get("replies", []):
                    r_author = r.get("author", {}).get("displayName", "Unknown")
                    r_date = r.get("createdTime", "")[:10]
                    output += f"  ↳ **{r_author}** ({r_date}): {r.get('content', '')}\n"
                output += "\n"

            return output

        elif params.action == "add":
            if not params.content:
                return "Error: 'content' is required for adding a comment."
            body = {"content": params.content}
            if params.quoted_text:
                body["quotedFileContent"] = {"value": params.quoted_text, "mimeType": "text/plain"}

            comment = get_drive().comments().create(
                fileId=params.file_id,
                body=body,
                fields="id, content, createdTime",
            ).execute()
            return f"Comment added (ID: `{comment['id']}`)."

        elif params.action == "resolve":
            if not params.comment_id:
                return "Error: 'comment_id' is required for resolving a comment."
            get_drive().comments().update(
                fileId=params.file_id,
                commentId=params.comment_id,
                body={"resolved": True},
                fields="id",
            ).execute()
            return f"Comment `{params.comment_id}` resolved."

        else:
            return "Error: action must be 'list', 'add', or 'resolve'."
    except Exception as e:
        return f"Error: {e}"


# ── Tool: Version History ────────────────────────────────────────────────────

class VersionsInput(BaseModel):
    """Input for viewing file version history."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(..., description="The Google Drive file ID.", min_length=1)
    max_results: int = Field(default=20, description="Maximum revisions to return.", ge=1, le=100)


@mcp.tool(
    name="gdrive_versions",
    annotations={"title": "File Version History", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def gdrive_versions(params: VersionsInput) -> str:
    """List revision history of a Google Drive file.

    Shows who modified the file, when, and revision IDs.

    Args:
        params: File ID and max results.

    Returns:
        List of revisions with dates and authors.
    """
    try:
        result = get_drive().revisions().list(
            fileId=params.file_id,
            fields="revisions(id, modifiedTime, lastModifyingUser(displayName, emailAddress))",
            pageSize=params.max_results,
        ).execute()

        revisions = result.get("revisions", [])
        if not revisions:
            return "No revision history available for this file."

        # Get file name
        file_meta = get_drive().files().get(fileId=params.file_id, fields="name", supportsAllDrives=True).execute()
        name = file_meta.get("name", "Unknown")

        output = f"## Revision History: {name}\n\n"
        output += f"Showing {len(revisions)} revision(s):\n\n"

        for i, rev in enumerate(reversed(revisions), 1):
            date = rev.get("modifiedTime", "?")[:19].replace("T", " ")
            user = rev.get("lastModifyingUser", {})
            author = user.get("displayName", user.get("emailAddress", "Unknown"))
            output += f"{i}. **{date}** — {author} (rev `{rev['id']}`)\n"

        return output
    except Exception as e:
        return f"Error fetching revisions: {e}"


# ══════════════════════════════════════════════════════════════════════════════
# SLIDES TOOLS
# ══════════════════════════════════════════════════════════════════════════════

# ── Tool: Read Slides ────────────────────────────────────────────────────────

class ReadSlidesInput(BaseModel):
    """Input for reading a Google Slides presentation."""
    model_config = ConfigDict(str_strip_whitespace=True)

    file_id: str = Field(..., description="The Google Drive file ID of the presentation.", min_length=1)
    slide_numbers: Optional[List[int]] = Field(
        default=None,
        description="Specific slide numbers to read (1-based). If omitted, reads all slides.",
    )


def extract_slide_text(slide: dict) -> str:
    """Extract text content from a slide's page elements."""
    texts = []
    for element in slide.get("pageElements", []):
        shape = element.get("shape", {})
        text_content = shape.get("text", {})
        for text_element in text_content.get("textElements", []):
            text_run = text_element.get("textRun", {})
            content = text_run.get("content", "")
            if content.strip():
                style = text_run.get("style", {})
                if style.get("bold"):
                    content = f"**{content.strip()}**"
                if style.get("italic"):
                    content = f"*{content.strip()}*"
                if style.get("link", {}).get("url"):
                    content = f"[{content.strip()}]({style['link']['url']})"
                texts.append(content)

        # Handle tables in slides
        table = element.get("table", {})
        if table:
            md_rows = []
            for row in table.get("tableRows", []):
                cells = []
                for cell in row.get("tableCells", []):
                    cell_text = ""
                    for tc in cell.get("text", {}).get("textElements", []):
                        cell_text += tc.get("textRun", {}).get("content", "")
                    cells.append(cell_text.strip().replace("|", "\\|"))
                md_rows.append("| " + " | ".join(cells) + " |")
            if md_rows:
                num_cols = len(table.get("tableRows", [{}])[0].get("tableCells", []))
                md_rows.insert(1, "| " + " | ".join(["---"] * num_cols) + " |")
                texts.append("\n".join(md_rows))

    return "\n".join(texts)


@mcp.tool(
    name="gdrive_read_slides",
    annotations={"title": "Read Google Slides", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def gdrive_read_slides(params: ReadSlidesInput) -> str:
    """Read a Google Slides presentation and return content as markdown.

    Extracts text, formatting, tables, and speaker notes from each slide.
    Optionally read specific slides by number.

    Args:
        params: File ID and optional slide numbers.

    Returns:
        Presentation content as markdown with slide separators.
    """
    try:
        presentation = get_slides().presentations().get(presentationId=params.file_id).execute()
    except Exception as e:
        return f"Error reading presentation: {e}"

    title = presentation.get("title", "Untitled Presentation")
    slides = presentation.get("slides", [])
    total = len(slides)

    if not slides:
        return f"Presentation '{title}' has no slides."

    output = f"# {title}\n\n*{total} slides*\n\n"

    for i, slide in enumerate(slides, 1):
        if params.slide_numbers and i not in params.slide_numbers:
            continue

        output += f"---\n\n## Slide {i}\n\n"

        # Slide content
        content = extract_slide_text(slide)
        if content.strip():
            output += content + "\n\n"
        else:
            output += "*[Empty slide]*\n\n"

        # Speaker notes
        notes_page = slide.get("slideProperties", {}).get("notesPage", {})
        for element in notes_page.get("pageElements", []):
            shape = element.get("shape", {})
            if shape.get("shapeType") == "TEXT_BOX":
                notes_text = ""
                for te in shape.get("text", {}).get("textElements", []):
                    notes_text += te.get("textRun", {}).get("content", "")
                if notes_text.strip():
                    output += f"**Speaker Notes:** {notes_text.strip()}\n\n"

    return output


# ══════════════════════════════════════════════════════════════════════════════
# APPS SCRIPT TOOLS
# ══════════════════════════════════════════════════════════════════════════════

# ── Tool: Run Script ─────────────────────────────────────────────────────────

class RunScriptInput(BaseModel):
    """Input for running an Apps Script function."""
    model_config = ConfigDict(str_strip_whitespace=True)

    script_id: str = Field(
        ...,
        description="The Apps Script project ID. Find this in the script editor URL or project settings.",
        min_length=1,
    )
    function_name: str = Field(
        ...,
        description="The function name to execute.",
        min_length=1,
    )
    parameters: Optional[List] = Field(
        default=None,
        description="Parameters to pass to the function as a JSON array.",
    )


@mcp.tool(
    name="gdrive_run_script",
    annotations={"title": "Run Apps Script", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
)
async def gdrive_run_script(params: RunScriptInput) -> str:
    """Execute a function in a Google Apps Script project.

    The script must be deployed as an API executable. Use this for advanced
    operations like complex formatting, chart creation, or custom automations
    that aren't possible through the Sheets/Docs APIs directly.

    Prerequisites:
    1. Open the script project in the Apps Script editor
    2. Deploy > New deployment > API Executable
    3. Use the script project ID (not the deployment ID)

    Args:
        params: Script project ID, function name, and optional parameters.

    Returns:
        The function's return value or error details.
    """
    body = {
        "function": params.function_name,
        "devMode": True,
    }
    if params.parameters:
        body["parameters"] = params.parameters

    try:
        response = get_scripts().scripts().run(
            scriptId=params.script_id,
            body=body,
        ).execute()

        if "error" in response:
            error = response["error"]
            details = error.get("details", [{}])
            error_msg = details[0].get("errorMessage", str(error)) if details else str(error)
            error_type = details[0].get("errorType", "UNKNOWN") if details else "UNKNOWN"
            return f"Script error ({error_type}): {error_msg}"

        result = response.get("response", {}).get("result")
        if result is None:
            return f"Function `{params.function_name}` executed successfully (no return value)."

        if isinstance(result, (dict, list)):
            return f"Function `{params.function_name}` returned:\n\n```json\n{json.dumps(result, indent=2)}\n```"

        return f"Function `{params.function_name}` returned: {result}"

    except Exception as e:
        error_str = str(e)
        if "not been deployed as an API Executable" in error_str:
            return (
                f"Error: Script is not deployed as an API Executable.\n\n"
                f"To fix:\n"
                f"1. Open: https://script.google.com/d/{params.script_id}/edit\n"
                f"2. Click Deploy > New deployment\n"
                f"3. Select 'API Executable'\n"
                f"4. Click Deploy\n"
                f"5. Try again"
            )
        return f"Error running script: {e}"


# ── Tool: Create Script ──────────────────────────────────────────────────────

class CreateScriptInput(BaseModel):
    """Input for creating an Apps Script project."""
    model_config = ConfigDict(str_strip_whitespace=True)

    title: str = Field(..., description="Name for the Apps Script project.", min_length=1)
    parent_id: Optional[str] = Field(
        default=None,
        description="File ID of a Google Doc/Sheet/Slides to bind the script to. If omitted, creates a standalone script.",
    )
    code: str = Field(
        ...,
        description="The JavaScript/Apps Script code to add to the project.",
        min_length=1,
    )


@mcp.tool(
    name="gdrive_create_script",
    annotations={"title": "Create Apps Script", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
)
async def gdrive_create_script(params: CreateScriptInput) -> str:
    """Create a new Google Apps Script project, optionally bound to a file.

    Creates the script project and pushes the provided code. The script
    can then be deployed and executed via gdrive_run_script.

    Args:
        params: Title, optional parent file ID, and the script code.

    Returns:
        Script project ID and link.
    """
    try:
        # Create project
        body = {"title": params.title}
        if params.parent_id:
            body["parentId"] = params.parent_id

        project = get_scripts().projects().create(body=body).execute()
        script_id = project["scriptId"]

        # Push the code
        content = {
            "files": [
                {
                    "name": "Code",
                    "type": "SERVER_JS",
                    "source": params.code,
                },
                {
                    "name": "appsscript",
                    "type": "JSON",
                    "source": json.dumps({
                        "timeZone": "America/New_York",
                        "dependencies": {},
                        "exceptionLogging": "STACKDRIVER",
                        "runtimeVersion": "V8",
                        "oauthScopes": [
                            "https://www.googleapis.com/auth/spreadsheets",
                            "https://www.googleapis.com/auth/documents",
                            "https://www.googleapis.com/auth/drive",
                        ],
                    }),
                },
            ]
        }

        get_scripts().projects().updateContent(
            scriptId=script_id,
            body=content,
        ).execute()

        link = f"https://script.google.com/d/{script_id}/edit"
        bound = f" (bound to file `{params.parent_id}`)" if params.parent_id else " (standalone)"

        return (
            f"Created Apps Script project: **{params.title}**{bound}\n\n"
            f"Script ID: `{script_id}`\n"
            f"Editor: {link}\n\n"
            f"To execute via API, deploy as 'API Executable' from the editor."
        )

    except Exception as e:
        return f"Error creating script: {e}"


# ── Run ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    mcp.run()
