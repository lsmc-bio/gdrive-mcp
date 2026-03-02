"""Drive read tools for gdrive-mcp."""

import re
from datetime import datetime, timedelta, timezone
from typing import Optional, List
from enum import Enum

from pydantic import BaseModel, Field, ConfigDict

from services import get_drive, get_docs, get_sheets
from helpers import format_file_entry, drive_query_files, GOOGLE_MIME_LABELS


def register(mcp):
    """Register all drive read tools with the MCP server."""

    # ── Shared enums/dicts for this module ───────────────────────────────

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

    # ── Tool: Search ─────────────────────────────────────────────────────

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
            parts.append(f"fullText contains '{query_text.replace(chr(39), chr(92)+chr(39))}'")
        else:
            parts.append(f"name contains '{query_text.replace(chr(39), chr(92)+chr(39))}'")


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

    # ── Tool: Read Google Doc ────────────────────────────────────────────

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

    # ── Tool: Read Google Sheet ──────────────────────────────────────────

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

    # ── Tool: List Folder ────────────────────────────────────────────────

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

    # ── Tool: Recent Files ───────────────────────────────────────────────

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

    # ── Tool: File Info ──────────────────────────────────────────────────

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
