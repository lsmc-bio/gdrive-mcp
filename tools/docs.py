"""Google Docs tools for gdrive-mcp."""

import re
from typing import Optional, List
from enum import Enum

from pydantic import BaseModel, Field, ConfigDict
from googleapiclient.http import MediaInMemoryUpload

from services import get_docs, get_drive
from helpers import find_heading_end_index, find_heading_section_range, find_text_indices, hex_to_color


def register(mcp):
    """Register all Google Docs tools with the MCP server."""

    # ── Tool: Find & Replace ─────────────────────────────────────────────

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

    # ── Tool: Insert Text ────────────────────────────────────────────────

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
        Supports rich formatting: bold, italic, underline, strikethrough, font family/size,
        text color, background color, hyperlinks, and heading-level styling.

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

            return f"Inserted {len(params.text)} chars {location_desc[params.location]}{fmt_str}."

        except Exception as e:
            return f"Error inserting text: {e}"

    # ── Tool: Delete Text ────────────────────────────────────────────────

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

    # ── Tool: Create Doc ─────────────────────────────────────────────────

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

    # ── Tool: Read Section ───────────────────────────────────────────────

    class ReadSectionInput(BaseModel):
        """Input for reading a section of a Google Doc."""
        model_config = ConfigDict(str_strip_whitespace=True)

        file_id: str = Field(..., description="The Google Drive file ID.", min_length=1)
        heading: str = Field(
            ...,
            description="The heading text to read under (case-insensitive partial match). Returns all content from this heading until the next heading of equal or higher level.",
            min_length=1,
        )

    # We need doc_element_to_markdown from drive module — define it locally
    # since it's used both for read_doc and read_section
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
                nesting = bullet.get("nestingLevel", 0)
                indent = "  " * nesting
                list_id = bullet.get("listId", "")

                if list_id not in lists_state:
                    lists_state[list_id] = {}
                if nesting not in lists_state[list_id]:
                    lists_state[list_id][nesting] = 0
                lists_state[list_id][nesting] += 1

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

                if len(md_rows) > 0:
                    num_cols = len(rows[0].get("tableCells", []))
                    separator = "| " + " | ".join(["---"] * num_cols) + " |"
                    md_rows.insert(1, separator)

                output = "\n".join(md_rows) + "\n\n"

        return output

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

    # ── Tool: Insert Table ───────────────────────────────────────────────

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

    # ── Tool: Format Text ────────────────────────────────────────────────

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
