"""Google Sheets tools for gdrive-mcp."""

from typing import Optional, List
from enum import Enum

from pydantic import BaseModel, Field, ConfigDict

from services import get_sheets, get_drive
from helpers import hex_to_color, a1_to_grid_range, get_sheet_id


def register(mcp):
    """Register all Google Sheets tools with the MCP server."""

    # ── Tool: Write Sheet ────────────────────────────────────────────────

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

    # ── Tool: Append Sheet ───────────────────────────────────────────────

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

    # ── Tool: Format Cells ───────────────────────────────────────────────

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

    # ── Tool: Manage Sheets (tabs) ───────────────────────────────────────

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

    # ── Tool: Create Spreadsheet ─────────────────────────────────────────

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

    # ── Tool: Insert Rows/Columns ────────────────────────────────────────

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

    # ── Tool: Delete Rows/Columns ────────────────────────────────────────

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
