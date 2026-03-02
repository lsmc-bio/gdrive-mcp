"""Drive file management tools for gdrive-mcp."""

import os
import tempfile
from typing import Optional

from pydantic import BaseModel, Field, ConfigDict

from services import get_drive


def register(mcp):
    """Register all file management tools with the MCP server."""

    # ── Tool: Move / Copy ────────────────────────────────────────────────

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

    # ── Tool: Share / Permissions ────────────────────────────────────────

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

    # ── Tool: Export ─────────────────────────────────────────────────────

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

    # ── Tool: Version History ────────────────────────────────────────────

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
