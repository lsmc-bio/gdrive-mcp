"""Google Drive comments tools for gdrive-mcp."""

from typing import Optional

from pydantic import BaseModel, Field, ConfigDict

from services import get_drive


def register(mcp):
    """Register all comments tools with the MCP server."""

    # ── Tool: Comments ───────────────────────────────────────────────────

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
