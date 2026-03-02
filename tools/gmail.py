"""Gmail tools for gdrive-mcp (read-only)."""

import base64
import email
from typing import Optional, List
from enum import Enum

from pydantic import BaseModel, Field, ConfigDict

from services import get_gmail


# ── Gmail helper functions (module-level) ────────────────────────────────────


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


# ── Tool registration ────────────────────────────────────────────────────────


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
