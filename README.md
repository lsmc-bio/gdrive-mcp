# gdrive-mcp

A complete Google Workspace MCP server. One server, one OAuth flow, full read-write access to Drive, Docs, Sheets, Slides, Gmail, Calendar, and Apps Script.

## Why this exists

Coding agents like Claude Code are increasingly capable of doing real work -- not just writing code, but managing projects, drafting documents, updating spreadsheets, scheduling meetings, and triaging email. The missing piece is access to the tools where that work actually lives. For most teams, that's Google Workspace.

The MCP (Model Context Protocol) ecosystem has several Google integrations, but in practice they tend to be incomplete or frustrating to use. Some only cover Drive search and Docs reading. Others spread Google functionality across multiple servers that each require separate auth flows. Many lack write access entirely, or cover Docs but not Sheets, or Sheets but not Slides. The result is that you end up wiring together multiple partial solutions, each with its own OAuth config, and you still can't do half of what you need.

gdrive-mcp takes a different approach: **one server that covers all of Google Workspace with real depth**. Not just "read a doc" but insert formatted text at a specific heading, manage sheet tabs, create slides with positioned text boxes, search Gmail threads, schedule calendar events with Google Meet links, and deploy Apps Script projects. The goal is that an agent using this server can do anything a human would do in Google Workspace, through the same APIs, with the same permissions.

## What this unlocks

When you give a coding agent access to Google Workspace, the workflows that become possible go well beyond "read me that doc":

- **Project management from the terminal.** Ask your agent to check your calendar for the day, pull up the relevant project doc, update the status spreadsheet, and draft a summary -- all without leaving your editor.
- **Document creation and editing.** Generate formatted Google Docs from code, data, or conversation. Insert tables, apply heading styles, add comments. Create slide decks with structured content and speaker notes.
- **Spreadsheet automation.** Read data from sheets, write results back, format cells, manage tabs -- useful for anything from tracking deployments to updating shared team data.
- **Email triage.** Search and read Gmail (read-only by design) to pull context into your workflow. Find that thread from last week, read the attachment list, batch-read a set of messages.
- **Calendar integration.** Check availability, create events with attendees and Meet links, update or cancel meetings. Useful for agents that coordinate across people or schedule follow-ups.
- **Apps Script as an escape hatch.** For anything the APIs don't cover directly, create and run Apps Script functions. Build custom automations, complex formatting, or chart generation -- and manage deployments programmatically.

The common thread is that the agent operates on your real documents, in your real Drive, with your real permissions. Nothing is sandboxed or simulated.

## Tools

44 tools across 9 Google Workspace domains:

| Domain | Tools | Access |
|---|---|---|
| **Drive** | Search, list folder, recent files, file info | Read |
| **Docs** | Read, create, find/replace, insert/delete text, format text, insert table, read section | Read/Write |
| **Sheets** | Read, write, append, format cells, manage tabs, insert/delete rows/cols, create | Read/Write |
| **Slides** | Read, get elements, replace text, set element text, add text box, speaker notes, create/duplicate/delete/reorder slides | Read/Write |
| **Gmail** | Search, read, read thread, batch read, list labels | Read-only |
| **Calendar** | List events, get event, create/update/delete event, free/busy, list calendars | Read/Write |
| **Comments** | List, add, resolve comments on any Drive file | Read/Write |
| **File Management** | Move/copy, share/permissions, export, version history | Read/Write |
| **Apps Script** | Run, create, list, get, update, deploy script projects | Read/Write |

## Prerequisites

- Python 3.10+
- A Google Cloud project with the following APIs enabled:
  - Google Drive API
  - Google Docs API
  - Google Sheets API
  - Google Slides API
  - Gmail API
  - Google Calendar API
  - Apps Script API

## Setup

### 1. Google Cloud credentials

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project (or select an existing one)
3. Enable the APIs listed above (APIs & Services > Library)
4. Create OAuth 2.0 credentials:
   - Go to APIs & Services > Credentials
   - Click "Create Credentials" > "OAuth client ID"
   - Select "Desktop app" as the application type
   - Download the JSON file
5. Save the downloaded file as `credentials.json` in the project root

### 2. Install dependencies

```bash
# Using uv (recommended)
uv sync

# Or using pip
pip install -e .
```

### 3. Authenticate

```bash
python auth.py
```

This opens a browser window for Google sign-in. After authorizing, tokens are saved to `token.json` (automatically gitignored). The token refreshes automatically -- you only need to do this once.

### 4. Configure your MCP client

Add the server to your MCP configuration. For Claude Code, add to `.mcp.json`:

```json
{
  "mcpServers": {
    "gdrive": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/gdrive-mcp", "python", "server.py"]
    }
  }
}
```

This works with any MCP-compatible client -- Claude Code, Claude Desktop, Cursor, Windsurf, or any agent framework that speaks MCP.

To run standalone for testing:

```bash
python server.py
```

## OAuth Scopes

| Scope | Purpose |
|---|---|
| `drive` | Full Drive access (search, read, write, share) |
| `documents` | Google Docs read/write |
| `spreadsheets` | Google Sheets read/write |
| `presentations` | Google Slides read/write |
| `gmail.readonly` | Gmail search and read (no send/delete) |
| `calendar` | Google Calendar full access |
| `script.projects` | Apps Script project management |
| `script.deployments` | Apps Script deployment management |

Gmail is intentionally read-only. The server can search and read your email to pull context into workflows, but it cannot send, delete, or modify messages.

## Project Structure

```
gdrive-mcp/
├── server.py          # MCP server entry point
├── auth.py            # OAuth2 authentication
├── services.py        # Lazy-initialized Google API clients
├── helpers.py         # Shared utilities
├── pyproject.toml     # Project config and dependencies
└── tools/
    ├── __init__.py    # Tool registration
    ├── drive.py       # Drive search and read tools
    ├── docs.py        # Google Docs tools
    ├── sheets.py      # Google Sheets tools
    ├── slides.py      # Google Slides tools
    ├── gmail.py       # Gmail tools (read-only)
    ├── calendar.py    # Google Calendar tools
    ├── comments.py    # Drive comments tools
    ├── management.py  # File management (move, share, export)
    └── scripts.py     # Apps Script tools
```

## Security

- `credentials.json` and `token.json` are gitignored and must never be committed
- `token.json` is created with `0600` permissions (owner-only read/write)
- Gmail access is read-only by design
- All Google API calls use the authenticated user's permissions -- the server cannot access files the user doesn't have access to
- Drive API query inputs are escaped to prevent query injection

## License

MIT
