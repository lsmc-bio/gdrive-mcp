# gdrive-mcp

A local MCP (Model Context Protocol) server that gives Claude full read-write access to Google Drive, Docs, Sheets, Slides, Gmail (read-only), Calendar, and Apps Script.

## Features

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

This opens a browser window for Google sign-in. After authorizing, tokens are saved to `token.json` (automatically gitignored).

### 4. Configure Claude Code

Add the server to your MCP configuration (`.mcp.json`):

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

Or run standalone:

```bash
python server.py
```

## OAuth Scopes

The server requests the following scopes:

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

## Security Notes

- `credentials.json` and `token.json` are gitignored and must never be committed
- `token.json` is created with `0600` permissions (owner-only read/write)
- Gmail access is read-only by design
- All Google API calls use the authenticated user's permissions -- the server cannot access files the user doesn't have access to

## License

MIT
