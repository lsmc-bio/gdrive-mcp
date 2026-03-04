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

## Quick Start with Claude Code

If you already have your `credentials.json` file ready, you can paste something like this into Claude Code and it will handle the rest:

> Clone and set up the gdrive-mcp server from https://github.com/lsmc-bio/gdrive-mcp. My Google OAuth credentials file is at ~/Downloads/credentials.json. Install it globally (or: for this project).

Claude Code will clone the repo, move your credentials into place, install dependencies, run the auth flow, and add the MCP server to your config. You shouldn't need to touch any config files yourself.

## Setup (Manual)

If you prefer to set things up by hand, or if you want to understand what each step does, follow the instructions below.

### Step 1: Create a Google Cloud project

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Click the project dropdown at the top of the page, then **New Project**
3. Name it something recognizable (e.g. `gdrive-mcp` or `workspace-tools`) -- the name is just for your reference
4. Click **Create** and wait for it to provision, then select it from the project dropdown

### Step 2: Enable the APIs

1. Go to **APIs & Services > Library** (or search "API Library" in the console search bar)
2. Search for and enable each of the following. Click the API name, then click **Enable**:
   - Google Drive API
   - Google Docs API
   - Google Sheets API
   - Google Slides API
   - Gmail API
   - Google Calendar API
   - Apps Script API

### Step 3: Configure the OAuth consent screen

Before you can create credentials, Google requires you to configure a consent screen. This is what users see when they authorize the app.

1. Go to **APIs & Services > OAuth consent screen**
2. Choose a user type:
   - **Internal** (recommended if you're on Google Workspace): Only users in your organization can authorize. No app review needed. Tokens don't expire after 7 days.
   - **External**: Anyone with a Google account can authorize, but the app starts in "Testing" mode with limitations (see below).
3. Fill in the required fields:
   - **App name**: Anything descriptive (e.g. `gdrive-mcp`)
   - **User support email**: Your email
   - **Developer contact email**: Your email
4. Click **Save and Continue** through the Scopes and Test Users pages (you don't need to add scopes here -- they're requested at runtime)
5. Click **Back to Dashboard**

**Important -- Testing mode vs. published apps:**

If you chose **External**, your app starts in "Testing" mode. This works fine, but has one limitation: **refresh tokens expire after 7 days**, which means you'll need to re-run `python auth.py` weekly. To make tokens permanent, you have two options:

- **Publish the app.** Click **Publish App** on the consent screen page. For personal or small-team use, Google won't require verification -- you'll just see an "unverified app" warning during sign-in, which you can click through. Once published, tokens persist indefinitely.
- **Switch to Internal** (Workspace orgs only). Internal apps have no token expiration and no verification requirement.

If you chose **Internal**, you're already set -- tokens won't expire.

### Step 4: Create OAuth credentials

1. Go to **APIs & Services > Credentials**
2. Click **Create Credentials > OAuth client ID**
3. For **Application type**, select **Desktop app**
4. Name it anything (e.g. `gdrive-mcp`)
5. Click **Create**
6. Click **Download JSON** on the confirmation dialog (or find it in the credentials list and click the download icon)
7. Save the file as `credentials.json` in the gdrive-mcp project root

### Step 5: Clone and install

```bash
git clone https://github.com/lsmc-bio/gdrive-mcp.git
cd gdrive-mcp

# Move your credentials into place
cp ~/Downloads/client_secret_*.json credentials.json

# Install dependencies (using uv, recommended)
uv sync

# Or using pip
pip install -e .
```

### Step 6: Authenticate

```bash
python auth.py
```

This opens a browser window. Sign in with your Google account and grant the requested permissions. If you see an "unverified app" warning, click **Advanced > Go to [app name]** to proceed.

After authorizing, tokens are saved to `token.json` (automatically gitignored). The token refreshes automatically on each use. If you configured your consent screen as Internal or published, you only need to do this once.

### Step 7: Configure your MCP client

Add the server to your MCP configuration. For **Claude Code**, add to your `.mcp.json` (project-level or `~/.claude/.mcp.json` for global):

```json
{
  "mcpServers": {
    "gdrive": {
      "command": "uv",
      "args": ["run", "--directory", "/absolute/path/to/gdrive-mcp", "python", "server.py"]
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
