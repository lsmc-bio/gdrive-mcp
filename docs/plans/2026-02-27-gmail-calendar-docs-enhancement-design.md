# gdrive-mcp Enhancement: Gmail, Calendar, Docs Formatting, Apps Script

**Date:** 2026-02-27
**Status:** Approved

## Overview

Extend the gdrive-mcp server (currently 27 tools) with Gmail read-only integration, full Calendar read/write, enhanced Docs text formatting, and complete Apps Script management. Target: 46 tools total, organized into domain-specific modules.

## Architecture

### Project Structure

```
gdrive-mcp/
├── server.py              # Thin coordinator — creates FastMCP, imports tool modules
├── auth.py                # OAuth (add 3 new scopes)
├── services.py            # Extracted service getters (get_drive, get_docs, get_gmail, etc.)
├── helpers.py             # Shared utilities (hex_to_color, a1_to_grid_range, etc.)
├── tools/
│   ├── __init__.py
│   ├── drive.py           # search, list_folder, recent, file_info, move_copy, share, export, versions
│   ├── docs.py            # read_doc, read_section, create_doc, insert_text (enhanced), delete_text, find_replace, insert_table, format_text (NEW)
│   ├── sheets.py          # read_sheet, create_sheet, write_sheet, append_sheet, format_cells, manage_sheets, insert/delete rows_cols
│   ├── slides.py          # read_slides
│   ├── comments.py        # comments
│   ├── gmail.py           # search, read, read_thread, read_batch, labels (ALL NEW)
│   ├── calendar.py        # list_events, get_event, create_event, update_event, delete_event, free_busy, list_calendars (ALL NEW)
│   └── scripts.py         # create_script, run_script (existing), list_scripts, get_script, update_script, deploy_script (NEW)
├── credentials.json
├── token.json
└── pyproject.toml
```

### New OAuth Scopes

Add to `auth.py` SCOPES list:
- `https://www.googleapis.com/auth/gmail.readonly`
- `https://www.googleapis.com/auth/calendar`
- `https://www.googleapis.com/auth/script.deployments`

Users must re-run `python auth.py` once to authorize the new scopes. Existing `token.json` will be replaced.

### New Service Getters

Add to `services.py`:
- `get_gmail()` → `build("gmail", "v1", ...)`
- `get_calendar()` → `build("calendar", "v3", ...)`

Existing getters (`get_drive`, `get_docs`, `get_sheets`, `get_slides`, `get_scripts`) move from `server.py` to `services.py`.

---

## Gmail Tools (Read-Only) — 5 tools

### gmail_search

Smart search with optional inline body previews.

**Input:**
- `query: str` — Full Gmail search syntax (from:, to:, subject:, is:, label:, has:, before:, after:)
- `max_results: int` — Default 20, max 100
- `include_body: bool` — Default False. When True, includes body preview per result.
- `body_length: int` — Default 500. How many chars of body to include when include_body=True.

**Output:** Markdown-formatted list with subject, from, to, date, snippet, labels, thread_id, message_id, and optional body preview.

**Improvement over Google Workspace:** No second call needed for triage. Search returns actionable content.

### gmail_read

Read one email with granularity control.

**Input:**
- `message_id: str`
- `format: str` — "full" (default), "metadata" (headers only), "summary" (first N chars)
- `max_body_length: Optional[int]` — Truncate body. Avoids dumping 50KB newsletters.
- `include_headers: bool` — Default True
- `include_attachments: bool` — Default False. Lists attachment names/sizes (not content).

**Improvement over Google Workspace:** `max_body_length` and `format: "summary"` prevent overwhelming context with large emails.

### gmail_read_thread

Entire conversation in one call.

**Input:**
- `thread_id: str`
- `max_messages: int` — Default 10
- `max_body_length: Optional[int]` — Per-message truncation
- `newest_first: bool` — Default True

**Improvement over Google Workspace:** They have no thread-level tool at all.

### gmail_read_batch

Multiple emails at once.

**Input:**
- `message_ids: list[str]` — Up to 50
- `format: str` — "full", "metadata", "summary"
- `max_body_length: Optional[int]`

**Improvement over Google Workspace:** 50 vs their 25 cap, with body truncation controls.

### gmail_labels

List all labels with message counts.

**Input:** None.

**Output:** All user labels with name, ID, message count, unread count.

**Improvement over Google Workspace:** They don't expose labels at all.

---

## Calendar Tools (Full Read/Write) — 7 tools

### gcal_list_events

Flexible event listing with response status filtering.

**Input:**
- `calendar_id: str` — Default "primary"
- `time_min: Optional[str]` — RFC3339 or date string
- `time_max: Optional[str]`
- `query: Optional[str]` — Keyword search
- `max_results: int` — Default 25, max 250
- `include_recurring: bool` — Default True (expand recurring into instances)
- `detailed: bool` — Default False (when True: attendees, description, attachments)
- `status_filter: Optional[str]` — "accepted", "tentative", "declined", "needsAction"

**Improvement over Google Workspace:** `status_filter` for "what haven't I responded to?", `include_recurring` toggle.

### gcal_get_event

Single event full detail.

**Input:**
- `event_id: str`
- `calendar_id: str` — Default "primary"

**Output:** Title, time, location, description, attendees with response status, recurrence rules, Meet link, attachments, reminders, organizer.

### gcal_create_event

Full-featured event creation with recurrence.

**Input:**
- `summary: str`
- `start: str`, `end: str` — RFC3339 or date string
- `calendar_id: str` — Default "primary"
- `timezone: Optional[str]` — IANA format, defaults from calendar
- `description: Optional[str]`
- `location: Optional[str]`
- `attendees: Optional[list[str]]` — Just email strings
- `recurrence: Optional[str]` — "daily", "weekly", "monthly", "yearly", or raw RRULE
- `recurrence_count: Optional[int]` — Repeat N times
- `recurrence_until: Optional[str]` — Repeat until date
- `reminders: Optional[list[dict]]` — [{"method": "popup", "minutes": 15}]
- `add_google_meet: bool` — Default False
- `visibility: Optional[str]` — "default", "public", "private"
- `transparency: Optional[str]` — "opaque" (busy), "transparent" (free)
- `guests_can_modify: Optional[bool]`
- `send_notifications: str` — "all" (default), "none", "external_only"

**Improvement over Google Workspace:** Recurrence is first-class (simple strings auto-convert to RRULE). Attendees are plain email strings.

### gcal_update_event

Partial updates with add/remove attendees.

**Input:**
- `event_id: str`
- `calendar_id: str` — Default "primary"
- All optional: `summary`, `start`, `end`, `timezone`, `description`, `location`, `recurrence`, `reminders`, `add_google_meet`, `visibility`, `transparency`
- `add_attendees: Optional[list[str]]` — Emails to ADD
- `remove_attendees: Optional[list[str]]` — Emails to REMOVE
- `send_notifications: str` — Default "all"

**Improvement over Google Workspace:** `add_attendees`/`remove_attendees` instead of requiring the full attendee list. Server fetches, merges, submits internally.

### gcal_delete_event

Missing entirely from Google Workspace.

**Input:**
- `event_id: str`
- `calendar_id: str` — Default "primary"
- `send_notifications: str` — Default "all"
- `scope: str` — "this" (single occurrence), "following" (this and future), "all" (entire series)

**Improvement over Google Workspace:** They literally don't have event deletion. The `scope` field handles recurring event deletion cleanly.

### gcal_free_busy

Check availability before creating events.

**Input:**
- `calendars: list[str]` — Calendar IDs, default ["primary"]
- `time_min: str`
- `time_max: str`
- `timezone: Optional[str]`

**Output:** Busy blocks per calendar.

### gcal_list_calendars

List available calendars.

**Input:** None.

**Output:** All calendars with ID, name, primary status, access role, timezone.

---

## Enhanced Docs Formatting — 1 new tool + 1 modified

### Modified: gdrive_insert_text

Add new optional fields to existing `InsertTextInput`:
- `underline: bool` — Default False
- `strikethrough: bool` — Default False
- `font_family: Optional[str]` — "Arial", "Times New Roman", etc.
- `font_size: Optional[int]` — Points (1-400)
- `text_color: Optional[str]` — Hex "#FF0000"
- `bg_color: Optional[str]` — Highlight color hex
- `link_url: Optional[str]` — Make inserted text a hyperlink

Implementation: expand the existing `text_style` dict and `fields` mask in the batchUpdate call. Minimal code change.

### New: gdrive_format_text

Restyle existing document content with three targeting modes.

**Input:**

Target selection (one of three modes):
- `target: str` — "match", "range", or "heading"
- For "match": `match_text: str`, `match_occurrence: Optional[int]` (default all, 1=first, -1=last)
- For "range": `start_index: int`, `end_index: int`
- For "heading": `heading_text: str` (partial match, case-insensitive)

Text formatting (all optional):
- `bold`, `italic`, `underline`, `strikethrough` — Optional[bool]
- `font_family: Optional[str]`
- `font_size: Optional[int]`
- `text_color: Optional[str]` — Hex
- `bg_color: Optional[str]` — Hex
- `link_url: Optional[str]`
- `remove_link: Optional[bool]`

Paragraph formatting (all optional, applied to whole paragraphs in range):
- `named_style: Optional[str]` — "HEADING_1" through "HEADING_6", "NORMAL_TEXT", "TITLE", "SUBTITLE"
- `alignment: Optional[str]` — "START", "CENTER", "END", "JUSTIFIED"
- `line_spacing: Optional[float]` — 1.0 = single, 1.5, 2.0
- `indent_start: Optional[float]` — Points from left margin

**Improvements over Google Workspace `modify_doc_text`:**

| Google Workspace | gdrive_format_text |
|---|---|
| Character index only | Match by text, heading, or index |
| No paragraph formatting | Heading styles, alignment, spacing, indentation |
| No "find and format" | match_text + match_occurrence |
| Single range per call | Match mode hits all occurrences |
| No link removal | Explicit remove_link |

**Implementation:** For match mode, fetch doc → scan for text → convert to index ranges → build batchUpdate. For heading mode, reuse existing `find_heading_end_index` helper. The Docs API batchUpdate already supports all formatting ops.

---

## Enhanced Apps Script — 4 new tools

### gdrive_list_scripts

**Input:**
- `max_results: int` — Default 20
- `bound_to: Optional[str]` — File ID to filter scripts bound to a specific doc/sheet
- `query: Optional[str]` — Search by name

**Improvement over Google Workspace:** Filtering by bound document. Theirs is flat pagination only.

### gdrive_get_script

**Input:**
- `script_id: str`
- `file_name: Optional[str]` — One file or all
- `include_manifest: bool` — Default True (appsscript.json)

**Improvement over Google Workspace:** Combined single-file and whole-project read. Manifest included by default.

### gdrive_update_script

**Input:**
- `script_id: str`
- `files: list[{name, source, type?}]` — Type auto-detected from extension
- `merge: bool` — Default True (update only specified files, keep the rest)

**Improvement over Google Workspace:** Merge mode prevents accidental file deletion. Auto-detected file types from extension.

### gdrive_deploy_script

**Input:**
- `script_id: str`
- `action: str` — "create", "list", "update", "delete"
- `deployment_id: Optional[str]` — For update/delete
- `description: Optional[str]`
- `type: str` — "API_EXECUTABLE" (default), "ADDON", "WEBAPP"
- `version: Optional[int]` — Specific version or HEAD

**Improvement over Google Workspace:** They can only generate trigger code as a string. We manage real deployments.

---

## Tool Inventory Summary

| Domain | Existing | New | Modified | Total |
|---|---|---|---|---|
| Drive | 8 | 0 | 0 | 8 |
| Docs | 6 | 1 (format_text) | 1 (insert_text) | 7 |
| Sheets | 7 | 0 | 0 | 7 |
| Slides | 1 | 0 | 0 | 1 |
| Comments | 1 | 0 | 0 | 1 |
| Gmail | 0 | 5 | 0 | 5 |
| Calendar | 0 | 7 | 0 | 7 |
| Apps Script | 2 | 4 | 0 | 6 |
| **Total** | **25** | **17** | **1** | **42** |

Note: 42 unique tool functions. Some were double-counted in the original 27 estimate due to the refactor consolidation.

## Implementation Order

1. **Phase 1 — Refactor:** Extract services.py, helpers.py, split tools/ modules. Zero new functionality, just reorganization. Verify all 27 existing tools still work.
2. **Phase 2 — Docs enhancement:** Extend insert_text, add format_text. Smallest change, immediate value.
3. **Phase 3 — Gmail:** Add all 5 gmail tools. New scope, new service getter.
4. **Phase 4 — Calendar:** Add all 7 calendar tools. New scope, new service getter.
5. **Phase 5 — Apps Script:** Add 4 new script tools. New scope for deployments.
