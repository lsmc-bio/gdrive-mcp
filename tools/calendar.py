"""Google Calendar tools for gdrive-mcp."""

from datetime import datetime, timezone
from typing import Optional, List

from pydantic import BaseModel, Field, ConfigDict

from services import get_calendar


def _parse_datetime(dt_str: str) -> str:
    """Parse a flexible datetime string to RFC3339.

    Accepts: RFC3339 ('2025-03-15T14:00:00Z'), date ('2025-03-15'),
    or date with simple time ('2025-03-15 2:30 PM' — treated as UTC).
    Returns RFC3339 string.
    """
    if not dt_str:
        return dt_str

    # Already RFC3339
    if "T" in dt_str:
        return dt_str

    # Just a date
    if len(dt_str) == 10:
        return dt_str

    # Try parsing "YYYY-MM-DD HH:MM" or "YYYY-MM-DD H:MM PM"
    for fmt in ["%Y-%m-%d %I:%M %p", "%Y-%m-%d %H:%M", "%Y-%m-%d %I:%M%p", "%Y-%m-%d %H:%M:%S"]:
        try:
            dt = datetime.strptime(dt_str.strip(), fmt)
            return dt.strftime("%Y-%m-%dT%H:%M:%S")
        except ValueError:
            continue

    return dt_str  # Return as-is; API will validate


def _build_recurrence(recurrence: str, count: Optional[int], until: Optional[str]) -> list:
    """Build RRULE list from simple string or raw RRULE."""
    if recurrence.startswith("RRULE:"):
        return [recurrence]

    freq_map = {"daily": "DAILY", "weekly": "WEEKLY", "monthly": "MONTHLY", "yearly": "YEARLY"}
    freq = freq_map.get(recurrence.lower())
    if not freq:
        return [recurrence]  # Assume raw RRULE without prefix

    rule = f"RRULE:FREQ={freq}"
    if count:
        rule += f";COUNT={count}"
    if until:
        until_clean = until.replace("-", "")
        if len(until_clean) == 8:
            until_clean += "T235959Z"
        rule += f";UNTIL={until_clean}"

    return [rule]


def _format_event(event: dict, detailed: bool = False) -> str:
    """Format a calendar event as markdown."""
    summary = event.get("summary", "(no title)")
    start = event.get("start", {})
    end = event.get("end", {})
    start_str = start.get("dateTime", start.get("date", "?"))
    end_str = end.get("dateTime", end.get("date", "?"))

    output = f"**{summary}**\n"
    output += f"  When: {start_str} → {end_str}\n"
    output += f"  Event ID: `{event.get('id', '?')}`\n"

    if event.get("location"):
        output += f"  Location: {event['location']}\n"
    if event.get("htmlLink"):
        output += f"  Link: {event['htmlLink']}\n"

    status = event.get("status", "")
    if status and status != "confirmed":
        output += f"  Status: {status}\n"

    if event.get("recurringEventId"):
        output += f"  Recurring (series ID: `{event['recurringEventId']}`)\n"

    organizer = event.get("organizer", {})
    if organizer:
        output += f"  Organizer: {organizer.get('displayName', organizer.get('email', '?'))}\n"

    if detailed:
        if event.get("description"):
            output += f"  Description: {event['description']}\n"
        attendees = event.get("attendees", [])
        if attendees:
            output += f"  Attendees ({len(attendees)}):\n"
            for a in attendees:
                name = a.get("displayName", a.get("email", "?"))
                resp = a.get("responseStatus", "?")
                opt = " (optional)" if a.get("optional") else ""
                output += f"    - {name} [{resp}]{opt}\n"

        hangout = event.get("hangoutLink") or event.get("conferenceData", {}).get("entryPoints", [{}])[0].get("uri")
        if hangout:
            output += f"  Meet: {hangout}\n"

        if event.get("recurrence"):
            for rule in event["recurrence"]:
                output += f"  Recurrence: {rule}\n"

        reminders = event.get("reminders", {})
        if reminders.get("overrides"):
            for r in reminders["overrides"]:
                output += f"  Reminder: {r['method']} {r['minutes']}min before\n"

        if event.get("attachments"):
            output += "  Attachments:\n"
            for att in event["attachments"]:
                output += f"    - {att.get('title', 'untitled')} ({att.get('mimeType', '?')})\n"

    return output


def register(mcp):
    """Register all Calendar tools with the MCP server."""

    # ── gcal_list_events ─────────────────────────────────────────────────────

    class CalListEventsInput(BaseModel):
        """Input for listing calendar events."""
        model_config = ConfigDict(str_strip_whitespace=True)

        calendar_id: str = Field(default="primary", description="Calendar ID. Use 'primary' for your main calendar.")
        time_min: Optional[str] = Field(default=None, description="Start of time range (RFC3339, date, or 'YYYY-MM-DD HH:MM'). Defaults to now.")
        time_max: Optional[str] = Field(default=None, description="End of time range (same formats).")
        query: Optional[str] = Field(default=None, description="Keyword search across event title, description, and location.")
        max_results: int = Field(default=25, description="Maximum events to return.", ge=1, le=250)
        include_recurring: bool = Field(default=True, description="If true, expand recurring events into individual instances.")
        detailed: bool = Field(default=False, description="If true, include attendees, description, attachments, and Meet links.")
        status_filter: Optional[str] = Field(
            default=None,
            description="Filter by YOUR response status: 'accepted', 'tentative', 'declined', 'needsAction'.",
        )

    @mcp.tool(
        name="gcal_list_events",
        annotations={"title": "List Calendar Events", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gcal_list_events(params: CalListEventsInput) -> str:
        """List events from a Google Calendar with flexible filtering.

        Supports time range, keyword search, recurring event expansion,
        and response status filtering.

        Args:
            params: Calendar ID, time range, search query, and display options.

        Returns:
            Markdown-formatted event listing.
        """
        try:
            svc = get_calendar()
            kwargs = {
                "calendarId": params.calendar_id,
                "maxResults": params.max_results,
                "singleEvents": params.include_recurring,
                "orderBy": "startTime" if params.include_recurring else "updated",
            }

            if params.time_min:
                parsed = _parse_datetime(params.time_min)
                if "T" not in parsed:
                    parsed += "T00:00:00Z"
                if not parsed.endswith("Z") and "+" not in parsed and "-" not in parsed[10:]:
                    parsed += "Z"
                kwargs["timeMin"] = parsed
            else:
                kwargs["timeMin"] = datetime.now(timezone.utc).isoformat()

            if params.time_max:
                parsed = _parse_datetime(params.time_max)
                if "T" not in parsed:
                    parsed += "T23:59:59Z"
                if not parsed.endswith("Z") and "+" not in parsed and "-" not in parsed[10:]:
                    parsed += "Z"
                kwargs["timeMax"] = parsed

            if params.query:
                kwargs["q"] = params.query

            result = svc.events().list(**kwargs).execute()
            events = result.get("items", [])

            if params.status_filter:
                filtered = []
                for ev in events:
                    for att in ev.get("attendees", []):
                        if att.get("self") and att.get("responseStatus") == params.status_filter:
                            filtered.append(ev)
                            break
                events = filtered

            if not events:
                return "No events found."

            output = f"## Calendar Events ({len(events)})\n\n"
            for ev in events:
                output += _format_event(ev, detailed=params.detailed) + "\n"

            return output

        except Exception as e:
            return f"Error listing events: {e}"

    # ── gcal_get_event ───────────────────────────────────────────────────────

    class CalGetEventInput(BaseModel):
        """Input for getting a single calendar event."""
        model_config = ConfigDict(str_strip_whitespace=True)

        event_id: str = Field(..., description="Event ID.", min_length=1)
        calendar_id: str = Field(default="primary", description="Calendar ID.")

    @mcp.tool(
        name="gcal_get_event",
        annotations={"title": "Get Calendar Event", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gcal_get_event(params: CalGetEventInput) -> str:
        """Get full details for a single calendar event.

        Args:
            params: Event ID and calendar ID.

        Returns:
            Detailed event information.
        """
        try:
            event = get_calendar().events().get(
                calendarId=params.calendar_id,
                eventId=params.event_id,
            ).execute()
            return _format_event(event, detailed=True)
        except Exception as e:
            return f"Error getting event: {e}"

    # ── gcal_create_event ────────────────────────────────────────────────────

    class CalCreateEventInput(BaseModel):
        """Input for creating a calendar event."""
        model_config = ConfigDict(str_strip_whitespace=True)

        summary: str = Field(..., description="Event title.", min_length=1)
        start: str = Field(..., description="Start time (RFC3339, date 'YYYY-MM-DD', or 'YYYY-MM-DD HH:MM').")
        end: str = Field(..., description="End time (same formats as start).")
        calendar_id: str = Field(default="primary", description="Calendar ID.")
        timezone: Optional[str] = Field(default=None, description="IANA timezone (e.g., 'America/New_York'). Uses calendar default if omitted.")
        description: Optional[str] = Field(default=None, description="Event description.")
        location: Optional[str] = Field(default=None, description="Event location.")
        attendees: Optional[List[str]] = Field(default=None, description="Attendee email addresses.")
        recurrence: Optional[str] = Field(
            default=None,
            description="Recurrence: 'daily', 'weekly', 'monthly', 'yearly', or raw RRULE string.",
        )
        recurrence_count: Optional[int] = Field(default=None, description="Number of times to repeat.", ge=1)
        recurrence_until: Optional[str] = Field(default=None, description="Repeat until this date (YYYY-MM-DD).")
        reminders: Optional[List[dict]] = Field(
            default=None,
            description="Reminder list: [{'method': 'popup', 'minutes': 15}]. Method: 'popup' or 'email'.",
        )
        add_google_meet: bool = Field(default=False, description="If true, add a Google Meet video conference.")
        visibility: Optional[str] = Field(default=None, description="'default', 'public', or 'private'.")
        transparency: Optional[str] = Field(default=None, description="'opaque' (busy) or 'transparent' (free).")
        guests_can_modify: Optional[bool] = Field(default=None, description="Allow guests to modify the event.")
        send_notifications: str = Field(default="all", description="'all', 'none', or 'externalOnly'.")

    @mcp.tool(
        name="gcal_create_event",
        annotations={"title": "Create Calendar Event", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
    )
    async def gcal_create_event(params: CalCreateEventInput) -> str:
        """Create a new Google Calendar event with full options.

        Supports recurrence (simple strings like 'weekly' auto-convert to RRULE),
        Google Meet, attendees (just email strings), reminders, and visibility settings.

        Args:
            params: Event details.

        Returns:
            Confirmation with event ID and link.
        """
        try:
            start_parsed = _parse_datetime(params.start)
            end_parsed = _parse_datetime(params.end)

            is_all_day = len(start_parsed) == 10

            event_body = {"summary": params.summary}

            if is_all_day:
                event_body["start"] = {"date": start_parsed}
                event_body["end"] = {"date": end_parsed}
            else:
                start_obj = {"dateTime": start_parsed}
                end_obj = {"dateTime": end_parsed}
                if params.timezone:
                    start_obj["timeZone"] = params.timezone
                    end_obj["timeZone"] = params.timezone
                event_body["start"] = start_obj
                event_body["end"] = end_obj

            if params.description:
                event_body["description"] = params.description
            if params.location:
                event_body["location"] = params.location
            if params.attendees:
                event_body["attendees"] = [{"email": e} for e in params.attendees]
            if params.recurrence:
                event_body["recurrence"] = _build_recurrence(
                    params.recurrence, params.recurrence_count, params.recurrence_until
                )
            if params.reminders:
                event_body["reminders"] = {"useDefault": False, "overrides": params.reminders}
            if params.visibility:
                event_body["visibility"] = params.visibility
            if params.transparency:
                event_body["transparency"] = params.transparency
            if params.guests_can_modify is not None:
                event_body["guestsCanModify"] = params.guests_can_modify

            kwargs = {
                "calendarId": params.calendar_id,
                "body": event_body,
                "sendUpdates": params.send_notifications,
            }

            if params.add_google_meet:
                event_body["conferenceData"] = {
                    "createRequest": {
                        "conferenceSolutionKey": {"type": "hangoutsMeet"},
                        "requestId": f"gdrive-mcp-{datetime.now().strftime('%Y%m%d%H%M%S')}",
                    }
                }
                kwargs["conferenceDataVersion"] = 1

            event = get_calendar().events().insert(**kwargs).execute()

            output = f"Event created: **{event.get('summary')}**\n\n"
            output += f"Event ID: `{event['id']}`\n"
            output += f"Link: {event.get('htmlLink', 'N/A')}\n"

            hangout = event.get("hangoutLink")
            if hangout:
                output += f"Google Meet: {hangout}\n"

            return output

        except Exception as e:
            return f"Error creating event: {e}"

    # ── gcal_update_event ────────────────────────────────────────────────────

    class CalUpdateEventInput(BaseModel):
        """Input for updating a calendar event."""
        model_config = ConfigDict(str_strip_whitespace=True)

        event_id: str = Field(..., description="Event ID to update.", min_length=1)
        calendar_id: str = Field(default="primary", description="Calendar ID.")
        summary: Optional[str] = Field(default=None, description="New title.")
        start: Optional[str] = Field(default=None, description="New start time.")
        end: Optional[str] = Field(default=None, description="New end time.")
        timezone: Optional[str] = Field(default=None, description="IANA timezone.")
        description: Optional[str] = Field(default=None, description="New description.")
        location: Optional[str] = Field(default=None, description="New location.")
        add_attendees: Optional[List[str]] = Field(default=None, description="Email addresses to ADD to the event.")
        remove_attendees: Optional[List[str]] = Field(default=None, description="Email addresses to REMOVE from the event.")
        recurrence: Optional[str] = Field(default=None, description="New recurrence rule.")
        reminders: Optional[List[dict]] = Field(default=None, description="New reminders.")
        add_google_meet: Optional[bool] = Field(default=None, description="Add (true) or keep (null) Google Meet.")
        visibility: Optional[str] = Field(default=None, description="'default', 'public', or 'private'.")
        transparency: Optional[str] = Field(default=None, description="'opaque' or 'transparent'.")
        send_notifications: str = Field(default="all", description="'all', 'none', or 'externalOnly'.")

    @mcp.tool(
        name="gcal_update_event",
        annotations={"title": "Update Calendar Event", "readOnlyHint": False, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gcal_update_event(params: CalUpdateEventInput) -> str:
        """Update an existing calendar event with partial changes.

        Use add_attendees/remove_attendees to modify the guest list without
        having to re-submit the entire list. Only include fields you want to change.

        Args:
            params: Event ID and fields to update.

        Returns:
            Confirmation with updated event details.
        """
        try:
            svc = get_calendar()

            # Fetch existing event
            event = svc.events().get(
                calendarId=params.calendar_id,
                eventId=params.event_id,
            ).execute()

            # Apply changes
            updated = []

            if params.summary is not None:
                event["summary"] = params.summary
                updated.append("title")
            if params.description is not None:
                event["description"] = params.description
                updated.append("description")
            if params.location is not None:
                event["location"] = params.location
                updated.append("location")
            if params.visibility:
                event["visibility"] = params.visibility
                updated.append("visibility")
            if params.transparency:
                event["transparency"] = params.transparency
                updated.append("transparency")

            if params.start or params.end:
                if params.start:
                    parsed = _parse_datetime(params.start)
                    is_all_day = len(parsed) == 10
                    if is_all_day:
                        event["start"] = {"date": parsed}
                    else:
                        obj = {"dateTime": parsed}
                        if params.timezone:
                            obj["timeZone"] = params.timezone
                        event["start"] = obj
                if params.end:
                    parsed = _parse_datetime(params.end)
                    is_all_day = len(parsed) == 10
                    if is_all_day:
                        event["end"] = {"date": parsed}
                    else:
                        obj = {"dateTime": parsed}
                        if params.timezone:
                            obj["timeZone"] = params.timezone
                        event["end"] = obj
                updated.append("time")

            # Attendee management — add/remove without full list
            if params.add_attendees or params.remove_attendees:
                current = event.get("attendees", [])
                current_emails = {a["email"] for a in current}

                if params.remove_attendees:
                    remove_set = set(params.remove_attendees)
                    current = [a for a in current if a["email"] not in remove_set]

                if params.add_attendees:
                    for email_addr in params.add_attendees:
                        if email_addr not in current_emails:
                            current.append({"email": email_addr})

                event["attendees"] = current
                updated.append("attendees")

            if params.recurrence:
                event["recurrence"] = _build_recurrence(params.recurrence, None, None)
                updated.append("recurrence")

            if params.reminders:
                event["reminders"] = {"useDefault": False, "overrides": params.reminders}
                updated.append("reminders")

            kwargs = {
                "calendarId": params.calendar_id,
                "eventId": params.event_id,
                "body": event,
                "sendUpdates": params.send_notifications,
            }

            if params.add_google_meet:
                event["conferenceData"] = {
                    "createRequest": {
                        "conferenceSolutionKey": {"type": "hangoutsMeet"},
                        "requestId": f"gdrive-mcp-update-{datetime.now().strftime('%Y%m%d%H%M%S')}",
                    }
                }
                kwargs["conferenceDataVersion"] = 1
                updated.append("Google Meet")

            result = svc.events().update(**kwargs).execute()

            return f"Updated event **{result.get('summary')}**: {', '.join(updated)}.\nLink: {result.get('htmlLink', 'N/A')}"

        except Exception as e:
            return f"Error updating event: {e}"

    # ── gcal_delete_event ────────────────────────────────────────────────────

    class CalDeleteEventInput(BaseModel):
        """Input for deleting a calendar event."""
        model_config = ConfigDict(str_strip_whitespace=True)

        event_id: str = Field(..., description="Event ID to delete.", min_length=1)
        calendar_id: str = Field(default="primary", description="Calendar ID.")
        send_notifications: str = Field(default="all", description="'all', 'none', or 'externalOnly'.")

    @mcp.tool(
        name="gcal_delete_event",
        annotations={"title": "Delete Calendar Event", "readOnlyHint": False, "destructiveHint": True, "idempotentHint": True, "openWorldHint": True},
    )
    async def gcal_delete_event(params: CalDeleteEventInput) -> str:
        """Delete a calendar event.

        Args:
            params: Event ID and notification preference.

        Returns:
            Confirmation of deletion.
        """
        try:
            get_calendar().events().delete(
                calendarId=params.calendar_id,
                eventId=params.event_id,
                sendUpdates=params.send_notifications,
            ).execute()
            return f"Deleted event `{params.event_id}`."
        except Exception as e:
            return f"Error deleting event: {e}"

    # ── gcal_free_busy ───────────────────────────────────────────────────────

    class CalFreeBusyInput(BaseModel):
        """Input for checking free/busy status."""
        model_config = ConfigDict(str_strip_whitespace=True)

        calendars: List[str] = Field(default=["primary"], description="Calendar IDs to check.")
        time_min: str = Field(..., description="Start of time range.")
        time_max: str = Field(..., description="End of time range.")
        timezone: Optional[str] = Field(default=None, description="IANA timezone.")

    @mcp.tool(
        name="gcal_free_busy",
        annotations={"title": "Check Free/Busy", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gcal_free_busy(params: CalFreeBusyInput) -> str:
        """Check free/busy status for one or more calendars.

        Args:
            params: Calendar IDs and time range.

        Returns:
            Busy time blocks for each calendar.
        """
        try:
            time_min = _parse_datetime(params.time_min)
            time_max = _parse_datetime(params.time_max)

            if "T" not in time_min:
                time_min += "T00:00:00Z"
            if "T" not in time_max:
                time_max += "T23:59:59Z"

            for s in [time_min, time_max]:
                if not s.endswith("Z") and "+" not in s and "-" not in s[10:]:
                    s += "Z"

            body = {
                "timeMin": time_min,
                "timeMax": time_max,
                "items": [{"id": c} for c in params.calendars],
            }
            if params.timezone:
                body["timeZone"] = params.timezone

            result = get_calendar().freebusy().query(body=body).execute()

            output = "## Free/Busy\n\n"
            for cal_id, cal_data in result.get("calendars", {}).items():
                busy = cal_data.get("busy", [])
                output += f"**{cal_id}**\n"
                if not busy:
                    output += "  Free during this period.\n"
                else:
                    for block in busy:
                        output += f"  Busy: {block['start']} → {block['end']}\n"
                output += "\n"

            return output

        except Exception as e:
            return f"Error checking free/busy: {e}"

    # ── gcal_list_calendars ──────────────────────────────────────────────────

    class CalListCalendarsInput(BaseModel):
        """Input for listing calendars."""
        model_config = ConfigDict(str_strip_whitespace=True)

    @mcp.tool(
        name="gcal_list_calendars",
        annotations={"title": "List Calendars", "readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
    )
    async def gcal_list_calendars(params: CalListCalendarsInput) -> str:
        """List all available Google Calendars.

        Returns:
            All calendars with ID, name, access role, and timezone.
        """
        try:
            result = get_calendar().calendarList().list().execute()
            calendars = result.get("items", [])

            if not calendars:
                return "No calendars found."

            output = "## Your Calendars\n\n"
            output += "| Calendar | ID | Role | Timezone |\n|---|---|---|---|\n"
            for cal in calendars:
                name = cal.get("summary", "?")
                primary = " (primary)" if cal.get("primary") else ""
                cal_id = cal.get("id", "?")
                role = cal.get("accessRole", "?")
                tz = cal.get("timeZone", "?")
                output += f"| {name}{primary} | `{cal_id}` | {role} | {tz} |\n"

            return output

        except Exception as e:
            return f"Error listing calendars: {e}"
