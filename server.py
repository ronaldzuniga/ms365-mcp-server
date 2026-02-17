"""
Microsoft 365 MCP Server for Claude Desktop and Claude Code.
Provides email (Outlook) and calendar tools via Microsoft Graph API.
"""

import os
import sys
from typing import Optional, List

from mcp.server.fastmcp import FastMCP

import auth
import graph_client

# ── Configuration ────────────────────────────────────────────────

MS_CLIENT_ID = os.environ.get("MS_CLIENT_ID")
MS_TENANT_ID = os.environ.get("MS_TENANT_ID")

if not MS_CLIENT_ID or not MS_TENANT_ID:
    print("ERROR: MS_CLIENT_ID and MS_TENANT_ID environment variables are required.", file=sys.stderr)
    print("Set them in your Claude Desktop or Claude Code MCP config.", file=sys.stderr)
    sys.exit(1)


# ── Global token storage ──────────────────────────────────────────

_access_token: Optional[str] = None


def _get_token() -> str:
    """Get or refresh the access token."""
    global _access_token
    if not _access_token:
        _access_token = auth.get_access_token(MS_CLIENT_ID, MS_TENANT_ID)
    return _access_token


mcp = FastMCP("ms365_mcp")


def _handle_error(e: Exception) -> str:
    """Consistent error formatting."""
    import httpx as _httpx
    if isinstance(e, _httpx.HTTPStatusError):
        status = e.response.status_code
        if status == 401:
            return "Error: Authentication expired. Please restart the MCP server to re-authenticate."
        if status == 403:
            return "Error: Permission denied. Check the app's API permissions in Azure AD."
        if status == 404:
            return "Error: Resource not found. Please verify the ID."
        if status == 429:
            return "Error: Rate limited by Microsoft. Please wait a moment and retry."
        return f"Error: Microsoft Graph API returned status {status}: {e.response.text[:200]}"
    return f"Error: {type(e).__name__}: {str(e)}"


# ═══════════════════════════════════════════════════════════════════
# EMAIL TOOLS
# ═══════════════════════════════════════════════════════════════════


@mcp.tool(
    name="ms365_list_emails",
    annotations={"readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def ms365_list_emails(
    folder: str = "inbox",
    top: int = 10,
    skip: int = 0,
    unread_only: bool = False,
) -> str:
    """List emails from a mail folder (default: inbox). Returns subject, sender, date, and preview.

    Args:
        folder: Mail folder to read from: 'inbox', 'sentitems', 'drafts', 'deleteditems', or a folder ID.
        top: Number of emails to return (1-50, default: 10).
        skip: Number of emails to skip for pagination (default: 0).
        unread_only: If true, return only unread emails (default: false).

    Returns:
        Markdown-formatted list of emails with key metadata.
    """
    token = _get_token()
    try:
        # Validate parameters
        if top < 1 or top > 50:
            return "Error: 'top' must be between 1 and 50."
        if skip < 0:
            return "Error: 'skip' must be >= 0."

        filter_q = "isRead eq false" if unread_only else None
        data = await graph_client.list_messages(
            token, top=top, skip=skip, filter_query=filter_q, folder=folder
        )
        messages = data.get("value", [])
        if not messages:
            return "No emails found."

        lines = [f"**Showing {len(messages)} email(s) from '{folder}':**\n"]
        for msg in messages:
            sender = msg.get("from", {}).get("emailAddress", {})
            sender_str = f"{sender.get('name', 'Unknown')} <{sender.get('address', '')}>"
            read_flag = "" if msg.get("isRead") else " 🔵"
            attach = " 📎" if msg.get("hasAttachments") else ""
            lines.append(
                f"- **{msg['subject']}**{read_flag}{attach}\n"
                f"  From: {sender_str}\n"
                f"  Date: {msg['receivedDateTime']}\n"
                f"  Preview: {msg.get('bodyPreview', '')[:120]}...\n"
                f"  ID: `{msg['id']}`\n"
            )

        total = data.get("@odata.count")
        if total:
            lines.append(f"\n_Total: {total} | Showing {skip + 1}–{skip + len(messages)}_")
        return "\n".join(lines)
    except Exception as e:
        return _handle_error(e)


@mcp.tool(
    name="ms365_read_email",
    annotations={"readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def ms365_read_email(message_id: str) -> str:
    """Read the full content of an email message by its ID.

    Args:
        message_id: The email message ID (from ms365_list_emails or ms365_search_emails).

    Returns:
        Full email with headers and body content.
    """
    token = _get_token()
    try:
        if not message_id or not message_id.strip():
            return "Error: message_id is required."
        msg = await graph_client.get_message(token, message_id)
        sender = msg.get("from", {}).get("emailAddress", {})
        to_list = ", ".join(
            f"{r['emailAddress'].get('name', '')} <{r['emailAddress']['address']}>"
            for r in msg.get("toRecipients", [])
        )
        cc_list = ", ".join(
            f"{r['emailAddress'].get('name', '')} <{r['emailAddress']['address']}>"
            for r in msg.get("ccRecipients", [])
        )

        body_content = msg.get("body", {}).get("content", "No body")
        body_type = msg.get("body", {}).get("contentType", "Text")

        lines = [
            f"# {msg['subject']}\n",
            f"**From:** {sender.get('name', 'Unknown')} <{sender.get('address', '')}>",
            f"**To:** {to_list}",
        ]
        if cc_list:
            lines.append(f"**CC:** {cc_list}")
        lines.extend([
            f"**Date:** {msg['receivedDateTime']}",
            f"**Read:** {'Yes' if msg.get('isRead') else 'No'}",
            f"**Attachments:** {'Yes' if msg.get('hasAttachments') else 'No'}",
            f"\n---\n",
        ])

        if body_type == "html":
            lines.append(f"_(HTML email — showing raw HTML)_\n\n{body_content}")
        else:
            lines.append(body_content)

        return "\n".join(lines)
    except Exception as e:
        return _handle_error(e)


@mcp.tool(
    name="ms365_search_emails",
    annotations={"readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def ms365_search_emails(query: str, top: int = 10) -> str:
    """Search emails across all folders by keyword.

    Args:
        query: Search query (searches subject, body, sender, etc.).
        top: Max results to return (1-25, default: 10).

    Returns:
        Markdown-formatted list of matching emails.
    """
    token = _get_token()
    try:
        if not query or not query.strip():
            return "Error: query is required."
        if top < 1 or top > 25:
            return "Error: 'top' must be between 1 and 25."
        data = await graph_client.search_messages(token, query, top=top)
        messages = data.get("value", [])
        if not messages:
            return f"No emails found matching '{query}'."

        lines = [f"**Found {len(messages)} email(s) matching '{query}':**\n"]
        for msg in messages:
            sender = msg.get("from", {}).get("emailAddress", {})
            sender_str = f"{sender.get('name', 'Unknown')} <{sender.get('address', '')}>"
            lines.append(
                f"- **{msg['subject']}**\n"
                f"  From: {sender_str} | Date: {msg['receivedDateTime']}\n"
                f"  Preview: {msg.get('bodyPreview', '')[:120]}...\n"
                f"  ID: `{msg['id']}`\n"
            )
        return "\n".join(lines)
    except Exception as e:
        return _handle_error(e)


@mcp.tool(
    name="ms365_send_email",
    annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
)
async def ms365_send_email(
    to: List[str],
    subject: str,
    body: str,
    cc: Optional[List[str]] = None,
    html: bool = False,
) -> str:
    """Send an email from your Microsoft 365 account.

    Args:
        to: List of recipient email addresses.
        subject: Email subject line.
        body: Email body content.
        cc: Optional CC recipients.
        html: If true, body is treated as HTML (default: false).

    Returns:
        Confirmation message.
    """
    token = _get_token()
    try:
        if not to or len(to) == 0:
            return "Error: At least one recipient is required in 'to' field."
        if not subject or not subject.strip():
            return "Error: 'subject' is required."

        await graph_client.send_message(
            token,
            to_recipients=to,
            subject=subject,
            body=body,
            cc_recipients=cc,
            is_html=html,
        )
        to_str = ", ".join(to)
        return f"Email sent successfully to {to_str} with subject '{subject}'."
    except Exception as e:
        return _handle_error(e)


@mcp.tool(
    name="ms365_reply_email",
    annotations={"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True},
)
async def ms365_reply_email(message_id: str, comment: str) -> str:
    """Reply to an email message.

    Args:
        message_id: The ID of the message to reply to.
        comment: Reply text content.

    Returns:
        Confirmation message.
    """
    token = _get_token()
    try:
        if not message_id or not message_id.strip():
            return "Error: message_id is required."
        if not comment or not comment.strip():
            return "Error: comment is required."
        await graph_client.reply_to_message(token, message_id, comment)
        return "Reply sent successfully."
    except Exception as e:
        return _handle_error(e)


@mcp.tool(
    name="ms365_list_mail_folders",
    annotations={"readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def ms365_list_mail_folders() -> str:
    """List all mail folders with message counts.

    Returns:
        Markdown-formatted list of folders.
    """
    token = _get_token()
    try:
        data = await graph_client.list_mail_folders(token)
        folders = data.get("value", [])
        if not folders:
            return "No mail folders found."

        lines = ["**Mail Folders:**\n"]
        for f in folders:
            unread = f.get("unreadItemCount", 0)
            total = f.get("totalItemCount", 0)
            unread_str = f" ({unread} unread)" if unread > 0 else ""
            lines.append(f"- **{f['displayName']}** — {total} messages{unread_str} | ID: `{f['id']}`")
        return "\n".join(lines)
    except Exception as e:
        return _handle_error(e)


# ═══════════════════════════════════════════════════════════════════
# CALENDAR TOOLS
# ═══════════════════════════════════════════════════════════════════


@mcp.tool(
    name="ms365_list_events",
    annotations={"readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def ms365_list_events(
    start_datetime: str,
    end_datetime: str,
    top: int = 25,
    calendar_id: Optional[str] = None,
) -> str:
    """List calendar events within a date/time range.

    Args:
        start_datetime: Start of time range in ISO 8601 format (e.g. '2026-02-16T00:00:00Z').
        end_datetime: End of time range in ISO 8601 format (e.g. '2026-02-17T23:59:59Z').
        top: Max events to return (1-50, default: 25).
        calendar_id: Specific calendar ID. Omit for default calendar.

    Returns:
        Markdown-formatted list of events.
    """
    token = _get_token()
    try:
        if not start_datetime or not start_datetime.strip():
            return "Error: start_datetime is required."
        if not end_datetime or not end_datetime.strip():
            return "Error: end_datetime is required."
        if top < 1 or top > 50:
            return "Error: 'top' must be between 1 and 50."

        data = await graph_client.list_events(
            token,
            start_datetime=start_datetime,
            end_datetime=end_datetime,
            top=top,
            calendar_id=calendar_id,
        )
        events = data.get("value", [])
        if not events:
            return f"No events found between {start_datetime} and {end_datetime}."

        lines = [f"**{len(events)} event(s) found:**\n"]
        for ev in events:
            start = ev.get("start", {})
            end = ev.get("end", {})
            organizer = ev.get("organizer", {}).get("emailAddress", {})
            location = ev.get("location", {}).get("displayName", "")
            all_day = " (All day)" if ev.get("isAllDay") else ""
            loc_str = f"\n  Location: {location}" if location else ""
            attendee_count = len(ev.get("attendees", []))
            meeting_link = ""
            if ev.get("onlineMeeting"):
                join_url = ev["onlineMeeting"].get("joinUrl", "")
                if join_url:
                    meeting_link = f"\n  Join: {join_url}"

            lines.append(
                f"- **{ev['subject']}**{all_day}\n"
                f"  {start.get('dateTime', '')} → {end.get('dateTime', '')}{loc_str}\n"
                f"  Organizer: {organizer.get('name', '')} <{organizer.get('address', '')}>\n"
                f"  Attendees: {attendee_count}{meeting_link}\n"
                f"  ID: `{ev['id']}`\n"
            )
        return "\n".join(lines)
    except Exception as e:
        return _handle_error(e)


@mcp.tool(
    name="ms365_get_event",
    annotations={"readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def ms365_get_event(event_id: str) -> str:
    """Get full details of a calendar event.

    Args:
        event_id: Calendar event ID.

    Returns:
        Full event details including body and attendees.
    """
    token = _get_token()
    try:
        if not event_id or not event_id.strip():
            return "Error: event_id is required."
        ev = await graph_client.get_event(token, event_id)
        organizer = ev.get("organizer", {}).get("emailAddress", {})
        start = ev.get("start", {})
        end = ev.get("end", {})
        location = ev.get("location", {}).get("displayName", "No location")
        body_content = ev.get("body", {}).get("content", "No body")

        lines = [
            f"# {ev['subject']}\n",
            f"**When:** {start.get('dateTime', '')} → {end.get('dateTime', '')}",
            f"**Location:** {location}",
            f"**Organizer:** {organizer.get('name', '')} <{organizer.get('address', '')}>",
            f"**All Day:** {'Yes' if ev.get('isAllDay') else 'No'}",
        ]

        attendees = ev.get("attendees", [])
        if attendees:
            lines.append("\n**Attendees:**")
            for att in attendees:
                email_info = att.get("emailAddress", {})
                status = att.get("status", {}).get("response", "none")
                lines.append(f"  - {email_info.get('name', '')} <{email_info.get('address', '')}> ({status})")

        meeting = ev.get("onlineMeeting")
        if meeting and meeting.get("joinUrl"):
            lines.append(f"\n**Join Link:** {meeting['joinUrl']}")

        lines.append(f"\n---\n\n{body_content}")
        return "\n".join(lines)
    except Exception as e:
        return _handle_error(e)


@mcp.tool(
    name="ms365_list_calendars",
    annotations={"readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def ms365_list_calendars() -> str:
    """List all available calendars.

    Returns:
        Markdown-formatted list of calendars.
    """
    token = _get_token()
    try:
        data = await graph_client.list_calendars(token)
        calendars = data.get("value", [])
        if not calendars:
            return "No calendars found."

        lines = ["**Your Calendars:**\n"]
        for cal in calendars:
            default = " ⭐ (default)" if cal.get("isDefaultCalendar") else ""
            owner = cal.get("owner", {}).get("address", "")
            lines.append(f"- **{cal['name']}**{default} | Owner: {owner} | ID: `{cal['id']}`")
        return "\n".join(lines)
    except Exception as e:
        return _handle_error(e)


@mcp.tool(
    name="ms365_find_free_time",
    annotations={"readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def ms365_find_free_time(
    emails: List[str],
    start_datetime: str,
    end_datetime: str,
    timezone: str = "America/Costa_Rica",
) -> str:
    """Check free/busy availability for one or more people.

    Args:
        emails: List of email addresses to check availability for.
        start_datetime: Start of range in ISO 8601 (e.g. '2026-02-17T08:00:00').
        end_datetime: End of range in ISO 8601 (e.g. '2026-02-17T18:00:00').
        timezone: IANA timezone (e.g. 'America/Costa_Rica', 'UTC'). Default: 'America/Costa_Rica'.

    Returns:
        Availability view for each person (0=free, 1=tentative, 2=busy, 3=out of office, 4=working elsewhere).
    """
    token = _get_token()
    try:
        if not emails or len(emails) == 0:
            return "Error: At least one email address is required."
        if not start_datetime or not start_datetime.strip():
            return "Error: start_datetime is required."
        if not end_datetime or not end_datetime.strip():
            return "Error: end_datetime is required."

        data = await graph_client.find_free_busy(
            token,
            schedules=emails,
            start_datetime=start_datetime,
            end_datetime=end_datetime,
            timezone=timezone,
        )
        schedules = data.get("value", [])
        if not schedules:
            return "No availability data returned."

        legend = "Legend: 0=Free, 1=Tentative, 2=Busy, 3=OOF, 4=Working Elsewhere"
        lines = [f"**Availability ({start_datetime} → {end_datetime}, {timezone}):**\n", legend, ""]

        for sched in schedules:
            email = sched.get("scheduleId", "Unknown")
            view = sched.get("availabilityView", "")
            lines.append(f"**{email}:** `{view}`")

            items = sched.get("scheduleItems", [])
            if items:
                for item in items:
                    status = item.get("status", "")
                    subj = item.get("subject", "")
                    s = item.get("start", {}).get("dateTime", "")
                    e = item.get("end", {}).get("dateTime", "")
                    lines.append(f"  - [{status}] {subj}: {s} → {e}")
            lines.append("")

        return "\n".join(lines)
    except Exception as e:
        return _handle_error(e)


# ═══════════════════════════════════════════════════════════════════
# USER PROFILE
# ═══════════════════════════════════════════════════════════════════


@mcp.tool(
    name="ms365_get_profile",
    annotations={"readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": True},
)
async def ms365_get_profile() -> str:
    """Get the authenticated user's Microsoft 365 profile.

    Returns:
        User profile info (name, email, job title, etc.).
    """
    token = _get_token()
    try:
        me = await graph_client.get_me(token)
        lines = [
            "**Your Microsoft 365 Profile:**\n",
            f"- **Name:** {me.get('displayName', 'N/A')}",
            f"- **Email:** {me.get('mail', me.get('userPrincipalName', 'N/A'))}",
            f"- **Job Title:** {me.get('jobTitle', 'N/A')}",
            f"- **Office:** {me.get('officeLocation', 'N/A')}",
            f"- **Phone:** {me.get('businessPhones', ['N/A'])[0] if me.get('businessPhones') else 'N/A'}",
        ]
        return "\n".join(lines)
    except Exception as e:
        return _handle_error(e)


# ═══════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    _get_token()
    mcp.run()
