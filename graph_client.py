"""
Microsoft Graph API client for email and calendar operations.
"""

import httpx
from typing import Optional

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TIMEOUT = 30.0


async def graph_request(
    access_token: str,
    method: str,
    endpoint: str,
    params: Optional[dict] = None,
    json_body: Optional[dict] = None,
) -> dict:
    """Make an authenticated request to the Microsoft Graph API."""
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    url = f"{GRAPH_BASE}{endpoint}"

    async with httpx.AsyncClient(timeout=TIMEOUT) as client:
        response = await client.request(
            method=method,
            url=url,
            headers=headers,
            params=params,
            json=json_body,
        )
        response.raise_for_status()
        # Handle empty response bodies (204 No Content, 202 Accepted)
        if response.status_code in (202, 204) or len(response.content) == 0:
            return {"status": "success"}
        return response.json()


# ── Email helpers ────────────────────────────────────────────────


async def list_messages(
    access_token: str,
    top: int = 10,
    skip: int = 0,
    filter_query: Optional[str] = None,
    search_query: Optional[str] = None,
    folder: str = "inbox",
) -> dict:
    """List email messages from a folder."""
    params = {
        "$top": top,
        "$skip": skip,
        "$orderby": "receivedDateTime desc",
        "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,bodyPreview,hasAttachments",
    }
    if filter_query:
        params["$filter"] = filter_query
    if search_query:
        params["$search"] = f'"{search_query}"'

    return await graph_request(access_token, "GET", f"/me/mailFolders/{folder}/messages", params=params)


async def get_message(access_token: str, message_id: str) -> dict:
    """Get a single email message with full body."""
    params = {
        "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,body,hasAttachments,attachments",
    }
    return await graph_request(access_token, "GET", f"/me/messages/{message_id}", params=params)


async def search_messages(access_token: str, query: str, top: int = 10) -> dict:
    """Search emails using Microsoft Graph $search."""
    params = {
        "$search": f'"{query}"',
        "$top": top,
        "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,bodyPreview,hasAttachments",
    }
    return await graph_request(access_token, "GET", "/me/messages", params=params)


async def send_message(
    access_token: str,
    to_recipients: list[str],
    subject: str,
    body: str,
    cc_recipients: Optional[list[str]] = None,
    is_html: bool = False,
) -> dict:
    """Send an email message."""
    message = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML" if is_html else "Text",
                "content": body,
            },
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_recipients],
        }
    }
    if cc_recipients:
        message["message"]["ccRecipients"] = [
            {"emailAddress": {"address": addr}} for addr in cc_recipients
        ]

    return await graph_request(access_token, "POST", "/me/sendMail", json_body=message)


async def reply_to_message(
    access_token: str,
    message_id: str,
    comment: str,
) -> dict:
    """Reply to an email message."""
    body = {"comment": comment}
    return await graph_request(access_token, "POST", f"/me/messages/{message_id}/reply", json_body=body)


async def list_mail_folders(access_token: str) -> dict:
    """List mail folders."""
    params = {"$select": "id,displayName,totalItemCount,unreadItemCount"}
    return await graph_request(access_token, "GET", "/me/mailFolders", params=params)


# ── Calendar helpers ─────────────────────────────────────────────


async def list_events(
    access_token: str,
    start_datetime: str,
    end_datetime: str,
    top: int = 25,
    calendar_id: Optional[str] = None,
) -> dict:
    """List calendar events in a time range (ISO 8601 format)."""
    endpoint = f"/me/calendars/{calendar_id}/calendarView" if calendar_id else "/me/calendarView"
    params = {
        "startDateTime": start_datetime,
        "endDateTime": end_datetime,
        "$top": top,
        "$orderby": "start/dateTime",
        "$select": "id,subject,organizer,start,end,location,attendees,isAllDay,bodyPreview,onlineMeeting,recurrence",
    }
    return await graph_request(access_token, "GET", endpoint, params=params)


async def get_event(access_token: str, event_id: str) -> dict:
    """Get a single calendar event with full details."""
    params = {
        "$select": "id,subject,organizer,start,end,location,attendees,isAllDay,body,onlineMeeting,recurrence",
    }
    return await graph_request(access_token, "GET", f"/me/events/{event_id}", params=params)


async def list_calendars(access_token: str) -> dict:
    """List all calendars for the user."""
    params = {"$select": "id,name,color,isDefaultCalendar,owner"}
    return await graph_request(access_token, "GET", "/me/calendars", params=params)


async def find_free_busy(
    access_token: str,
    schedules: list[str],
    start_datetime: str,
    end_datetime: str,
    timezone: str = "UTC",
) -> dict:
    """Get free/busy status for one or more users."""
    body = {
        "schedules": schedules,
        "startTime": {"dateTime": start_datetime, "timeZone": timezone},
        "endTime": {"dateTime": end_datetime, "timeZone": timezone},
        "availabilityViewInterval": 30,
    }
    return await graph_request(access_token, "POST", "/me/calendar/getSchedule", json_body=body)


# ── User profile ─────────────────────────────────────────────────


async def get_me(access_token: str) -> dict:
    """Get the authenticated user's profile."""
    return await graph_request(access_token, "GET", "/me")
