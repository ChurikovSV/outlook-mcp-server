from __future__ import annotations

from datetime import datetime

import pythoncom
import win32com.client as win32


OL_APPOINTMENT_ITEM = 1
OL_FOLDER_CALENDAR = 9
OL_MEETING = 1
OL_BUSY = 2

# Keep COM setup aligned with the mail implementation.
win32.gencache.is_readonly = True


def _get_outlook():
    pythoncom.CoInitialize()
    return win32.Dispatch("Outlook.Application")


def _parse_datetime(value: str) -> datetime:
    try:
        return datetime.fromisoformat(value)
    except ValueError as exc:
        raise ValueError(
            f"Invalid datetime '{value}'. Use ISO format, for example 2026-09-10T15:30:00"
        ) from exc


def _serialize_event(item) -> dict:
    attendees: list[str] = []
    try:
        for recipient in item.Recipients:
            address = None
            try:
                address = recipient.Address
            except Exception:
                pass
            attendees.append(address or recipient.Name)
    except Exception:
        pass

    start = getattr(item, "Start", None)
    end = getattr(item, "End", None)

    return {
        "entry_id": getattr(item, "EntryID", None),
        "subject": getattr(item, "Subject", ""),
        "start": start.isoformat() if hasattr(start, "isoformat") else str(start) if start else None,
        "end": end.isoformat() if hasattr(end, "isoformat") else str(end) if end else None,
        "location": getattr(item, "Location", ""),
        "body": getattr(item, "Body", ""),
        "all_day": bool(getattr(item, "AllDayEvent", False)),
        "busy_status": getattr(item, "BusyStatus", None),
        "organizer": getattr(item, "Organizer", None),
        "attendees": attendees,
    }


def list_calendar_events(start: str, end: str, limit: int = 100) -> dict:
    """List calendar items in the requested local datetime range."""
    start_dt = _parse_datetime(start)
    end_dt = _parse_datetime(end)
    if end_dt <= start_dt:
        raise ValueError("end must be later than start")
    if limit < 1 or limit > 500:
        raise ValueError("limit must be between 1 and 500")

    outlook = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")
    calendar = namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    # Outlook Restrict expects a locale-style date string. This US-style format is
    # the most portable form for Outlook's Jet restriction parser.
    start_filter = start_dt.strftime("%m/%d/%Y %I:%M %p")
    end_filter = end_dt.strftime("%m/%d/%Y %I:%M %p")
    restricted = items.Restrict(
        f"[Start] < '{end_filter}' AND [End] > '{start_filter}'"
    )

    events: list[dict] = []
    for item in restricted:
        events.append(_serialize_event(item))
        if len(events) >= limit:
            break

    return {
        "status": "ok",
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "count": len(events),
        "events": events,
    }


def create_calendar_event(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    attendees: list[str] | None = None,
    all_day: bool = False,
    reminder_minutes: int | None = 15,
) -> dict:
    """Create and save a calendar event without sending invitations."""
    start_dt = _parse_datetime(start)
    end_dt = _parse_datetime(end)
    if end_dt <= start_dt:
        raise ValueError("end must be later than start")

    outlook = _get_outlook()
    item = outlook.CreateItem(OL_APPOINTMENT_ITEM)
    item.Subject = subject
    item.Start = start_dt
    item.End = end_dt
    item.Location = location
    item.Body = body
    item.AllDayEvent = all_day
    item.BusyStatus = OL_BUSY

    if reminder_minutes is None:
        item.ReminderSet = False
    else:
        if reminder_minutes < 0:
            raise ValueError("reminder_minutes must be >= 0 or null")
        item.ReminderSet = True
        item.ReminderMinutesBeforeStart = reminder_minutes

    attendee_list = attendees or []
    if attendee_list:
        item.MeetingStatus = OL_MEETING
        for attendee in attendee_list:
            recipient = item.Recipients.Add(attendee)
            recipient.Type = 1

    # Deliberately Save only. Never call Send(); invitations remain unsent.
    item.Save()

    return {
        "status": "calendar_event_created",
        "entry_id": getattr(item, "EntryID", None),
        "subject": subject,
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "attendees": attendee_list,
        "invitations_sent": False,
    }


def update_calendar_event(
    entry_id: str,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    location: str | None = None,
    body: str | None = None,
    all_day: bool | None = None,
    reminder_minutes: int | None = None,
    disable_reminder: bool = False,
) -> dict:
    """Update an existing calendar item and save it without sending updates."""
    outlook = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")
    item = namespace.GetItemFromID(entry_id)

    new_start = _parse_datetime(start) if start is not None else None
    new_end = _parse_datetime(end) if end is not None else None

    if new_start is not None:
        item.Start = new_start
    if new_end is not None:
        item.End = new_end

    # Validate final interval after applying optional changes.
    current_start = getattr(item, "Start")
    current_end = getattr(item, "End")
    if current_end <= current_start:
        raise ValueError("end must be later than start")

    if subject is not None:
        item.Subject = subject
    if location is not None:
        item.Location = location
    if body is not None:
        item.Body = body
    if all_day is not None:
        item.AllDayEvent = all_day

    if disable_reminder:
        item.ReminderSet = False
    elif reminder_minutes is not None:
        if reminder_minutes < 0:
            raise ValueError("reminder_minutes must be >= 0")
        item.ReminderSet = True
        item.ReminderMinutesBeforeStart = reminder_minutes

    # Deliberately Save only. Never call Send().
    item.Save()

    result = _serialize_event(item)
    result.update({"status": "calendar_event_updated", "updates_sent": False})
    return result


def delete_calendar_event(entry_id: str) -> dict:
    """Delete a calendar item locally. No cancellation message is sent."""
    outlook = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")
    item = namespace.GetItemFromID(entry_id)
    subject = getattr(item, "Subject", "")
    item.Delete()
    return {
        "status": "calendar_event_deleted",
        "entry_id": entry_id,
        "subject": subject,
        "cancellation_sent": False,
    }
