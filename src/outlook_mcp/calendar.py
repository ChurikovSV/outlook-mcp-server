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


def _safe_get(item, name: str, default=None):
    try:
        return getattr(item, name)
    except Exception:
        return default


def _normalize_com_datetime(value) -> datetime | None:
    """Normalize Outlook/pywintypes datetime values for local naive comparisons.

    Outlook COM can return timezone-aware pywintypes datetime objects while MCP
    request datetimes are local naive values. Python refuses to compare aware and
    naive datetimes, which previously caused every calendar item to be skipped.
    """
    if value is None:
        return None

    try:
        if isinstance(value, datetime):
            return value.replace(tzinfo=None)
    except Exception:
        pass

    try:
        parsed = datetime.fromisoformat(str(value))
        return parsed.replace(tzinfo=None)
    except Exception:
        return None


def _serialize_datetime(value):
    normalized = _normalize_com_datetime(value)
    if normalized is not None:
        return normalized.isoformat()
    if value is None:
        return None
    try:
        return str(value)
    except Exception:
        return None


def _serialize_event(item) -> dict:
    """Serialize an Outlook calendar item without letting one COM property abort the whole request."""
    attendees: list[str] = []
    recipients = _safe_get(item, "Recipients")
    if recipients is not None:
        try:
            count = recipients.Count
            for index in range(1, count + 1):
                try:
                    recipient = recipients.Item(index)
                    address = _safe_get(recipient, "Address")
                    name = _safe_get(recipient, "Name", "")
                    if address or name:
                        attendees.append(address or name)
                except Exception:
                    continue
        except Exception:
            pass

    return {
        "entry_id": _safe_get(item, "EntryID"),
        "subject": _safe_get(item, "Subject", ""),
        "start": _serialize_datetime(_safe_get(item, "Start")),
        "end": _serialize_datetime(_safe_get(item, "End")),
        "location": _safe_get(item, "Location", ""),
        "body": _safe_get(item, "Body", ""),
        "all_day": bool(_safe_get(item, "AllDayEvent", False)),
        "busy_status": _safe_get(item, "BusyStatus"),
        "organizer": _safe_get(item, "Organizer"),
        "attendees": attendees,
    }


def list_calendar_events(start: str, end: str, limit: int = 100) -> dict:
    """List calendar items in the requested local datetime range.

    Filtering is performed in Python to avoid Outlook Restrict() locale-dependent
    date parsing. Outlook COM datetime values are normalized to local naive
    datetimes before comparison.
    """
    start_dt = _parse_datetime(start).replace(tzinfo=None)
    end_dt = _parse_datetime(end).replace(tzinfo=None)
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

    events: list[dict] = []
    skipped = 0
    scanned = 0
    comparison_errors = 0

    try:
        item = items.GetFirst()
    except Exception as exc:
        raise RuntimeError(f"Unable to read first calendar item: {exc}") from exc

    max_scan = 10000

    while item is not None and scanned < max_scan:
        scanned += 1

        raw_start = _safe_get(item, "Start")
        raw_end = _safe_get(item, "End")
        item_start = _normalize_com_datetime(raw_start)
        item_end = _normalize_com_datetime(raw_end)

        if item_start is None or item_end is None:
            skipped += 1
        else:
            try:
                # Items are sorted ascending by Start. Once we move beyond the
                # requested range there is no need to continue scanning.
                if item_start >= end_dt:
                    break

                if item_start < end_dt and item_end > start_dt:
                    events.append(_serialize_event(item))
                    if len(events) >= limit:
                        break
            except Exception:
                comparison_errors += 1
                skipped += 1

        try:
            item = items.GetNext()
        except Exception:
            skipped += 1
            break

    return {
        "status": "ok",
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "count": len(events),
        "events": events,
        "scanned": scanned,
        "skipped": skipped,
        "comparison_errors": comparison_errors,
        "filter_mode": "python_datetime_normalized",
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
        "entry_id": _safe_get(item, "EntryID"),
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

    current_start = _normalize_com_datetime(item.Start)
    current_end = _normalize_com_datetime(item.End)
    if current_start is None or current_end is None or current_end <= current_start:
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
    subject = _safe_get(item, "Subject", "")
    item.Delete()
    return {
        "status": "calendar_event_deleted",
        "entry_id": entry_id,
        "subject": subject,
        "cancellation_sent": False,
    }
