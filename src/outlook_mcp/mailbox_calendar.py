from __future__ import annotations

from datetime import datetime

import pythoncom
import win32com.client as win32


OL_FOLDER_CALENDAR = 9
win32.gencache.is_readonly = True


def _safe_get(obj, name: str, default=None):
    try:
        return getattr(obj, name)
    except Exception:
        return default


def _parse_datetime(value: str) -> datetime:
    try:
        return datetime.fromisoformat(value).replace(tzinfo=None)
    except ValueError as exc:
        raise ValueError(
            f"Invalid datetime '{value}'. Use ISO format, for example 2026-09-16T09:00:00"
        ) from exc


def _normalize_com_datetime(value) -> datetime | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    try:
        return datetime.fromisoformat(str(value)).replace(tzinfo=None)
    except Exception:
        return None


def _serialize_datetime(value):
    normalized = _normalize_com_datetime(value)
    return normalized.isoformat() if normalized is not None else None


def _select_store(namespace, store_name: str):
    requested = store_name.strip().casefold()
    exact = None
    partial = []
    available = []

    count = namespace.Stores.Count
    for index in range(1, count + 1):
        store = namespace.Stores.Item(index)
        display_name = str(_safe_get(store, "DisplayName", ""))
        available.append(display_name)
        normalized = display_name.strip().casefold()
        if normalized == requested:
            exact = store
            break
        if requested and requested in normalized:
            partial.append(store)

    if exact is not None:
        return exact, available
    if len(partial) == 1:
        return partial[0], available
    return None, available


def _serialize_event(item) -> dict:
    attendees: list[str] = []
    required = _safe_get(item, "RequiredAttendees", "")
    if required:
        attendees = [part.strip() for part in str(required).split(";") if part.strip()]

    return {
        "entry_id": _safe_get(item, "EntryID"),
        "subject": str(_safe_get(item, "Subject", "")),
        "start": _serialize_datetime(_safe_get(item, "Start")),
        "end": _serialize_datetime(_safe_get(item, "End")),
        "location": str(_safe_get(item, "Location", "")),
        "body": str(_safe_get(item, "Body", "")),
        "all_day": bool(_safe_get(item, "AllDayEvent", False)),
        "busy_status": _safe_get(item, "BusyStatus"),
        "organizer": _safe_get(item, "Organizer"),
        "attendees": attendees,
    }


def list_mailbox_calendar_events(
    store_name: str,
    start: str,
    end: str,
    limit: int = 100,
) -> dict:
    """List events from an additional/shared Outlook Store calendar in read-only mode."""
    start_dt = _parse_datetime(start)
    end_dt = _parse_datetime(end)
    if end_dt <= start_dt:
        raise ValueError("end must be later than start")
    if limit < 1 or limit > 500:
        raise ValueError("limit must be between 1 and 500")

    pythoncom.CoInitialize()
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    store, available_stores = _select_store(namespace, store_name)
    if store is None:
        return {
            "status": "store_not_found",
            "store_name": store_name,
            "available_stores": available_stores,
            "start": start_dt.isoformat(),
            "end": end_dt.isoformat(),
            "events": [],
        }

    selected_store = str(_safe_get(store, "DisplayName", store_name))
    calendar = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    events: list[dict] = []
    scanned = 0
    skipped = 0
    comparison_errors = 0

    try:
        item = items.GetFirst()
    except Exception as exc:
        raise RuntimeError(f"Unable to read first calendar item from '{selected_store}': {exc}") from exc

    while item is not None and scanned < 10000:
        scanned += 1
        item_start = _normalize_com_datetime(_safe_get(item, "Start"))
        item_end = _normalize_com_datetime(_safe_get(item, "End"))

        if item_start is None or item_end is None:
            skipped += 1
        else:
            try:
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
        "calendar_source": "outlook_store",
        "store_name": selected_store,
        "folder_name": str(_safe_get(calendar, "Name", "")),
        "folder_path": str(_safe_get(calendar, "FolderPath", "")),
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "count": len(events),
        "events": events,
        "scanned": scanned,
        "skipped": skipped,
        "comparison_errors": comparison_errors,
        "filter_mode": "python_datetime_normalized",
    }
