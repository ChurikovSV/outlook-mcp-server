from __future__ import annotations

from datetime import datetime, timedelta

import pythoncom
import win32com.client


OL_FOLDER_CALENDAR = 9


def _get_outlook_with_steps(result: dict):
    try:
        pythoncom.CoInitialize()
        result["steps"].append({"step": "coinitialize", "status": "ok"})
    except Exception as exc:
        result["steps"].append({"step": "coinitialize", "status": "error", "error": repr(exc)})
        return None

    outlook = None

    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
        result["steps"].append({"step": "get_active_outlook", "status": "ok"})
    except Exception as exc:
        result["steps"].append({"step": "get_active_outlook", "status": "not_available", "error": repr(exc)})

    if outlook is None:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            result["steps"].append({"step": "dispatch_outlook", "status": "ok"})
        except Exception as exc:
            result["steps"].append({"step": "dispatch_outlook", "status": "error", "error": repr(exc)})
            return None

    return outlook


def diagnose_outlook() -> dict:
    """Run step-by-step Outlook COM diagnostics without raising COM errors."""
    result: dict = {"status": "error", "steps": []}
    outlook = _get_outlook_with_steps(result)
    if outlook is None:
        return result

    try:
        version = str(outlook.Version)
        result["steps"].append({"step": "read_outlook_version", "status": "ok", "version": version})
    except Exception as exc:
        result["steps"].append({"step": "read_outlook_version", "status": "error", "error": repr(exc)})

    try:
        namespace = outlook.GetNamespace("MAPI")
        result["steps"].append({"step": "get_mapi_namespace", "status": "ok"})
    except Exception as exc:
        result["steps"].append({"step": "get_mapi_namespace", "status": "error", "error": repr(exc)})
        return result

    try:
        accounts = []
        for account in namespace.Accounts:
            try:
                accounts.append(str(account.SmtpAddress))
            except Exception:
                accounts.append(str(account.DisplayName))
        result["steps"].append({"step": "read_accounts", "status": "ok", "accounts": accounts})
    except Exception as exc:
        result["steps"].append({"step": "read_accounts", "status": "error", "error": repr(exc)})
        return result

    result["status"] = "ok"
    return result


def diagnose_calendar() -> dict:
    """Diagnose Outlook calendar access step by step without raising COM errors."""
    result: dict = {"status": "error", "steps": []}
    outlook = _get_outlook_with_steps(result)
    if outlook is None:
        return result

    try:
        namespace = outlook.GetNamespace("MAPI")
        result["steps"].append({"step": "get_mapi_namespace", "status": "ok"})
    except Exception as exc:
        result["steps"].append({"step": "get_mapi_namespace", "status": "error", "error": repr(exc)})
        return result

    try:
        calendar = namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)
        folder_info = {
            "name": str(getattr(calendar, "Name", "")),
            "folder_path": str(getattr(calendar, "FolderPath", "")),
            "entry_id": str(getattr(calendar, "EntryID", "")),
        }
        result["steps"].append({"step": "get_default_calendar", "status": "ok", **folder_info})
    except Exception as exc:
        result["steps"].append({"step": "get_default_calendar", "status": "error", "error": repr(exc)})
        return result

    try:
        items = calendar.Items
        result["steps"].append({"step": "get_calendar_items", "status": "ok"})
    except Exception as exc:
        result["steps"].append({"step": "get_calendar_items", "status": "error", "error": repr(exc)})
        return result

    try:
        count = int(items.Count)
        result["steps"].append({"step": "read_items_count", "status": "ok", "count": count})
    except Exception as exc:
        result["steps"].append({"step": "read_items_count", "status": "error", "error": repr(exc)})
        return result

    try:
        items.IncludeRecurrences = True
        result["steps"].append({"step": "set_include_recurrences", "status": "ok"})
    except Exception as exc:
        result["steps"].append({"step": "set_include_recurrences", "status": "error", "error": repr(exc)})
        return result

    try:
        items.Sort("[Start]")
        result["steps"].append({"step": "sort_by_start", "status": "ok"})
    except Exception as exc:
        result["steps"].append({"step": "sort_by_start", "status": "error", "error": repr(exc)})
        return result

    start_dt = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    end_dt = start_dt + timedelta(days=1)
    start_filter = start_dt.strftime("%m/%d/%Y %I:%M %p")
    end_filter = end_dt.strftime("%m/%d/%Y %I:%M %p")
    restriction = f"[Start] < '{end_filter}' AND [End] > '{start_filter}'"

    try:
        restricted = items.Restrict(restriction)
        restricted_count = int(restricted.Count)
        result["steps"].append({
            "step": "restrict_today",
            "status": "ok",
            "filter": restriction,
            "count": restricted_count,
        })
    except Exception as exc:
        result["steps"].append({
            "step": "restrict_today",
            "status": "error",
            "filter": restriction,
            "error": repr(exc),
        })
        return result

    try:
        first = restricted.GetFirst()
        if first is None:
            result["steps"].append({"step": "read_first_event", "status": "ok", "event": None})
        else:
            result["steps"].append({
                "step": "read_first_event",
                "status": "ok",
                "event": {
                    "subject": str(getattr(first, "Subject", "")),
                    "start": str(getattr(first, "Start", "")),
                    "end": str(getattr(first, "End", "")),
                },
            })
    except Exception as exc:
        result["steps"].append({"step": "read_first_event", "status": "error", "error": repr(exc)})
        return result

    result["status"] = "ok"
    return result
