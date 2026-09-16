from __future__ import annotations

import pythoncom
import win32com.client as win32


OL_FOLDER_CALENDAR = 9

win32.gencache.is_readonly = True


def _safe_get(obj, name: str, default=None):
    try:
        return getattr(obj, name)
    except Exception:
        return default


def _serialize_datetime(value):
    if value is None:
        return None
    try:
        return value.isoformat()
    except Exception:
        try:
            return str(value)
        except Exception:
            return None


def diagnose_mailbox_calendar(store_name: str = "Управление программами", sample_limit: int = 5) -> dict:
    """Diagnose read-only access to a calendar in an additional Outlook Store.

    The function never creates, saves, updates, deletes, or sends anything.
    It enumerates Outlook Stores, selects the requested Store by display name,
    opens that Store's default Calendar folder, and reads a small sample of items.
    """
    result: dict = {
        "status": "error",
        "requested_store": store_name,
        "steps": [],
    }

    if sample_limit < 0 or sample_limit > 20:
        raise ValueError("sample_limit must be between 0 and 20")

    try:
        pythoncom.CoInitialize()
        result["steps"].append({"step": "coinitialize", "status": "ok"})

        outlook = win32.Dispatch("Outlook.Application")
        result["steps"].append({"step": "get_outlook", "status": "ok"})

        namespace = outlook.GetNamespace("MAPI")
        result["steps"].append({"step": "get_namespace", "status": "ok"})

        stores = []
        selected_store = None
        requested = store_name.strip().casefold()

        try:
            count = namespace.Stores.Count
        except Exception as exc:
            result["steps"].append({
                "step": "read_stores",
                "status": "blocked",
                "error": repr(exc),
            })
            result["status"] = "stores_blocked"
            return result

        for index in range(1, count + 1):
            try:
                store = namespace.Stores.Item(index)
                display_name = str(_safe_get(store, "DisplayName", ""))
                stores.append(display_name)
                if display_name.strip().casefold() == requested:
                    selected_store = store
            except Exception as exc:
                stores.append(f"<blocked store {index}: {exc!r}>")

        result["steps"].append({
            "step": "read_stores",
            "status": "ok",
            "stores": stores,
        })

        if selected_store is None:
            # Friendly fallback: allow an unambiguous substring match.
            matches = []
            for index in range(1, count + 1):
                try:
                    store = namespace.Stores.Item(index)
                    display_name = str(_safe_get(store, "DisplayName", ""))
                    if requested and requested in display_name.strip().casefold():
                        matches.append(store)
                except Exception:
                    continue
            if len(matches) == 1:
                selected_store = matches[0]

        if selected_store is None:
            result["status"] = "store_not_found"
            result["available_stores"] = stores
            return result

        selected_name = str(_safe_get(selected_store, "DisplayName", store_name))
        result["selected_store"] = selected_name
        result["steps"].append({
            "step": "select_store",
            "status": "ok",
            "store": selected_name,
        })

        try:
            calendar = selected_store.GetDefaultFolder(OL_FOLDER_CALENDAR)
            result["steps"].append({
                "step": "get_calendar_folder",
                "status": "ok",
                "folder_name": str(_safe_get(calendar, "Name", "")),
                "folder_path": str(_safe_get(calendar, "FolderPath", "")),
            })
        except Exception as exc:
            result["steps"].append({
                "step": "get_calendar_folder",
                "status": "blocked",
                "error": repr(exc),
            })
            result["status"] = "calendar_blocked"
            return result

        try:
            items = calendar.Items
            item_count = int(items.Count)
            result["steps"].append({
                "step": "read_calendar_items",
                "status": "ok",
                "count": item_count,
            })
        except Exception as exc:
            result["steps"].append({
                "step": "read_calendar_items",
                "status": "blocked",
                "error": repr(exc),
            })
            result["status"] = "calendar_items_blocked"
            return result

        sample = []
        if sample_limit:
            try:
                items.Sort("[Start]")
            except Exception:
                pass

            try:
                item = items.GetFirst()
            except Exception as exc:
                result["steps"].append({
                    "step": "read_sample",
                    "status": "blocked",
                    "error": repr(exc),
                })
                result["status"] = "calendar_read_supported"
                result["calendar_item_count"] = item_count
                return result

            while item is not None and len(sample) < sample_limit:
                sample.append({
                    "subject": str(_safe_get(item, "Subject", "")),
                    "start": _serialize_datetime(_safe_get(item, "Start")),
                    "end": _serialize_datetime(_safe_get(item, "End")),
                    "location": str(_safe_get(item, "Location", "")),
                })
                try:
                    item = items.GetNext()
                except Exception:
                    break

        result["steps"].append({
            "step": "read_sample",
            "status": "ok",
            "items": sample,
        })
        result["status"] = "calendar_read_supported"
        result["calendar_item_count"] = item_count
        result["sample"] = sample
        return result

    except Exception as exc:
        result["error"] = repr(exc)
        return result
