from __future__ import annotations

from datetime import datetime, timedelta

import pythoncom
import win32com.client as win32


FREE_BUSY_LABELS = {
    "0": "free",
    "1": "tentative",
    "2": "busy",
    "3": "out_of_office",
    "4": "working_elsewhere",
}

win32.gencache.is_readonly = True


def _get_outlook():
    pythoncom.CoInitialize()
    return win32.Dispatch("Outlook.Application")


def _parse_datetime(value: str) -> datetime:
    try:
        return datetime.fromisoformat(value).replace(tzinfo=None)
    except ValueError as exc:
        raise ValueError(
            f"Invalid datetime '{value}'. Use ISO format, for example 2026-09-11T09:00:00"
        ) from exc


def _try_free_busy(recipient, start_dt: datetime, slot_minutes: int) -> tuple[str | None, str | None]:
    """Try FreeBusy directly. Some Outlook/Exchange setups can resolve implicitly here."""
    try:
        return str(recipient.FreeBusy(start_dt, slot_minutes, True)), None
    except Exception as exc:
        return None, repr(exc)


def diagnose_free_busy(email: str, slot_minutes: int = 30) -> dict:
    """Test CreateRecipient, Resolve and direct FreeBusy without modifying calendar data."""
    if not email.strip():
        raise ValueError("email is required")
    if slot_minutes < 5 or slot_minutes > 1440:
        raise ValueError("slot_minutes must be between 5 and 1440")

    result: dict = {"status": "error", "email": email, "steps": []}

    try:
        outlook = _get_outlook()
        result["steps"].append({"step": "get_outlook", "status": "ok"})

        namespace = outlook.GetNamespace("MAPI")
        result["steps"].append({"step": "get_mapi_namespace", "status": "ok"})

        recipient = namespace.CreateRecipient(email)
        result["steps"].append({"step": "create_recipient", "status": "ok"})

        resolved = False
        try:
            resolved = bool(recipient.Resolve())
            result["steps"].append({
                "step": "resolve_recipient",
                "status": "ok" if resolved else "not_resolved",
                "resolved": resolved,
                "name": str(getattr(recipient, "Name", "")),
            })
        except Exception as exc:
            result["steps"].append({
                "step": "resolve_recipient",
                "status": "blocked",
                "error": repr(exc),
            })

        start = datetime.now().replace(second=0, microsecond=0)
        raw, free_busy_error = _try_free_busy(recipient, start, slot_minutes)
        if raw is not None:
            result["steps"].append({
                "step": "free_busy",
                "status": "ok",
                "slot_minutes": slot_minutes,
                "characters": len(raw),
                "sample": raw[:48],
                "implicit_resolution": not resolved,
            })
            result["status"] = "ok"
            result["mode"] = "resolved_recipient" if resolved else "free_busy_without_explicit_resolve"
            return result

        result["steps"].append({
            "step": "free_busy",
            "status": "blocked",
            "error": free_busy_error,
            "attempted_without_explicit_resolve": not resolved,
        })
        result["status"] = "free_busy_blocked"
        result["error"] = free_busy_error
        return result

    except Exception as exc:
        result["error"] = repr(exc)
        return result


def _merge_slots(start_dt: datetime, states: str, slot_minutes: int) -> list[dict]:
    if not states:
        return []

    merged: list[dict] = []
    run_state = states[0]
    run_start_index = 0

    def append_run(end_index: int) -> None:
        run_start = start_dt + timedelta(minutes=run_start_index * slot_minutes)
        run_end = start_dt + timedelta(minutes=end_index * slot_minutes)
        merged.append({
            "start": run_start.isoformat(),
            "end": run_end.isoformat(),
            "status": FREE_BUSY_LABELS.get(run_state, "unknown"),
            "code": run_state,
        })

    for index, state in enumerate(states[1:], start=1):
        if state != run_state:
            append_run(index)
            run_state = state
            run_start_index = index

    append_run(len(states))
    return merged


def get_employee_free_busy(
    email: str,
    start: str,
    end: str,
    slot_minutes: int = 30,
) -> dict:
    """Return Exchange/Outlook free-busy information for one employee."""
    if not email.strip():
        raise ValueError("email is required")
    if slot_minutes < 5 or slot_minutes > 1440:
        raise ValueError("slot_minutes must be between 5 and 1440")

    start_dt = _parse_datetime(start)
    end_dt = _parse_datetime(end)
    if end_dt <= start_dt:
        raise ValueError("end must be later than start")

    requested_minutes = int((end_dt - start_dt).total_seconds() // 60)
    slots_needed = (requested_minutes + slot_minutes - 1) // slot_minutes
    if slots_needed > 1440:
        raise ValueError("requested range is too large")

    outlook = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")
    recipient = namespace.CreateRecipient(email)

    resolved = False
    resolve_error = None
    try:
        resolved = bool(recipient.Resolve())
    except Exception as exc:
        resolve_error = repr(exc)

    raw, free_busy_error = _try_free_busy(recipient, start_dt, slot_minutes)
    if raw is None:
        return {
            "status": "free_busy_blocked",
            "email": email,
            "start": start_dt.isoformat(),
            "end": end_dt.isoformat(),
            "slot_minutes": slot_minutes,
            "resolved": resolved,
            "resolve_error": resolve_error,
            "free_busy_error": free_busy_error,
            "slots": [],
            "free_slots": [],
            "intervals": [],
        }

    states = raw[:slots_needed]

    slots: list[dict] = []
    free_slots: list[dict] = []
    for index, state in enumerate(states):
        slot_start = start_dt + timedelta(minutes=index * slot_minutes)
        slot_end = min(slot_start + timedelta(minutes=slot_minutes), end_dt)
        row = {
            "start": slot_start.isoformat(),
            "end": slot_end.isoformat(),
            "status": FREE_BUSY_LABELS.get(state, "unknown"),
            "code": state,
        }
        slots.append(row)
        if state == "0":
            free_slots.append(row)

    intervals = _merge_slots(start_dt, states, slot_minutes)
    if intervals and intervals[-1]["end"] > end_dt.isoformat():
        intervals[-1]["end"] = end_dt.isoformat()

    return {
        "status": "ok",
        "email": email,
        "resolved": resolved,
        "resolve_error": resolve_error,
        "mode": "resolved_recipient" if resolved else "free_busy_without_explicit_resolve",
        "resolved_name": str(getattr(recipient, "Name", "")) if resolved else "",
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "slot_minutes": slot_minutes,
        "slots": slots,
        "free_slots": free_slots,
        "intervals": intervals,
        "raw_length": len(raw),
    }
