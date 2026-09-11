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


def diagnose_free_busy(email: str, slot_minutes: int = 30) -> dict:
    """Test CreateRecipient -> Resolve -> FreeBusy without modifying calendar data."""
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

        resolved = bool(recipient.Resolve())
        result["steps"].append({
            "step": "resolve_recipient",
            "status": "ok" if resolved else "not_resolved",
            "resolved": resolved,
            "name": str(getattr(recipient, "Name", "")),
        })
        if not resolved:
            result["status"] = "recipient_not_resolved"
            return result

        start = datetime.now().replace(second=0, microsecond=0)
        raw = str(recipient.FreeBusy(start, slot_minutes, True))
        result["steps"].append({
            "step": "free_busy",
            "status": "ok",
            "slot_minutes": slot_minutes,
            "characters": len(raw),
            "sample": raw[:48],
        })
        result["status"] = "ok"
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

    if not bool(recipient.Resolve()):
        return {
            "status": "recipient_not_resolved",
            "email": email,
            "start": start_dt.isoformat(),
            "end": end_dt.isoformat(),
            "slot_minutes": slot_minutes,
            "slots": [],
            "intervals": [],
        }

    raw = str(recipient.FreeBusy(start_dt, slot_minutes, True))
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
        "resolved_name": str(getattr(recipient, "Name", "")),
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "slot_minutes": slot_minutes,
        "slots": slots,
        "free_slots": free_slots,
        "intervals": intervals,
        "raw_length": len(raw),
    }
