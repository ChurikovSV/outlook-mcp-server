from __future__ import annotations

import json
import random
from datetime import datetime, timedelta
from urllib.parse import quote

import requests
from requests_negotiate_sspi import HttpNegotiateAuth


OWA_SERVICE_URL = "https://mail.sberbank.ru/owa/service.svc"
FREE_BUSY_LABELS = {
    "0": "free",
    "1": "tentative",
    "2": "busy",
    "3": "out_of_office",
    "4": "working_elsewhere",
}


def _parse_datetime(value: str) -> datetime:
    try:
        return datetime.fromisoformat(value).replace(tzinfo=None)
    except ValueError as exc:
        raise ValueError(
            f"Invalid datetime '{value}'. Use ISO format, for example 2026-09-11T09:00:00"
        ) from exc


def _build_payload(emails: list[str], start_dt: datetime, end_dt: datetime, slot_minutes: int) -> dict:
    return {
        "request": {
            "__type": "GetUserAvailabilityInternalJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "Exchange2013",
                "TimeZoneContext": {
                    "__type": "TimeZoneContext:#Exchange",
                    "TimeZoneDefinition": {
                        "__type": "TimeZoneDefinitionType:#Exchange",
                        "Id": "Russian Standard Time",
                    },
                },
            },
            "Body": {
                "__type": "GetUserAvailabilityRequest:#Exchange",
                "MailboxDataArray": [
                    {
                        "__type": "MailboxData:#Exchange",
                        "Email": {
                            "__type": "EmailAddress:#Exchange",
                            "Address": email,
                        },
                    }
                    for email in emails
                ],
                "FreeBusyViewOptions": {
                    "__type": "FreeBusyViewOptions:#Exchange",
                    "MergedFreeBusyIntervalInMinutes": slot_minutes,
                    "RequestedView": "DetailedMerged",
                    "TimeWindow": {
                        "__type": "Duration:#Exchange",
                        "StartTime": start_dt.isoformat(timespec="seconds"),
                        "EndTime": end_dt.isoformat(timespec="seconds"),
                    },
                },
            },
        }
    }


def _make_request(emails: list[str], start_dt: datetime, end_dt: datetime, slot_minutes: int):
    payload = _build_payload(emails, start_dt, end_dt, slot_minutes)
    encoded = quote(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), safe="")
    action_id = -random.randint(1000, 999999)
    url = f"{OWA_SERVICE_URL}?action=GetUserAvailabilityInternal&EP=1&ID={action_id}&AC=1"
    headers = {
        "Accept": "*/*",
        "Action": "GetUserAvailabilityInternal",
        "Content-Type": "application/json; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "X-OWA-ActionId": str(action_id),
        "X-OWA-ActionName": "GetUserAvailabilityInternalAction",
        "X-OWA-Attempt": "1",
        "X-OWA-UrlPostData": encoded,
    }
    response = requests.post(
        url,
        headers=headers,
        data=b"",
        auth=HttpNegotiateAuth(),
        timeout=20,
    )
    return response, payload


def diagnose_owa_free_busy(email: str) -> dict:
    if not email.strip():
        raise ValueError("email is required")

    start_dt = datetime.now().replace(minute=0, second=0, microsecond=0)
    end_dt = start_dt + timedelta(hours=2)
    result = {
        "status": "error",
        "email": email,
        "auth_mode": "windows_integrated_auth",
        "steps": [],
    }

    try:
        result["steps"].append({"step": "build_request", "status": "ok"})
        response, _ = _make_request([email], start_dt, end_dt, 30)
        result["steps"].append({
            "step": "http_post",
            "status": "ok",
            "http_status": response.status_code,
            "content_type": response.headers.get("Content-Type", ""),
        })

        if response.status_code in (401, 403):
            result["status"] = "authentication_required"
            return result

        response.raise_for_status()
        try:
            data = response.json()
        except Exception as exc:
            result["status"] = "non_json_response"
            result["error"] = repr(exc)
            result["response_prefix"] = response.text[:300]
            return result

        body = data.get("Body") or {}
        result["steps"].append({
            "step": "parse_response",
            "status": "ok",
            "response_code": body.get("ResponseCode"),
            "response_class": body.get("ResponseClass"),
            "responses": len(body.get("Responses") or []),
        })
        result["status"] = "ok" if body.get("ResponseCode") == "NoError" else "owa_error"
        return result
    except Exception as exc:
        result["error"] = repr(exc)
        return result


def _merge_states(start_dt: datetime, states: str, slot_minutes: int, end_dt: datetime) -> list[dict]:
    if not states:
        return []

    intervals: list[dict] = []
    run_state = states[0]
    run_start = 0

    def append_run(end_index: int) -> None:
        start = start_dt + timedelta(minutes=run_start * slot_minutes)
        end = min(start_dt + timedelta(minutes=end_index * slot_minutes), end_dt)
        intervals.append({
            "start": start.isoformat(),
            "end": end.isoformat(),
            "status": FREE_BUSY_LABELS.get(run_state, "unknown"),
            "code": run_state,
        })

    for index, state in enumerate(states[1:], start=1):
        if state != run_state:
            append_run(index)
            run_state = state
            run_start = index
    append_run(len(states))
    return intervals


def get_owa_free_busy(
    emails: list[str],
    start: str,
    end: str,
    slot_minutes: int = 30,
) -> dict:
    clean_emails = [email.strip() for email in emails if email and email.strip()]
    if not clean_emails:
        raise ValueError("at least one email is required")
    if slot_minutes < 5 or slot_minutes > 1440:
        raise ValueError("slot_minutes must be between 5 and 1440")

    start_dt = _parse_datetime(start)
    end_dt = _parse_datetime(end)
    if end_dt <= start_dt:
        raise ValueError("end must be later than start")

    response, payload = _make_request(clean_emails, start_dt, end_dt, slot_minutes)
    if response.status_code in (401, 403):
        return {
            "status": "authentication_required",
            "http_status": response.status_code,
            "auth_mode": "windows_integrated_auth",
            "emails": clean_emails,
            "start": start_dt.isoformat(),
            "end": end_dt.isoformat(),
            "slot_minutes": slot_minutes,
        }

    response.raise_for_status()
    data = response.json()
    body = data.get("Body") or {}
    if body.get("ResponseCode") != "NoError":
        return {
            "status": "owa_error",
            "response_code": body.get("ResponseCode"),
            "response_class": body.get("ResponseClass"),
            "emails": clean_emails,
        }

    raw_responses = body.get("Responses") or []
    results: list[dict] = []
    for index, email in enumerate(clean_emails):
        entry = raw_responses[index] if index < len(raw_responses) else {}
        calendar_view = entry.get("CalendarView") or {}
        merged = str(calendar_view.get("MergedFreeBusy") or "")
        intervals = _merge_states(start_dt, merged, slot_minutes, end_dt)
        working_hours = entry.get("WorkingHours") or {}
        results.append({
            "email": email,
            "free_busy_view_type": calendar_view.get("FreeBusyViewType"),
            "merged_free_busy": merged,
            "intervals": intervals,
            "free_intervals": [row for row in intervals if row["status"] == "free"],
            "working_hours": working_hours,
            "items": calendar_view.get("Items") or [],
        })

    return {
        "status": "ok",
        "auth_mode": "windows_integrated_auth",
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "slot_minutes": slot_minutes,
        "results": results,
        "request_shape": payload,
    }
