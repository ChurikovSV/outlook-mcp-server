from __future__ import annotations

import json
import os
import random
from datetime import datetime, timedelta
from urllib.parse import quote

import requests
import truststore
from requests_negotiate_sspi import HttpNegotiateAuth


OWA_SERVICE_URL = "https://mail.sberbank.ru/owa/service.svc"
FREE_BUSY_LABELS = {
    "0": "free",
    "1": "tentative",
    "2": "busy",
    "3": "out_of_office",
    "4": "working_elsewhere",
}

# requests normally validates TLS against certifi's CA bundle. Corporate OWA
# certificates are often signed by an internal enterprise CA that Windows trusts
# but certifi does not know about. Inject the native Windows trust store into
# Python's ssl module instead of disabling certificate verification.
truststore.inject_into_ssl()


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


def _auth_config() -> tuple[str, object | None, dict[str, str]]:
    """Choose OWA authentication without persisting browser secrets in source code.

    If OUTLOOK_MCP_OWA_COOKIE and OUTLOOK_MCP_OWA_CANARY are present, reuse the
    user's active OWA browser session. Otherwise try Windows Integrated Auth.
    """
    cookie = os.getenv("OUTLOOK_MCP_OWA_COOKIE", "").strip()
    canary = os.getenv("OUTLOOK_MCP_OWA_CANARY", "").strip()

    if cookie and canary:
        return (
            "browser_session_env",
            None,
            {
                "Cookie": cookie,
                "X-OWA-CANARY": canary,
                "Origin": "https://mail.sberbank.ru",
            },
        )

    return "windows_integrated_auth", HttpNegotiateAuth(), {}


def _make_request(emails: list[str], start_dt: datetime, end_dt: datetime, slot_minutes: int):
    payload = _build_payload(emails, start_dt, end_dt, slot_minutes)
    encoded = quote(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), safe="")
    action_id = -random.randint(1000, 999999)
    url = f"{OWA_SERVICE_URL}?action=GetUserAvailabilityInternal&EP=1&ID={action_id}&AC=1"
    auth_mode, auth, auth_headers = _auth_config()
    headers = {
        "Accept": "*/*",
        "Action": "GetUserAvailabilityInternal",
        "Content-Type": "application/json; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "X-OWA-ActionId": str(action_id),
        "X-OWA-ActionName": "GetUserAvailabilityInternalAction",
        "X-OWA-Attempt": "1",
        "X-OWA-UrlPostData": encoded,
        **auth_headers,
    }
    response = requests.post(
        url,
        headers=headers,
        data=b"",
        auth=auth,
        timeout=20,
        allow_redirects=False,
    )
    return response, payload, auth_mode


def diagnose_owa_free_busy(email: str) -> dict:
    if not email.strip():
        raise ValueError("email is required")

    start_dt = datetime.now().replace(minute=0, second=0, microsecond=0)
    end_dt = start_dt + timedelta(hours=2)
    result = {
        "status": "error",
        "email": email,
        "tls_trust": "windows_system_store",
        "steps": [],
    }

    try:
        result["steps"].append({"step": "build_request", "status": "ok"})
        response, _, auth_mode = _make_request([email], start_dt, end_dt, 30)
        result["auth_mode"] = auth_mode
        result["steps"].append({
            "step": "http_post",
            "status": "ok",
            "http_status": response.status_code,
            "content_type": response.headers.get("Content-Type", ""),
            "location": response.headers.get("Location", ""),
        })

        if 300 <= response.status_code < 400:
            result["status"] = "authentication_redirect"
            return result

        if response.status_code in (401, 403):
            result["status"] = "authentication_required"
            return result

        if response.status_code >= 400:
            result["status"] = "http_error"
            result["response_prefix"] = response.text[:300]
            return result

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

    response, payload, auth_mode = _make_request(clean_emails, start_dt, end_dt, slot_minutes)

    if 300 <= response.status_code < 400:
        return {
            "status": "authentication_redirect",
            "http_status": response.status_code,
            "location": response.headers.get("Location", ""),
            "auth_mode": auth_mode,
            "tls_trust": "windows_system_store",
            "emails": clean_emails,
            "start": start_dt.isoformat(),
            "end": end_dt.isoformat(),
            "slot_minutes": slot_minutes,
        }

    if response.status_code in (401, 403):
        return {
            "status": "authentication_required",
            "http_status": response.status_code,
            "auth_mode": auth_mode,
            "tls_trust": "windows_system_store",
            "emails": clean_emails,
            "start": start_dt.isoformat(),
            "end": end_dt.isoformat(),
            "slot_minutes": slot_minutes,
        }

    if response.status_code >= 400:
        return {
            "status": "http_error",
            "http_status": response.status_code,
            "auth_mode": auth_mode,
            "tls_trust": "windows_system_store",
            "response_prefix": response.text[:300],
            "emails": clean_emails,
        }

    try:
        data = response.json()
    except Exception as exc:
        return {
            "status": "non_json_response",
            "http_status": response.status_code,
            "auth_mode": auth_mode,
            "tls_trust": "windows_system_store",
            "error": repr(exc),
            "response_prefix": response.text[:300],
            "emails": clean_emails,
        }

    body = data.get("Body") or {}
    if body.get("ResponseCode") != "NoError":
        return {
            "status": "owa_error",
            "response_code": body.get("ResponseCode"),
            "response_class": body.get("ResponseClass"),
            "auth_mode": auth_mode,
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
        "auth_mode": auth_mode,
        "tls_trust": "windows_system_store",
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "slot_minutes": slot_minutes,
        "results": results,
        "request_shape": payload,
    }
