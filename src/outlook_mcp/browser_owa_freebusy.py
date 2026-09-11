from __future__ import annotations

import json
import os
import random
from datetime import datetime, timedelta
from urllib.parse import quote

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


DEFAULT_CDP_URL = "http://127.0.0.1:9222"
OWA_ORIGIN = "https://mail.sberbank.ru"
OWA_SERVICE_URL = f"{OWA_ORIGIN}/owa/service.svc"
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


def _find_owa_page(browser):
    pages = []
    for context in browser.contexts:
        pages.extend(context.pages)
    for page in pages:
        if page.url.startswith(OWA_ORIGIN):
            return page
    return None


def _capture_canary(page) -> str | None:
    captured: dict[str, str | None] = {"value": None}

    def on_request(request) -> None:
        if captured["value"]:
            return
        if "mail.sberbank.ru/owa/" not in request.url:
            return
        try:
            headers = request.all_headers()
        except Exception:
            return
        for key, value in headers.items():
            if key.lower() == "x-owa-canary" and value:
                captured["value"] = value
                return

    page.on("request", on_request)
    try:
        page.reload(wait_until="domcontentloaded", timeout=20000)
        page.wait_for_timeout(3000)
    except PlaywrightTimeoutError:
        pass
    finally:
        page.remove_listener("request", on_request)

    if captured["value"]:
        return captured["value"]

    # Some OWA deployments also expose the canary as a cookie.
    try:
        for cookie in page.context.cookies():
            if cookie.get("name", "").lower() == "x-owa-canary":
                return cookie.get("value") or None
    except Exception:
        pass
    return None


def diagnose_browser_owa(cdp_url: str | None = None) -> dict:
    endpoint = (cdp_url or os.getenv("OUTLOOK_MCP_CDP_URL") or DEFAULT_CDP_URL).strip()
    result = {
        "status": "error",
        "cdp_url": endpoint,
        "steps": [],
    }

    try:
        with sync_playwright() as p:
            browser = p.chromium.connect_over_cdp(endpoint)
            result["steps"].append({"step": "connect_cdp", "status": "ok"})

            page = _find_owa_page(browser)
            if page is None:
                result["status"] = "owa_tab_not_found"
                result["pages"] = [
                    page.url
                    for context in browser.contexts
                    for page in context.pages
                ][:20]
                return result

            result["steps"].append({
                "step": "find_owa_tab",
                "status": "ok",
                "url": page.url,
            })

            canary = _capture_canary(page)
            result["steps"].append({
                "step": "capture_owa_canary",
                "status": "ok" if canary else "not_found",
            })
            result["status"] = "ok" if canary else "canary_not_found"
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


def get_browser_owa_free_busy(
    emails: list[str],
    start: str,
    end: str,
    slot_minutes: int = 30,
    cdp_url: str | None = None,
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

    endpoint = (cdp_url or os.getenv("OUTLOOK_MCP_CDP_URL") or DEFAULT_CDP_URL).strip()
    payload = _build_payload(clean_emails, start_dt, end_dt, slot_minutes)
    encoded = quote(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), safe="")
    action_id = -random.randint(1000, 999999)
    url = f"{OWA_SERVICE_URL}?action=GetUserAvailabilityInternal&EP=1&ID={action_id}&AC=1"

    try:
        with sync_playwright() as p:
            browser = p.chromium.connect_over_cdp(endpoint)
            page = _find_owa_page(browser)
            if page is None:
                return {
                    "status": "owa_tab_not_found",
                    "cdp_url": endpoint,
                }

            canary = _capture_canary(page)
            if not canary:
                return {
                    "status": "canary_not_found",
                    "cdp_url": endpoint,
                    "owa_url": page.url,
                }

            raw = page.evaluate(
                """async ({url, encoded, actionId, canary}) => {
                    const response = await fetch(url, {
                        method: 'POST',
                        credentials: 'include',
                        redirect: 'manual',
                        headers: {
                            'Accept': '*/*',
                            'Action': 'GetUserAvailabilityInternal',
                            'Content-Type': 'application/json; charset=UTF-8',
                            'X-Requested-With': 'XMLHttpRequest',
                            'X-OWA-ActionId': String(actionId),
                            'X-OWA-ActionName': 'GetUserAvailabilityInternalAction',
                            'X-OWA-Attempt': '1',
                            'X-OWA-CANARY': canary,
                            'X-OWA-UrlPostData': encoded
                        },
                        body: ''
                    });
                    const text = await response.text();
                    return {
                        status: response.status,
                        ok: response.ok,
                        url: response.url,
                        redirected: response.redirected,
                        contentType: response.headers.get('content-type') || '',
                        text
                    };
                }""",
                {
                    "url": url,
                    "encoded": encoded,
                    "actionId": action_id,
                    "canary": canary,
                },
            )

            http_status = int(raw.get("status") or 0)
            if http_status != 200:
                return {
                    "status": "http_error",
                    "http_status": http_status,
                    "browser_mode": "cdp_page_fetch",
                    "response_prefix": str(raw.get("text") or "")[:300],
                    "response_url": raw.get("url"),
                }

            try:
                data = json.loads(raw.get("text") or "")
            except Exception as exc:
                return {
                    "status": "non_json_response",
                    "http_status": http_status,
                    "browser_mode": "cdp_page_fetch",
                    "content_type": raw.get("contentType", ""),
                    "error": repr(exc),
                    "response_prefix": str(raw.get("text") or "")[:300],
                }

            body = data.get("Body") or {}
            if body.get("ResponseCode") != "NoError":
                return {
                    "status": "owa_error",
                    "response_code": body.get("ResponseCode"),
                    "response_class": body.get("ResponseClass"),
                    "browser_mode": "cdp_page_fetch",
                }

            raw_responses = body.get("Responses") or []
            results: list[dict] = []
            for index, email in enumerate(clean_emails):
                entry = raw_responses[index] if index < len(raw_responses) else {}
                calendar_view = entry.get("CalendarView") or {}
                merged = str(calendar_view.get("MergedFreeBusy") or "")
                intervals = _merge_states(start_dt, merged, slot_minutes, end_dt)
                results.append({
                    "email": email,
                    "free_busy_view_type": calendar_view.get("FreeBusyViewType"),
                    "merged_free_busy": merged,
                    "intervals": intervals,
                    "free_intervals": [row for row in intervals if row["status"] == "free"],
                    "working_hours": entry.get("WorkingHours") or {},
                    "items": calendar_view.get("Items") or [],
                })

            return {
                "status": "ok",
                "browser_mode": "cdp_page_fetch",
                "cdp_url": endpoint,
                "start": start_dt.isoformat(),
                "end": end_dt.isoformat(),
                "slot_minutes": slot_minutes,
                "results": results,
            }
    except Exception as exc:
        return {
            "status": "error",
            "browser_mode": "cdp_page_fetch",
            "cdp_url": endpoint,
            "error": repr(exc),
        }
