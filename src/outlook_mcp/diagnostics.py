from __future__ import annotations

import pythoncom
import win32com.client


def diagnose_outlook() -> dict:
    """Run step-by-step Outlook COM diagnostics without raising COM errors."""
    result: dict = {"status": "error", "steps": []}

    try:
        pythoncom.CoInitialize()
        result["steps"].append({"step": "coinitialize", "status": "ok"})
    except Exception as exc:
        result["steps"].append({"step": "coinitialize", "status": "error", "error": repr(exc)})
        return result

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
