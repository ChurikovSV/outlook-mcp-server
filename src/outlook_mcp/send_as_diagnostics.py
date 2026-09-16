from __future__ import annotations

import pythoncom
import win32com.client


def diagnose_send_as_account(email: str) -> dict:
    """Check whether Outlook can prepare a draft using another mailbox identity.

    This diagnostic is intentionally best-effort: corporate Outlook policy may block
    account enumeration or recipient resolution even when setting SentOnBehalfOfName
    on a MailItem still works. No message is sent and no persistent draft is saved.
    """
    result = {
        "status": "error",
        "email": email,
        "steps": [],
    }

    try:
        pythoncom.CoInitialize()
        result["steps"].append({"step": "coinitialize", "status": "ok"})
    except Exception as exc:
        result["steps"].append({"step": "coinitialize", "status": "error", "error": repr(exc)})
        result["error"] = repr(exc)
        return result

    try:
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
            result["steps"].append({"step": "get_outlook", "status": "ok", "mode": "active_object"})
        except Exception as active_exc:
            result["steps"].append({
                "step": "get_active_outlook",
                "status": "not_available",
                "error": repr(active_exc),
            })
            outlook = win32com.client.Dispatch("Outlook.Application")
            result["steps"].append({"step": "get_outlook", "status": "ok", "mode": "dispatch"})
    except Exception as exc:
        result["steps"].append({"step": "get_outlook", "status": "error", "error": repr(exc)})
        result["error"] = repr(exc)
        return result

    namespace = None
    try:
        namespace = outlook.GetNamespace("MAPI")
        result["steps"].append({"step": "get_namespace", "status": "ok"})
    except Exception as exc:
        result["steps"].append({"step": "get_namespace", "status": "error", "error": repr(exc)})

    if namespace is not None:
        try:
            accounts = []
            for account in namespace.Accounts:
                try:
                    smtp = str(getattr(account, "SmtpAddress", ""))
                except Exception as exc:
                    smtp = f"<blocked: {exc!r}>"
                try:
                    name = str(getattr(account, "DisplayName", ""))
                except Exception as exc:
                    name = f"<blocked: {exc!r}>"
                accounts.append({"smtp": smtp, "name": name})
            result["steps"].append({"step": "read_accounts", "status": "ok", "accounts": accounts})
        except Exception as exc:
            result["steps"].append({"step": "read_accounts", "status": "blocked", "error": repr(exc)})

        try:
            stores = []
            for store in namespace.Stores:
                try:
                    stores.append(str(getattr(store, "DisplayName", "")))
                except Exception as exc:
                    stores.append(f"<blocked: {exc!r}>")
            result["steps"].append({"step": "read_stores", "status": "ok", "stores": stores})
        except Exception as exc:
            result["steps"].append({"step": "read_stores", "status": "blocked", "error": repr(exc)})

        try:
            recipient = namespace.CreateRecipient(email)
            result["steps"].append({"step": "create_recipient", "status": "ok"})
            try:
                resolved = bool(recipient.Resolve())
                result["steps"].append({
                    "step": "resolve_recipient",
                    "status": "ok",
                    "resolved": resolved,
                })
            except Exception as exc:
                result["steps"].append({
                    "step": "resolve_recipient",
                    "status": "blocked",
                    "error": repr(exc),
                })
        except Exception as exc:
            result["steps"].append({"step": "create_recipient", "status": "blocked", "error": repr(exc)})

    try:
        mail = outlook.CreateItem(0)
        result["steps"].append({"step": "create_mail_item", "status": "ok"})
    except Exception as exc:
        result["steps"].append({"step": "create_mail_item", "status": "error", "error": repr(exc)})
        result["error"] = repr(exc)
        return result

    try:
        mail.SentOnBehalfOfName = email
        result["steps"].append({"step": "set_sent_on_behalf_of", "status": "ok"})
        result["status"] = "supported"
    except Exception as exc:
        result["steps"].append({
            "step": "set_sent_on_behalf_of",
            "status": "blocked",
            "error": repr(exc),
        })
        result["status"] = "blocked"

    return result
