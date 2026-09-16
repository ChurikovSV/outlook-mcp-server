from __future__ import annotations

import pythoncom
import win32com.client


def diagnose_send_as_account(email: str) -> dict:
    """Check whether Outlook can create a draft using another mailbox identity.

    Does not send mail and does not save a persistent draft.
    """
    result = {
        "status": "error",
        "email": email,
        "steps": [],
    }

    try:
        pythoncom.CoInitialize()
        result["steps"].append({"step": "coinitialize", "status": "ok"})

        outlook = win32com.client.GetActiveObject("Outlook.Application")
        result["steps"].append({"step": "get_outlook", "status": "ok"})

        namespace = outlook.GetNamespace("MAPI")
        result["steps"].append({"step": "get_namespace", "status": "ok"})

        accounts = []
        for account in namespace.Accounts:
            accounts.append({
                "smtp": str(getattr(account, "SmtpAddress", "")),
                "name": str(getattr(account, "DisplayName", "")),
            })
        result["steps"].append({"step": "read_accounts", "status": "ok", "accounts": accounts})

        recipient = namespace.CreateRecipient(email)
        result["steps"].append({"step": "create_recipient", "status": "ok"})

        try:
            resolved = bool(recipient.Resolve())
            result["steps"].append({"step": "resolve_recipient", "status": "ok", "resolved": resolved})
        except Exception as exc:
            result["steps"].append({"step": "resolve_recipient", "status": "error", "error": repr(exc)})

        mail = outlook.CreateItem(0)
        result["steps"].append({"step": "create_mail_item", "status": "ok"})

        try:
            mail.SentOnBehalfOfName = email
            result["steps"].append({"step": "set_sent_on_behalf_of", "status": "ok"})
        except Exception as exc:
            result["steps"].append({"step": "set_sent_on_behalf_of", "status": "error", "error": repr(exc)})

        result["status"] = "ok"
        return result

    except Exception as exc:
        result["error"] = repr(exc)
        return result
