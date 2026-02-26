import smtplib
import datetime
from pathlib import Path
from typing import List, Optional, Tuple
from email.message import EmailMessage


def _email_list(s: str) -> List[str]:
    """Converts a comma-separated string of emails into a cleaned list."""
    return [x.strip() for x in (s or "").split(",") if x.strip()]


def send_email_outlook(
    subject: str, body: str, to_list: List[str], attachments: List[Path]
) -> Tuple[bool, str]:
    """Sends an email using the local Outlook application."""
    try:
        import win32com.client as win32
        import pythoncom
    except ImportError:
        return False, "Outlook COM (pywin32) is not installed."

    try:
        pythoncom.CoInitialize()
        outlook = win32.gencache.EnsureDispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = "; ".join(to_list)
        mail.Subject = subject
        mail.Body = body
        for p in attachments:
            if p.exists() and not p.name.startswith("~$"):
                mail.Attachments.Add(str(p))
        mail.Send()
        return True, ""
    except Exception as e:
        return False, f"Outlook error: {str(e)}"
    finally:
        pythoncom.CoUninitialize()


def send_email_smtp(
    subject: str, body: str, to_list: List[str], attachments: List[Path], config: dict
) -> Tuple[bool, str]:
    """Sends an email using a standard SMTP server."""
    host = config.get("host")
    port = int(config.get("port", 587))
    if not host:
        return False, "SMTP host is missing."

    try:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = config.get("user") or "noreply@ecovis.hu"
        msg["To"] = ", ".join(to_list)
        msg.set_content(body)

        for p in attachments:
            if p.exists() and not p.name.startswith("~$"):
                msg.add_attachment(
                    p.read_bytes(),
                    maintype="application",
                    subtype="octet-stream",
                    filename=p.name,
                )

        with smtplib.SMTP(host, port, timeout=30) as s:
            if config.get("tls", True):
                s.starttls()
            if config.get("user"):
                s.login(config["user"], config.get("pwd", ""))
            s.send_message(msg)
        return True, ""
    except Exception as e:
        return False, f"SMTP error: {str(e)}"


def send_email(
    subject: str,
    body: str,
    to_list: List[str],
    attachments: List[Path],
    method: str,
    smtp_cfg: dict,
) -> Tuple[bool, str]:
    """Main entry point for sending emails with a fallback mechanism."""
    if method.lower() == "outlook":
        ok, err = send_email_outlook(subject, body, to_list, attachments)
        if ok:
            return True, ""
        # Fallback to SMTP if Outlook fails and host is provided
        if smtp_cfg.get("host"):
            return send_email_smtp(subject, body, to_list, attachments, smtp_cfg)
        return False, err

    return send_email_smtp(subject, body, to_list, attachments, smtp_cfg)
