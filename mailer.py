"""
Email delivery for the generated report.
Reads SMTP credentials from Streamlit secrets (see .streamlit/secrets.toml.example).
"""

import mimetypes
import smtplib
from email.message import EmailMessage

import streamlit as st


class EmailNotConfigured(Exception):
    pass


def _smtp_config():
    try:
        cfg = st.secrets["smtp"]
        return {
            "server": cfg["server"],
            "port": int(cfg.get("port", 587)),
            "username": cfg["username"],
            "password": cfg["password"],
            "sender": cfg.get("sender", cfg["username"]),
        }
    except (KeyError, FileNotFoundError) as exc:
        raise EmailNotConfigured(
            "SMTP is not configured. Add an [smtp] section to Streamlit secrets "
            "(server, port, username, password) — see .streamlit/secrets.toml.example."
        ) from exc


def send_report_email(to_addrs, subject, body, attachment_bytes, attachment_filename):
    """to_addrs: str or list[str] of recipient email addresses."""
    cfg = _smtp_config()

    if isinstance(to_addrs, str):
        to_addrs = [a.strip() for a in to_addrs.split(",") if a.strip()]

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = cfg["sender"]
    msg["To"] = ", ".join(to_addrs)
    msg.set_content(body)

    mime_type, _ = mimetypes.guess_type(attachment_filename)
    maintype, subtype = (mime_type or "application/octet-stream").split("/", 1)
    msg.add_attachment(
        attachment_bytes,
        maintype=maintype,
        subtype=subtype,
        filename=attachment_filename,
    )

    with smtplib.SMTP(cfg["server"], cfg["port"]) as smtp:
        smtp.starttls()
        smtp.login(cfg["username"], cfg["password"])
        smtp.send_message(msg)

    return to_addrs
