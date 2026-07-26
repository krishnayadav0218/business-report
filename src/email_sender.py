"""
email_sender.py
Sends the generated report as an email attachment via SMTP
(works with Gmail, Outlook, or SendGrid's SMTP relay).

IPv4 FIX: Render's free tier (and several other free hosts) advertise IPv6
connectivity that isn't actually routable outbound. Python's smtplib resolves
smtp.gmail.com to an IPv6 address first and tries that, which hangs for a long
time and then fails with "[Errno 101] Network is unreachable". Forcing DNS
resolution to IPv4-only fixes this and also makes the connection fast again.
"""

import os
import socket
import smtplib
import time
from datetime import date
from email.message import EmailMessage

# --- Force IPv4-only DNS resolution (see module docstring) ---
_original_getaddrinfo = socket.getaddrinfo


def _ipv4_only_getaddrinfo(host, port, family=0, type=0, proto=0, flags=0):
    return _original_getaddrinfo(host, port, socket.AF_INET, type, proto, flags)


socket.getaddrinfo = _ipv4_only_getaddrinfo
# ---------------------------------------------------------------

CONNECT_TIMEOUT_SECONDS = 15  # fail fast instead of hanging
MAX_RETRIES = 2


def send_report(attachment_path: str, subject: str = None, body: str = None,
                 recipients_override: list = None):
    smtp_host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER")
    smtp_password = os.getenv("SMTP_PASSWORD")

    if recipients_override is not None:
        recipients = [r.strip() for r in recipients_override if r.strip()]
    else:
        recipients = [r.strip() for r in os.getenv("REPORT_RECIPIENTS", "").split(",") if r.strip()]

    if not smtp_user or not smtp_password:
        raise RuntimeError("SMTP_USER / SMTP_PASSWORD not set -- check your .env / environment variables")
    if not recipients:
        raise RuntimeError("No recipients configured -- add at least one email (via the portal or REPORT_RECIPIENTS)")

    base_subject = subject or "Business Report"
    dated_subject = f"{base_subject} — {date.today().strftime('%d %b %Y')}"

    msg = EmailMessage()
    msg["Subject"] = dated_subject
    msg["From"] = smtp_user
    msg["To"] = ", ".join(recipients)
    msg.set_content(body or "Hi,\n\nPlease find the latest automated business report attached.\n\nRegards,\nReporting Bot")

    with open(attachment_path, "rb") as f:
        data = f.read()
    filename = os.path.basename(attachment_path)
    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename,
    )

    last_error = None
    for attempt in range(1, MAX_RETRIES + 2):
        try:
            with smtplib.SMTP(smtp_host, smtp_port, timeout=CONNECT_TIMEOUT_SECONDS) as server:
                server.starttls()
                server.login(smtp_user, smtp_password)
                server.send_message(msg)
            print(f"Report emailed to: {', '.join(recipients)}")
            return
        except (smtplib.SMTPException, OSError, socket.timeout) as e:
            last_error = e
            print(f"Email attempt {attempt} failed: {e}")
            if attempt <= MAX_RETRIES:
                time.sleep(2)

    raise RuntimeError(f"Could not send email after {MAX_RETRIES + 1} attempts: {last_error}")
