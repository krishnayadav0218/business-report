"""
email_sender.py
Sends the generated report as an email attachment.

TWO delivery methods -- pick with EMAIL_PROVIDER in your environment:

  EMAIL_PROVIDER=sendgrid  (recommended on Render / most free hosts)
      Sends over HTTPS (port 443) via SendGrid's Web API. Render's free tier
      blocks outbound SMTP ports (25/465/587) to prevent spam abuse, which is
      why plain SMTP times out there ("Could not send email... timed out").
      HTTPS isn't blocked, so this always works. Needs:
        SENDGRID_API_KEY   -- from sendgrid.com (free tier: 100 emails/day)
        SENDGRID_FROM_EMAIL -- an email you verified as a "Single Sender" in SendGrid

  EMAIL_PROVIDER=smtp  (default -- fine for local use / GitHub Actions / your own PC)
      Classic SMTP (Gmail, Outlook, etc.). Works locally and in GitHub Actions,
      but will likely time out on Render's free tier -- use sendgrid there instead.
"""

import os
import socket
import smtplib
import time
import base64
from datetime import date
from email.message import EmailMessage

# --- Force IPv4-only DNS resolution for the SMTP path ---
# Some hosts advertise IPv6 that isn't actually routable, causing slow hangs
# then "Network is unreachable". This keeps SMTP fast where it *is* allowed.
_original_getaddrinfo = socket.getaddrinfo


def _ipv4_only_getaddrinfo(host, port, family=0, type=0, proto=0, flags=0):
    return _original_getaddrinfo(host, port, socket.AF_INET, type, proto, flags)


socket.getaddrinfo = _ipv4_only_getaddrinfo
# ---------------------------------------------------------------

CONNECT_TIMEOUT_SECONDS = 15
MAX_RETRIES = 2


def _build_subject(subject):
    base_subject = subject or "Business Report"
    return f"{base_subject} — {date.today().strftime('%d %b %Y')}"


def _get_recipients(recipients_override):
    if recipients_override is not None:
        return [r.strip() for r in recipients_override if r.strip()]
    return [r.strip() for r in os.getenv("REPORT_RECIPIENTS", "").split(",") if r.strip()]


def send_report(attachment_path: str, subject: str = None, body: str = None,
                 recipients_override: list = None):
    provider = os.getenv("EMAIL_PROVIDER", "smtp").lower()
    recipients = _get_recipients(recipients_override)

    if not recipients:
        raise RuntimeError("No recipients configured -- add at least one email (via the portal or REPORT_RECIPIENTS)")

    if provider == "sendgrid":
        _send_via_sendgrid(attachment_path, recipients, subject, body)
    else:
        _send_via_smtp(attachment_path, recipients, subject, body)


def _send_via_sendgrid(attachment_path, recipients, subject, body):
    import requests

    api_key = os.getenv("SENDGRID_API_KEY")
    from_email = os.getenv("SENDGRID_FROM_EMAIL")
    if not api_key or not from_email:
        raise RuntimeError("SENDGRID_API_KEY / SENDGRID_FROM_EMAIL not set -- check your environment variables")

    with open(attachment_path, "rb") as f:
        encoded_file = base64.b64encode(f.read()).decode()
    filename = os.path.basename(attachment_path)

    payload = {
        "personalizations": [{"to": [{"email": r} for r in recipients]}],
        "from": {"email": from_email},
        "subject": _build_subject(subject),
        "content": [{
            "type": "text/plain",
            "value": body or "Hi,\n\nPlease find the latest automated business report attached.\n\nRegards,\nReporting Bot",
        }],
        "attachments": [{
            "content": encoded_file,
            "filename": filename,
            "type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "disposition": "attachment",
        }],
    }

    last_error = None
    for attempt in range(1, MAX_RETRIES + 2):
        try:
            resp = requests.post(
                "https://api.sendgrid.com/v3/mail/send",
                headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                json=payload,
                timeout=CONNECT_TIMEOUT_SECONDS,
            )
            if resp.status_code in (200, 202):
                print(f"Report emailed via SendGrid to: {', '.join(recipients)}")
                return
            last_error = f"SendGrid returned {resp.status_code}: {resp.text[:300]}"
            print(f"Email attempt {attempt} failed: {last_error}")
        except requests.RequestException as e:
            last_error = str(e)
            print(f"Email attempt {attempt} failed: {last_error}")
        if attempt <= MAX_RETRIES:
            time.sleep(2)

    raise RuntimeError(f"Could not send email via SendGrid after {MAX_RETRIES + 1} attempts: {last_error}")


def _send_via_smtp(attachment_path, recipients, subject, body):
    smtp_host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER")
    smtp_password = os.getenv("SMTP_PASSWORD")

    if not smtp_user or not smtp_password:
        raise RuntimeError("SMTP_USER / SMTP_PASSWORD not set -- check your .env / environment variables")

    msg = EmailMessage()
    msg["Subject"] = _build_subject(subject)
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
            if smtp_port == 465:
                with smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=CONNECT_TIMEOUT_SECONDS) as server:
                    server.login(smtp_user, smtp_password)
                    server.send_message(msg)
            else:
                with smtplib.SMTP(smtp_host, smtp_port, timeout=CONNECT_TIMEOUT_SECONDS) as server:
                    server.starttls()
                    server.login(smtp_user, smtp_password)
                    server.send_message(msg)
            print(f"Report emailed via SMTP to: {', '.join(recipients)}")
            return
        except (smtplib.SMTPException, OSError, socket.timeout) as e:
            last_error = e
            print(f"Email attempt {attempt} failed: {e}")
            if attempt <= MAX_RETRIES:
                time.sleep(2)

    raise RuntimeError(
        f"Could not send email after {MAX_RETRIES + 1} attempts: {last_error}. "
        f"If this is hosted on Render's free tier, outbound SMTP ports are usually blocked -- "
        f"set EMAIL_PROVIDER=sendgrid instead (see README)."
    )
