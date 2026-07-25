"""
email_sender.py
Sends the generated report as an email attachment via SMTP
(works with Gmail, Outlook, or SendGrid's SMTP relay).
"""

import os
import smtplib
from email.message import EmailMessage


def send_report(attachment_path: str, subject: str = None, body: str = None):
    smtp_host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER")
    smtp_password = os.getenv("SMTP_PASSWORD")
    recipients = os.getenv("REPORT_RECIPIENTS", "")
    recipients = [r.strip() for r in recipients.split(",") if r.strip()]

    if not smtp_user or not smtp_password:
        raise RuntimeError("SMTP_USER / SMTP_PASSWORD not set -- check your .env file")
    if not recipients:
        raise RuntimeError("REPORT_RECIPIENTS not set -- add at least one email in .env")

    msg = EmailMessage()
    msg["Subject"] = subject or "Weekly Business Report"
    msg["From"] = smtp_user
    msg["To"] = ", ".join(recipients)
    msg.set_content(body or "Hi,\n\nPlease find this week's automated business report attached.\n\nRegards,\nReporting Bot")

    with open(attachment_path, "rb") as f:
        data = f.read()
    filename = os.path.basename(attachment_path)
    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename,
    )

    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)

    print(f"Report emailed to: {', '.join(recipients)}")
