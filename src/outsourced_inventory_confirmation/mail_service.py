from __future__ import annotations

import smtplib
from datetime import date
from email.message import EmailMessage

import pythoncom
import win32com.client

from .mail_models import MailSettings, PreparedEmail

WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]


def format_report_date_kr(report_date: date) -> str:
    return f"{report_date:%y}년 {report_date:%m}월 {report_date:%d}일"


def quarter_text(report_date: date) -> str:
    return f"{((report_date.month - 1) // 3) + 1}분기"


def quarter_month_range(report_date: date) -> tuple[int, int]:
    quarter = ((report_date.month - 1) // 3) + 1
    start_month = (quarter - 1) * 3 + 1
    end_month = start_month + 2
    return start_month, end_month


def reply_due_text(reply_due_date: date) -> str:
    return f"{reply_due_date:%m/%d}({WEEKDAYS[reply_due_date.weekday()]})"


def build_template_context(
    vendor_name: str,
    report_date: date,
    reply_due_date: date,
    sender_name: str,
) -> dict[str, str]:
    start_month, end_month = quarter_month_range(report_date)
    return {
        "vendor_name": vendor_name,
        "sender_name": sender_name,
        "quarter": quarter_text(report_date),
        "period_year": f"{report_date:%Y}",
        "period_start_month": f"{start_month:02d}",
        "period_end_month": f"{end_month:02d}",
        "report_date_kr": format_report_date_kr(report_date),
        "reply_due_kr": reply_due_text(reply_due_date),
    }


def render_subject(template: str, context: dict[str, str]) -> str:
    return template.format(**context)


def render_body(template: str, context: dict[str, str]) -> str:
    return template.format(**context)


def send_via_gmail(settings: MailSettings, password: str, email: PreparedEmail) -> None:
    message = EmailMessage()
    message["From"] = settings.smtp_username
    message["To"] = email.recipient
    message["Subject"] = email.subject
    message.set_content(email.body)

    for attachment in email.attachments:
        with attachment.open("rb") as file:
            data = file.read()
        maintype = "application"
        subtype = "octet-stream"
        if attachment.suffix.lower() == ".pdf":
            subtype = "pdf"
        elif attachment.suffix.lower() in {".xlsx", ".xls"}:
            subtype = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        message.add_attachment(data, maintype=maintype, subtype=subtype, filename=attachment.name)

    with smtplib.SMTP(settings.smtp_host, settings.smtp_port, timeout=30) as smtp:
        if settings.smtp_use_tls:
            smtp.starttls()
        smtp.login(settings.smtp_username, password)
        smtp.send_message(message)


def send_via_outlook(email: PreparedEmail) -> None:
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.To = email.recipient
        message.Subject = email.subject
        message.Body = email.body
        for attachment in email.attachments:
            message.Attachments.Add(str(attachment))
        message.Send()
    finally:
        pythoncom.CoUninitialize()
