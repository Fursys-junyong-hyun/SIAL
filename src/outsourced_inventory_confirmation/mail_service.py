from __future__ import annotations

import re
import smtplib
from datetime import date, timedelta
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from pathlib import Path

import pythoncom
import win32com.client

from .mail_models import MailSettings, PreparedEmail
from .mail_oauth import build_xoauth2_string, refresh_access_token

WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]
EMAIL_PATTERN = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")


def is_valid_email(value: str) -> bool:
    return bool(EMAIL_PATTERN.match(value.strip())) if value else False


def suggested_reply_due_date(base: date | None = None) -> date:
    """다음 주 금요일을 기본 회신 기한으로 제안한다."""
    base = base or date.today()
    days_until_friday = (4 - base.weekday()) % 7
    if days_until_friday < 3:
        days_until_friday += 7
    return base + timedelta(days=days_until_friday)


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


def _format_from(settings: MailSettings, fallback_email: str = "") -> str:
    address = settings.smtp_username or settings.oauth_account_email or fallback_email
    if settings.sender_name and address:
        return formataddr((settings.sender_name, address))
    return address


def _build_message(settings: MailSettings, email: PreparedEmail) -> EmailMessage:
    message = EmailMessage()
    from_value = _format_from(settings)
    if from_value:
        message["From"] = from_value
    message["To"] = email.recipient
    message["Subject"] = email.subject
    message["Date"] = formatdate(localtime=True)
    message["Message-ID"] = make_msgid()
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
    return message


def send_via_gmail(settings: MailSettings, password: str, email: PreparedEmail) -> None:
    """Gmail SMTP + 앱 비밀번호 발송 (회사 정책상 차단된 환경에서는 OAuth/EML 사용)."""
    message = _build_message(settings, email)
    with smtplib.SMTP(settings.smtp_host, settings.smtp_port, timeout=30) as smtp:
        if settings.smtp_use_tls:
            smtp.starttls()
        smtp.login(settings.smtp_username, password)
        smtp.send_message(message)


def send_via_gmail_oauth(
    settings: MailSettings,
    refresh_token: str,
    client_id: str,
    client_secret: str,
    email: PreparedEmail,
    access_token: str | None = None,
) -> str:
    """Gmail SMTP + XOAUTH2 발송. 앱 비밀번호 차단 환경의 정공법.

    재발급된 access token 을 반환해 다음 호출 때 재사용할 수 있다.
    """
    if access_token is None:
        access_token = refresh_access_token(client_id, client_secret, refresh_token)

    account_email = settings.oauth_account_email or settings.smtp_username
    if not account_email:
        raise RuntimeError("OAuth 로 로그인된 Gmail 주소가 없습니다. 다시 로그인해 주세요.")

    auth_string = build_xoauth2_string(account_email, access_token)
    message = _build_message(settings, email)

    with smtplib.SMTP(settings.smtp_host, settings.smtp_port, timeout=30) as smtp:
        smtp.ehlo()
        if settings.smtp_use_tls:
            smtp.starttls()
            smtp.ehlo()
        code, response = smtp.docmd("AUTH", f"XOAUTH2 {auth_string}")
        if code != 235:
            raise RuntimeError(f"XOAUTH2 인증 실패 ({code}): {response!r}")
        smtp.send_message(message)
    return access_token


def save_as_eml(settings: MailSettings, email: PreparedEmail, output_path: Path) -> Path:
    """발송 대신 Outlook 등에서 더블클릭으로 열 수 있는 .eml 초안 파일로 저장한다."""
    message = _build_message(settings, email)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(bytes(message))
    return output_path


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


def verify_gmail_credentials(settings: MailSettings, password: str) -> None:
    """SMTP 로그인까지만 시도해 자격 증명을 검증한다. 실패 시 예외를 그대로 전파한다."""
    with smtplib.SMTP(settings.smtp_host, settings.smtp_port, timeout=15) as smtp:
        if settings.smtp_use_tls:
            smtp.starttls()
        smtp.login(settings.smtp_username, password)
