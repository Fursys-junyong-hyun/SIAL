from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class MailSettings:
    sender_name: str = ""
    default_method: str = "outlook"
    smtp_host: str = "smtp.gmail.com"
    smtp_port: int = 587
    smtp_username: str = ""
    smtp_use_tls: bool = True
    subject_template: str = "[{vendor_name}] 재고자산확인서 송부 (기준일: {report_date_kr})"
    body_template: str = (
        "유상사급 타처보관 자료요청\n\n"
        "안녕하세요. {sender_name} 입니다.\n\n"
        "퍼시스 결산을 위해 아래 내용으로 {quarter} 타처 보관 재고 자료 요청 드립니다.\n\n"
        "---------------------------------------아 래-----------------------------------------------\n\n"
        "1. 첨부드리는 양식으로 재고조사 진행하여 기입 후, 직인 날인하여 회신 부탁드립니다.\n"
        "2. 첨부된 양식 자재는 {period_year}년 {period_start_month}월 부터 {period_end_month}월까지 출고된 자재 내역입니다.\n"
        "3. 원활한 조사를 위해 엑셀 및 PDF파일을 함께 첨부드리니, 두 파일 모두 회신 바랍니다.\n"
        "4. {reply_due_kr} 오전까지 회신 주시면 감사하겠습니다.\n\n"
        "감사합니다."
    )


@dataclass(slots=True)
class VendorEmail:
    vendor_name: str
    email: str


@dataclass(slots=True)
class PreparedEmail:
    vendor_name: str
    recipient: str
    subject: str
    body: str
    attachments: list[Path]
