from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QDate, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDateEdit,
    QDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .mail_models import MailSettings, PreparedEmail, VendorEmail
from .mail_service import build_template_context, render_body, render_subject
from .models import VendorPreview


def default_reply_due_date() -> date:
    return date.today() + timedelta(days=7)


def build_preview_table(rows) -> QTableWidget:
    table = QTableWidget(len(rows), 5)
    table.setHorizontalHeaderLabels(["사업장", "업체명", "자재코드", "색상코드", "비고"])
    table.setEditTriggers(QAbstractItemView.NoEditTriggers)
    table.setSelectionBehavior(QAbstractItemView.SelectRows)
    table.verticalHeader().setVisible(False)
    table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
    table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
    table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
    table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
    table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)

    for row_index, row in enumerate(rows):
        values = [row.site_name, row.vendor_name, row.material_code, row.color_code, row.material_name]
        for col_index, value in enumerate(values):
            item = QTableWidgetItem(value)
            if col_index != 4:
                item.setTextAlignment(Qt.AlignCenter)
            table.setItem(row_index, col_index, item)
    return table


class PreviewDialog(QDialog):
    def __init__(self, previews: list[VendorPreview], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("전체 미리보기")
        self.resize(1080, 720)

        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        layout.addWidget(tabs)

        for preview in previews:
            tabs.addTab(build_preview_table(preview.rows), f"{preview.vendor_name} ({len(preview.rows)})")

        close_button = QPushButton("닫기")
        close_button.clicked.connect(self.accept)
        layout.addWidget(close_button, alignment=Qt.AlignRight)


class MailSendDialog(QDialog):
    def __init__(
        self,
        settings: MailSettings,
        previews: list[VendorPreview],
        vendor_email_map: dict[str, str],
        report_date: date,
        attachment_map: dict[str, list[Path]],
        save_vendor_emails: Callable[[list[VendorEmail]], None] | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("메일 발송")
        self.resize(1000, 760)

        self.settings = settings
        self.previews = previews
        self.vendor_email_map = dict(vendor_email_map)
        self.report_date = report_date
        self.attachment_map = attachment_map
        self.save_vendor_emails = save_vendor_emails
        self._body_seed_vendor = previews[0].vendor_name if previews else ""

        layout = QVBoxLayout(self)

        form_box = QGroupBox("발송 설정")
        form_layout = QFormLayout(form_box)

        self.method_combo = QComboBox()
        self.method_combo.addItem("Outlook", "outlook")
        self.method_combo.addItem("Gmail(SMTP)", "gmail")
        self.method_combo.setCurrentIndex(0 if settings.default_method == "outlook" else 1)

        self.reply_due_edit = QDateEdit()
        self.reply_due_edit.setCalendarPopup(True)
        seed = default_reply_due_date()
        self.reply_due_edit.setDate(QDate(seed.year, seed.month, seed.day))

        self.subject_edit = QLineEdit(settings.subject_template)
        self.body_edit = QTextEdit()
        self.body_edit.setPlainText(self._render_initial_body())

        form_layout.addRow("발송 방식", self.method_combo)
        form_layout.addRow("회신 요청일", self.reply_due_edit)
        form_layout.addRow("제목 템플릿", self.subject_edit)
        form_layout.addRow("본문", self.body_edit)
        layout.addWidget(form_box)

        guide = QLabel("본문은 실제 발송될 최종 양식입니다. 업체 이메일은 아래 표에서 바로 입력할 수 있습니다.")
        guide.setObjectName("helpLabel")
        layout.addWidget(guide)

        self.target_table = QTableWidget(0, 4)
        self.target_table.setHorizontalHeaderLabels(["업체명", "이메일", "첨부파일", "상태"])
        self.target_table.verticalHeader().setVisible(False)
        self.target_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.target_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.target_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.target_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.target_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        layout.addWidget(self.target_table, stretch=1)

        buttons = QHBoxLayout()
        refresh_button = QPushButton("대상 새로고침")
        send_button = QPushButton("발송")
        close_button = QPushButton("취소")
        refresh_button.clicked.connect(self.refresh_targets)
        send_button.clicked.connect(self.accept)
        close_button.clicked.connect(self.reject)
        buttons.addStretch(1)
        buttons.addWidget(refresh_button)
        buttons.addWidget(send_button)
        buttons.addWidget(close_button)
        layout.addLayout(buttons)

        self.reply_due_edit.dateChanged.connect(self._refresh_body_from_template)
        self.target_table.itemChanged.connect(self._handle_item_changed)
        self.refresh_targets()

    def _render_initial_body(self) -> str:
        context = build_template_context(
            vendor_name=self._body_seed_vendor,
            report_date=self.report_date,
            reply_due_date=self.reply_due_edit.date().toPython(),
            sender_name=self.settings.sender_name,
        )
        return render_body(self.settings.body_template, context)

    def _refresh_body_from_template(self) -> None:
        self.body_edit.setPlainText(self._render_initial_body())

    def _sync_vendor_email_map_from_table(self) -> None:
        for row in range(self.target_table.rowCount()):
            vendor_item = self.target_table.item(row, 0)
            email_item = self.target_table.item(row, 1)
            if not vendor_item:
                continue
            self.vendor_email_map[vendor_item.text().strip()] = email_item.text().strip() if email_item else ""

    def _handle_item_changed(self, item: QTableWidgetItem) -> None:
        if item.column() != 1:
            return
        self._sync_vendor_email_map_from_table()
        if self.save_vendor_emails is not None:
            self.save_vendor_emails(self.collect_vendor_emails())
        row = item.row()
        attachments = self.attachment_map.get(self.target_table.item(row, 0).text().strip(), [])
        status = "발송 가능"
        if not item.text().strip():
            status = "이메일 없음"
        elif len(attachments) != 2 or not all(path.exists() for path in attachments):
            status = "첨부파일 없음"
        status_item = self.target_table.item(row, 3)
        if status_item is not None:
            status_item.setText(status)

    def refresh_targets(self) -> None:
        self._sync_vendor_email_map_from_table()
        self.target_table.setRowCount(0)
        for preview in self.previews:
            row = self.target_table.rowCount()
            self.target_table.insertRow(row)
            email = self.vendor_email_map.get(preview.vendor_name, "")
            attachments = self.attachment_map.get(preview.vendor_name, [])
            status = "발송 가능"
            if not email:
                status = "이메일 없음"
            elif len(attachments) != 2 or not all(path.exists() for path in attachments):
                status = "첨부파일 없음"

            vendor_item = QTableWidgetItem(preview.vendor_name)
            vendor_item.setFlags(vendor_item.flags() & ~Qt.ItemIsEditable)
            attachment_item = QTableWidgetItem(str(len(attachments)))
            attachment_item.setFlags(attachment_item.flags() & ~Qt.ItemIsEditable)
            status_item = QTableWidgetItem(status)
            status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
            email_item = QTableWidgetItem(email)

            self.target_table.setItem(row, 0, vendor_item)
            self.target_table.setItem(row, 1, email_item)
            self.target_table.setItem(row, 2, attachment_item)
            self.target_table.setItem(row, 3, status_item)

    def collect_vendor_emails(self) -> list[VendorEmail]:
        self._sync_vendor_email_map_from_table()
        return [VendorEmail(vendor_name=preview.vendor_name, email=self.vendor_email_map.get(preview.vendor_name, "")) for preview in self.previews]

    def build_prepared_emails(self) -> tuple[list[PreparedEmail], list[str]]:
        self._sync_vendor_email_map_from_table()
        reply_due_date = self.reply_due_edit.date().toPython()
        subject_template = self.subject_edit.text().strip()
        final_body = self.body_edit.toPlainText().strip()

        prepared: list[PreparedEmail] = []
        excluded: list[str] = []
        for preview in self.previews:
            recipient = self.vendor_email_map.get(preview.vendor_name, "").strip()
            attachments = self.attachment_map.get(preview.vendor_name, [])
            if not recipient:
                excluded.append(f"{preview.vendor_name}: 이메일 미등록")
                continue
            if len(attachments) != 2 or not all(path.exists() for path in attachments):
                excluded.append(f"{preview.vendor_name}: 첨부파일 없음")
                continue

            context = build_template_context(
                vendor_name=preview.vendor_name,
                report_date=self.report_date,
                reply_due_date=reply_due_date,
                sender_name=self.settings.sender_name,
            )
            prepared.append(
                PreparedEmail(
                    vendor_name=preview.vendor_name,
                    recipient=recipient,
                    subject=render_subject(subject_template, context),
                    body=final_body,
                    attachments=attachments,
                )
            )
        return prepared, excluded

    def selected_method(self) -> str:
        return str(self.method_combo.currentData())
