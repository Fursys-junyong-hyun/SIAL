from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QDate, Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDateEdit,
    QDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .mail_models import MailSettings, PreparedEmail, VendorEmail
from .mail_service import (
    build_template_context,
    is_valid_email,
    render_body,
    render_subject,
    suggested_reply_due_date,
)
from .models import VendorPreview

STATUS_READY = "✅ 발송 가능"
STATUS_NO_EMAIL = "✉️ 이메일 미입력"
STATUS_INVALID_EMAIL = "⚠️ 이메일 형식 오류"
STATUS_NO_ATTACHMENT = "📎 첨부파일 없음"

COLOR_READY = QColor("#1f6f3a")
COLOR_WARN = QColor("#7a4a00")
COLOR_ERROR = QColor("#7a1a1a")


def evaluate_status(email: str, attachments: list[Path]) -> tuple[str, QColor]:
    if not email:
        return STATUS_NO_EMAIL, COLOR_WARN
    if not is_valid_email(email):
        return STATUS_INVALID_EMAIL, COLOR_ERROR
    if len(attachments) != 2 or not all(path.exists() for path in attachments):
        return STATUS_NO_ATTACHMENT, COLOR_WARN
    return STATUS_READY, COLOR_READY


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
        self.resize(1180, 860)

        self.settings = settings
        self.previews = previews
        self.vendor_email_map = dict(vendor_email_map)
        self.report_date = report_date
        self.attachment_map = attachment_map
        self.save_vendor_emails = save_vendor_emails

        outer = QVBoxLayout(self)
        outer.setContentsMargins(8, 8, 8, 8)
        outer.setSpacing(8)

        inner = QWidget()
        layout = QVBoxLayout(inner)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(12)
        layout.addWidget(self._build_settings_box())
        layout.addWidget(self._build_template_box(), stretch=2)
        layout.addWidget(self._build_target_box(), stretch=2)

        scroll = QScrollArea()
        scroll.setWidget(inner)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        outer.addWidget(scroll, stretch=1)
        outer.addLayout(self._build_buttons())

        self._wire_signals()
        self.refresh_targets()
        self._update_preview()

    def _build_settings_box(self) -> QGroupBox:
        box = QGroupBox("1단계 · 발송 설정")
        form = QFormLayout(box)

        self.method_combo = QComboBox()
        self.method_combo.addItem("Gmail (앱 비밀번호 / SMTP)", "gmail_smtp")
        self.method_combo.addItem("Gmail (Google 로그인 / OAuth)", "gmail_oauth")
        self.method_combo.addItem("Outlook (PC 클라이언트)", "outlook")
        self.method_combo.addItem("EML 초안 파일로 저장", "eml")
        self._select_method(self.settings.default_method)

        self.reply_due_edit = QDateEdit()
        self.reply_due_edit.setCalendarPopup(True)
        seed = suggested_reply_due_date()
        self.reply_due_edit.setDate(QDate(seed.year, seed.month, seed.day))

        self.preview_vendor_combo = QComboBox()
        for preview in self.previews:
            self.preview_vendor_combo.addItem(preview.vendor_name)

        form.addRow("발송 방식", self.method_combo)
        form.addRow("회신 요청일", self.reply_due_edit)
        form.addRow("미리보기 업체", self.preview_vendor_combo)
        return box

    def _build_template_box(self) -> QGroupBox:
        box = QGroupBox("2단계 · 메일 내용 (왼쪽 = 편집할 템플릿 · 오른쪽 = 선택 업체 기준 미리보기)")
        outer = QVBoxLayout(box)

        subject_row = QHBoxLayout()
        subject_row.addWidget(QLabel("제목 템플릿"))
        self.subject_edit = QLineEdit(self.settings.subject_template)
        subject_row.addWidget(self.subject_edit, stretch=1)
        outer.addLayout(subject_row)

        self.subject_preview_label = QLabel()
        self.subject_preview_label.setObjectName("helpLabel")
        self.subject_preview_label.setWordWrap(True)
        outer.addWidget(self.subject_preview_label)

        splitter = QSplitter(Qt.Horizontal)
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.addWidget(QLabel("본문 템플릿 (편집 가능)"))
        self.body_edit = QTextEdit()
        self.body_edit.setPlainText(self.settings.body_template)
        self.body_edit.setMinimumHeight(220)
        left_layout.addWidget(self.body_edit, stretch=1)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.addWidget(QLabel("본문 미리보기 (편집 불가)"))
        self.body_preview_edit = QTextEdit()
        self.body_preview_edit.setReadOnly(True)
        self.body_preview_edit.setMinimumHeight(220)
        right_layout.addWidget(self.body_preview_edit, stretch=1)

        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setSizes([520, 520])
        outer.addWidget(splitter, stretch=1)

        hint = QLabel(
            "치환 변수: {vendor_name}, {sender_name}, {quarter}, {period_year}, "
            "{period_start_month}, {period_end_month}, {report_date_kr}, {reply_due_kr}"
        )
        hint.setObjectName("helpLabel")
        hint.setWordWrap(True)
        outer.addWidget(hint)
        return box

    def _build_target_box(self) -> QGroupBox:
        box = QGroupBox("3단계 · 발송 대상 (이메일 칸을 클릭해 직접 입력/수정)")
        layout = QVBoxLayout(box)

        self.summary_label = QLabel()
        self.summary_label.setObjectName("helpLabel")
        layout.addWidget(self.summary_label)

        self.target_table = QTableWidget(0, 4)
        self.target_table.setHorizontalHeaderLabels(["업체명", "이메일", "첨부", "상태"])
        self.target_table.verticalHeader().setVisible(False)
        self.target_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.target_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.target_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.target_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.target_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.target_table.setMinimumHeight(240)
        layout.addWidget(self.target_table, stretch=1)
        return box

    def _build_buttons(self) -> QHBoxLayout:
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
        return buttons

    def _wire_signals(self) -> None:
        self.reply_due_edit.dateChanged.connect(self._update_preview)
        self.subject_edit.textChanged.connect(self._update_preview)
        self.body_edit.textChanged.connect(self._update_preview)
        self.preview_vendor_combo.currentTextChanged.connect(self._update_preview)
        self.target_table.itemChanged.connect(self._handle_item_changed)

    def _current_preview_vendor(self) -> str:
        return self.preview_vendor_combo.currentText() or (self.previews[0].vendor_name if self.previews else "")

    def _build_context(self, vendor_name: str) -> dict[str, str]:
        return build_template_context(
            vendor_name=vendor_name,
            report_date=self.report_date,
            reply_due_date=self.reply_due_edit.date().toPython(),
            sender_name=self.settings.sender_name or "(설정에서 사용자 이름 등록 필요)",
        )

    def _update_preview(self) -> None:
        vendor_name = self._current_preview_vendor()
        if not vendor_name:
            self.subject_preview_label.setText("(미리보기 대상 없음)")
            self.body_preview_edit.setPlainText("")
            return
        context = self._build_context(vendor_name)
        try:
            rendered_subject = render_subject(self.subject_edit.text(), context)
        except (KeyError, IndexError) as exc:
            rendered_subject = f"(템플릿 오류: {exc})"
        try:
            rendered_body = render_body(self.body_edit.toPlainText(), context)
        except (KeyError, IndexError) as exc:
            rendered_body = f"(템플릿 오류: {exc})"
        self.subject_preview_label.setText(f"제목 미리보기 → {rendered_subject}")
        self.body_preview_edit.setPlainText(rendered_body)

    def _sync_vendor_email_map_from_table(self) -> None:
        for row in range(self.target_table.rowCount()):
            vendor_item = self.target_table.item(row, 0)
            email_item = self.target_table.item(row, 1)
            if not vendor_item:
                continue
            self.vendor_email_map[vendor_item.text().strip()] = email_item.text().strip() if email_item else ""

    def _apply_status_to_row(self, row: int) -> None:
        vendor_item = self.target_table.item(row, 0)
        email_item = self.target_table.item(row, 1)
        status_item = self.target_table.item(row, 3)
        if not vendor_item or status_item is None:
            return
        email = (email_item.text().strip() if email_item else "")
        attachments = self.attachment_map.get(vendor_item.text().strip(), [])
        status_text, color = evaluate_status(email, attachments)
        status_item.setText(status_text)
        status_item.setForeground(color)

    def _handle_item_changed(self, item: QTableWidgetItem) -> None:
        if item.column() != 1:
            return
        self._sync_vendor_email_map_from_table()
        if self.save_vendor_emails is not None:
            self.save_vendor_emails(self.collect_vendor_emails())
        self._apply_status_to_row(item.row())
        self._update_summary()

    def _update_summary(self) -> None:
        ready = 0
        missing_email = 0
        invalid_email = 0
        missing_attachment = 0
        for row in range(self.target_table.rowCount()):
            status_item = self.target_table.item(row, 3)
            if status_item is None:
                continue
            text = status_item.text()
            if text == STATUS_READY:
                ready += 1
            elif text == STATUS_NO_EMAIL:
                missing_email += 1
            elif text == STATUS_INVALID_EMAIL:
                invalid_email += 1
            elif text == STATUS_NO_ATTACHMENT:
                missing_attachment += 1
        parts = [f"발송 가능 {ready}건"]
        if missing_email:
            parts.append(f"이메일 미입력 {missing_email}건")
        if invalid_email:
            parts.append(f"이메일 형식 오류 {invalid_email}건")
        if missing_attachment:
            parts.append(f"첨부파일 없음 {missing_attachment}건 (먼저 산출 실행 필요)")
        self.summary_label.setText(" · ".join(parts))

    def refresh_targets(self) -> None:
        self._sync_vendor_email_map_from_table()
        self.target_table.blockSignals(True)
        self.target_table.setRowCount(0)
        for preview in self.previews:
            row = self.target_table.rowCount()
            self.target_table.insertRow(row)
            email = self.vendor_email_map.get(preview.vendor_name, "")
            attachments = self.attachment_map.get(preview.vendor_name, [])
            status_text, color = evaluate_status(email, attachments)

            vendor_item = QTableWidgetItem(preview.vendor_name)
            vendor_item.setFlags(vendor_item.flags() & ~Qt.ItemIsEditable)
            email_item = QTableWidgetItem(email)
            attachment_item = QTableWidgetItem(f"{len(attachments)}/2")
            attachment_item.setFlags(attachment_item.flags() & ~Qt.ItemIsEditable)
            attachment_item.setTextAlignment(Qt.AlignCenter)
            status_item = QTableWidgetItem(status_text)
            status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
            status_item.setForeground(color)

            self.target_table.setItem(row, 0, vendor_item)
            self.target_table.setItem(row, 1, email_item)
            self.target_table.setItem(row, 2, attachment_item)
            self.target_table.setItem(row, 3, status_item)
        self.target_table.blockSignals(False)
        self._update_summary()

    def collect_vendor_emails(self) -> list[VendorEmail]:
        self._sync_vendor_email_map_from_table()
        return [VendorEmail(vendor_name=preview.vendor_name, email=self.vendor_email_map.get(preview.vendor_name, "")) for preview in self.previews]

    def build_prepared_emails(self) -> tuple[list[PreparedEmail], list[str]]:
        self._sync_vendor_email_map_from_table()
        subject_template = self.subject_edit.text().strip()
        body_template = self.body_edit.toPlainText().strip()

        prepared: list[PreparedEmail] = []
        excluded: list[str] = []
        for preview in self.previews:
            recipient = self.vendor_email_map.get(preview.vendor_name, "").strip()
            attachments = self.attachment_map.get(preview.vendor_name, [])
            if not recipient:
                excluded.append(f"{preview.vendor_name}: 이메일 미입력")
                continue
            if not is_valid_email(recipient):
                excluded.append(f"{preview.vendor_name}: 이메일 형식 오류 ({recipient})")
                continue
            if len(attachments) != 2 or not all(path.exists() for path in attachments):
                excluded.append(f"{preview.vendor_name}: 첨부파일 없음 (산출 실행 필요)")
                continue

            context = self._build_context(preview.vendor_name)
            try:
                subject = render_subject(subject_template, context)
                body = render_body(body_template, context)
            except (KeyError, IndexError) as exc:
                excluded.append(f"{preview.vendor_name}: 템플릿 변수 오류 ({exc})")
                continue
            prepared.append(
                PreparedEmail(
                    vendor_name=preview.vendor_name,
                    recipient=recipient,
                    subject=subject,
                    body=body,
                    attachments=attachments,
                )
            )
        return prepared, excluded

    def selected_method(self) -> str:
        return str(self.method_combo.currentData())

    def _select_method(self, method: str) -> None:
        # 레거시 "gmail" 값을 새 명칭으로 매핑
        alias = {"gmail": "gmail_smtp"}.get(method, method)
        for index in range(self.method_combo.count()):
            if self.method_combo.itemData(index) == alias:
                self.method_combo.setCurrentIndex(index)
                return
        self.method_combo.setCurrentIndex(0)
