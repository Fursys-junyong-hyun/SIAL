from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from PySide6.QtCore import QDate, Qt
from PySide6.QtGui import QColor, QDesktopServices
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
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

from .credential_store import (
    clear_oauth_refresh_token,
    load_oauth_client_secret,
    load_oauth_refresh_token,
    load_smtp_password,
    save_oauth_client_secret,
    save_oauth_refresh_token,
    save_smtp_password,
)
from .excel_parser import ParsedSource, build_vendor_preview, collect_vendor_names, parse_source_file
from .mail_models import MailSettings, VendorEmail
from .mail_oauth import refresh_access_token, run_authorization_flow
from .mail_repository import MailRepository
from .mail_service import (
    is_valid_email,
    save_as_eml,
    send_via_gmail,
    send_via_gmail_oauth,
    send_via_outlook,
    verify_gmail_credentials,
)
from .mail_ui import MailSendDialog, PreviewDialog, build_preview_table
from .models import RenderedDocument, SourceFile, VendorPreview
from .rendering import render_excel, safe_vendor_name
from .settings import load_last_output_dir, save_last_output_dir


GMAIL_APP_PASSWORD_URL = "https://myaccount.google.com/apppasswords"
GOOGLE_CLOUD_CONSOLE_URL = "https://console.cloud.google.com/apis/credentials"
EML_DRAFT_FOLDER_NAME = "_메일초안_EML"


def wrap_in_scroll_area(inner: QWidget) -> QScrollArea:
    scroll = QScrollArea()
    scroll.setWidget(inner)
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QFrame.NoFrame)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    return scroll


def make_collapsible(group_box: QGroupBox, expanded: bool = False) -> QGroupBox:
    """QGroupBox 의 체크박스 토글에 따라 내부 위젯을 실제로 숨기거나 보이게 한다."""
    group_box.setCheckable(True)
    title = group_box.title()
    arrow_expanded = "▼"
    arrow_collapsed = "▶"
    base_title = title.lstrip("▶▼ ").strip()

    def apply(checked: bool) -> None:
        for child in group_box.findChildren(QWidget):
            child.setVisible(checked)
        arrow = arrow_expanded if checked else arrow_collapsed
        group_box.setTitle(f"{arrow} {base_title}")

    group_box.setChecked(expanded)
    apply(expanded)
    group_box.toggled.connect(apply)
    return group_box


def guess_site_name(path: Path) -> str:
    stem = path.stem
    parts = [part.strip() for part in stem.split("_") if part.strip()]
    return parts[-1] if parts else stem


def default_report_date() -> date:
    today = date.today()
    first_of_month = today.replace(day=1)
    return first_of_month - timedelta(days=1)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("유상사급 타처보관 확인서 생성기")
        self.resize(1460, 960)

        self.parsed_sources: list[ParsedSource] = []
        self.generated_documents: dict[str, RenderedDocument] = {}
        self.mail_repository = MailRepository()
        self.mail_settings = self.mail_repository.load_settings()

        root = QWidget()
        self.setCentralWidget(root)
        page = QVBoxLayout(root)
        page.setContentsMargins(20, 20, 20, 20)
        page.setSpacing(14)

        self.main_tabs = QTabWidget()
        self.main_tabs.addTab(self._build_work_tab(), "업무")
        self.main_tabs.addTab(self._build_settings_tab(), "메일 설정")
        page.addWidget(self.main_tabs)

        self._apply_theme()
        self._restore_last_output_dir()
        self._load_mail_settings_into_form()
        self._load_vendor_emails_table()
        self._refresh_summary()
        self._refresh_button_states()
        self._prompt_for_setup_if_missing()

    # ---------- Work tab ----------

    def _build_work_tab(self) -> QWidget:
        container = QWidget()
        page = QVBoxLayout(container)
        page.setContentsMargins(4, 4, 4, 4)
        page.setSpacing(14)
        page.addWidget(self._build_top_summary())
        main_splitter = self._build_main_splitter()
        main_splitter.setMinimumHeight(320)
        page.addWidget(main_splitter, stretch=2)
        preview_group = self._build_preview_group()
        preview_group.setMinimumHeight(260)
        page.addWidget(preview_group, stretch=1)
        page.addLayout(self._build_actions())
        return wrap_in_scroll_area(container)

    def _build_top_summary(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        flow_label = QLabel(
            "진행 순서:  ① 소스 파일 추가  →  ② 거래처 불러오기  →  ③ 거래처 선택  →  "
            "④ 산출 실행  →  ⑤ 메일 발송"
        )
        flow_label.setObjectName("flowLabel")
        layout.addWidget(flow_label)

        self.summary_label = QLabel()
        self.summary_label.setObjectName("summaryLabel")
        layout.addWidget(self.summary_label)
        return panel

    def _build_main_splitter(self) -> QWidget:
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self._build_left_panel())
        splitter.addWidget(self._build_right_panel())
        splitter.setSizes([920, 420])
        return splitter

    def _build_left_panel(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)
        layout.addWidget(self._build_source_group())
        layout.addWidget(self._build_options_group())
        return container

    def _build_right_panel(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)
        layout.addWidget(self._build_vendor_group(), stretch=1)
        return container

    def _build_source_group(self) -> QWidget:
        box = QGroupBox("① 소스 파일")
        layout = QVBoxLayout(box)

        help_label = QLabel("사업장별 자재유형별수불집계 파일을 추가한 뒤, 필요 시 사업장명을 수정합니다.")
        help_label.setObjectName("helpLabel")
        layout.addWidget(help_label)

        self.source_table = QTableWidget(0, 2)
        self.source_table.setHorizontalHeaderLabels(["파일 경로", "사업장"])
        self.source_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.source_table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        self.source_table.verticalHeader().setVisible(False)
        self.source_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.source_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Fixed)
        self.source_table.setColumnWidth(1, 140)
        self.source_table.setAlternatingRowColors(True)
        self.source_table.setMinimumHeight(160)
        layout.addWidget(self.source_table, stretch=1)

        row = QHBoxLayout()
        self.add_source_button = QPushButton("파일 추가")
        self.remove_source_button = QPushButton("선택 제거")
        self.load_vendors_button = QPushButton("② 거래처 불러오기")
        self.load_vendors_button.setObjectName("primaryButton")
        self.add_source_button.clicked.connect(self.add_source_files)
        self.remove_source_button.clicked.connect(self.remove_selected_sources)
        self.load_vendors_button.clicked.connect(self.load_vendors)
        row.addWidget(self.add_source_button)
        row.addWidget(self.remove_source_button)
        row.addStretch(1)
        row.addWidget(self.load_vendors_button)
        layout.addLayout(row)
        return box

    def _build_vendor_group(self) -> QWidget:
        box = QGroupBox("③ 거래처 선택")
        layout = QVBoxLayout(box)

        search_row = QHBoxLayout()
        self.vendor_search = QLineEdit()
        self.vendor_search.setPlaceholderText("거래처명 검색")
        self.vendor_search.textChanged.connect(self.filter_vendors)
        self.select_all_checkbox = QCheckBox("현재 목록 전체 선택")
        self.select_all_checkbox.stateChanged.connect(self.toggle_visible_vendors)
        search_row.addWidget(self.vendor_search)
        search_row.addWidget(self.select_all_checkbox)
        layout.addLayout(search_row)

        self.vendor_count_label = QLabel("거래처를 먼저 불러와 주세요.")
        self.vendor_count_label.setObjectName("helpLabel")
        layout.addWidget(self.vendor_count_label)

        self.vendor_list = QListWidget()
        self.vendor_list.itemChanged.connect(lambda _item: (self._refresh_summary(), self._refresh_button_states()))
        self.vendor_list.setMinimumHeight(220)
        layout.addWidget(self.vendor_list, stretch=1)
        return box

    def _build_options_group(self) -> QWidget:
        box = QGroupBox("출력 옵션")
        layout = QGridLayout(box)
        layout.setColumnStretch(1, 1)

        self.report_date_edit = QDateEdit()
        self.report_date_edit.setCalendarPopup(True)
        seed_date = default_report_date()
        self.report_date_edit.setDate(QDate(seed_date.year, seed_date.month, seed_date.day))

        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setPlaceholderText("산출물 저장 폴더를 지정해 주세요.")
        self.output_dir_edit.textChanged.connect(self._refresh_button_states)
        browse_button = QPushButton("폴더 선택")
        browse_button.clicked.connect(self.select_output_dir)

        layout.addWidget(QLabel("기준일"), 0, 0)
        layout.addWidget(self.report_date_edit, 0, 1)
        layout.addWidget(QLabel("저장 폴더"), 1, 0)
        layout.addWidget(self.output_dir_edit, 1, 1)
        layout.addWidget(browse_button, 1, 2)
        return box

    def _build_preview_group(self) -> QWidget:
        box = QGroupBox("선택 업체 미리보기")
        layout = QVBoxLayout(box)

        hint = QLabel("'선택 업체 반영'을 누르면 아래 탭에 출력 예정 데이터가 표시됩니다. 검토 전용입니다.")
        hint.setObjectName("helpLabel")
        layout.addWidget(hint)

        self.preview_tabs = QTabWidget()
        self.preview_tabs.setMinimumHeight(220)
        layout.addWidget(self.preview_tabs, stretch=1)
        return box

    def _build_actions(self) -> QHBoxLayout:
        row = QHBoxLayout()
        self.refresh_preview_button = QPushButton("선택 업체 반영")
        self.preview_button = QPushButton("전체 미리보기")
        self.generate_button = QPushButton("④ 산출 실행")
        self.generate_button.setObjectName("primaryButton")
        self.mail_button = QPushButton("⑤ 메일 발송")
        self.mail_button.setObjectName("primaryButton")
        self.refresh_preview_button.clicked.connect(self.refresh_preview_tabs)
        self.preview_button.clicked.connect(self.show_preview_dialog)
        self.generate_button.clicked.connect(self.generate_documents)
        self.mail_button.clicked.connect(self.open_mail_dialog)
        row.addStretch(1)
        row.addWidget(self.refresh_preview_button)
        row.addWidget(self.preview_button)
        row.addWidget(self.generate_button)
        row.addWidget(self.mail_button)
        return row

    # ---------- Settings tab ----------

    def _build_settings_tab(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(14)

        layout.addWidget(self._build_sender_box())
        layout.addWidget(self._build_smtp_password_box())
        layout.addWidget(self._build_oauth_box())
        layout.addWidget(self._build_template_settings_box())
        layout.addWidget(self._build_vendor_email_box())
        layout.addLayout(self._build_settings_buttons())
        layout.addStretch(1)
        return wrap_in_scroll_area(container)

    def _build_sender_box(self) -> QWidget:
        box = QGroupBox("발신자 정보 / 기본 발송 방식")
        form = QFormLayout(box)

        self.sender_name_edit = QLineEdit()
        self.sender_name_edit.setPlaceholderText("예) 홍길동 (메일 본문에 표시됩니다)")

        self.default_method_combo = QComboBox()
        self.default_method_combo.addItem("Gmail (앱 비밀번호 / SMTP) — 권장", "gmail_smtp")
        self.default_method_combo.addItem("Gmail (Google 로그인 / OAuth)", "gmail_oauth")
        self.default_method_combo.addItem("Outlook (PC 클라이언트)", "outlook")
        self.default_method_combo.addItem("EML 초안 파일로 저장", "eml")

        form.addRow("사용자 이름", self.sender_name_edit)
        form.addRow("기본 발송 방식", self.default_method_combo)
        return box

    def _build_smtp_password_box(self) -> QWidget:
        box = QGroupBox("Gmail 앱 비밀번호 (권장)")
        self.smtp_password_group = box
        layout = QVBoxLayout(box)

        guide = QLabel(
            "Google 계정의 2단계 인증을 켠 뒤 발급받은 16자리 앱 비밀번호를 사용합니다.\n"
            "발급 페이지에서 비밀번호를 새로 만든 다음 아래에 그대로 붙여넣어 주세요 (공백은 자동 제거됩니다)."
        )
        guide.setObjectName("helpLabel")
        guide.setWordWrap(True)
        layout.addWidget(guide)

        app_password_button = QPushButton("Gmail 앱 비밀번호 발급 페이지 열기")
        app_password_button.setObjectName("linkButton")
        app_password_button.clicked.connect(lambda: self._open_url(GMAIL_APP_PASSWORD_URL))
        layout.addWidget(app_password_button, alignment=Qt.AlignLeft)

        form = QFormLayout()
        self.smtp_username_edit = QLineEdit()
        self.smtp_username_edit.setPlaceholderText("회사 Gmail 주소 (예: name@fursys.com)")
        self.smtp_password_edit = QLineEdit()
        self.smtp_password_edit.setEchoMode(QLineEdit.Password)
        self.smtp_password_edit.setPlaceholderText("Google 앱 비밀번호 16자리 (공백 무시)")
        form.addRow("Gmail 주소", self.smtp_username_edit)
        form.addRow("앱 비밀번호", self.smtp_password_edit)
        layout.addLayout(form)

        self.smtp_status_label = QLabel("연결 테스트 전")
        self.smtp_status_label.setObjectName("helpLabel")
        test_button = QPushButton("Gmail 연결 테스트")
        test_button.clicked.connect(self.test_gmail_connection)
        send_test_button = QPushButton("테스트 메일 보내기")
        send_test_button.setToolTip("저장된 Gmail 주소로 본인에게 테스트 메일을 1통 발송합니다.")
        send_test_button.clicked.connect(self.send_test_mail_to_self)
        test_row = QHBoxLayout()
        test_row.addStretch(1)
        test_row.addWidget(self.smtp_status_label)
        test_row.addWidget(test_button)
        test_row.addWidget(send_test_button)
        layout.addLayout(test_row)

        # SMTP 호스트/포트/TLS는 평소엔 감춰둔 고급 항목 — 클릭으로 펼치기
        advanced = QGroupBox("고급 옵션 (변경 불필요)")
        adv_form = QFormLayout(advanced)
        self.smtp_host_edit = QLineEdit()
        self.smtp_port_edit = QLineEdit()
        self.smtp_tls_checkbox = QCheckBox("TLS 사용")
        adv_form.addRow("SMTP 호스트", self.smtp_host_edit)
        adv_form.addRow("SMTP 포트", self.smtp_port_edit)
        adv_form.addRow("", self.smtp_tls_checkbox)
        make_collapsible(advanced, expanded=False)
        layout.addWidget(advanced)
        return box

    def _build_oauth_box(self) -> QWidget:
        box = QGroupBox("Gmail Google 로그인 (대체 옵션)")
        self.oauth_group = box
        layout = QVBoxLayout(box)

        guide = QLabel(
            "앱 비밀번호가 막힌 환경의 대체 수단입니다. Google Cloud Console 에 '데스크톱 앱' OAuth 클라이언트를 "
            "사전에 등록한 뒤(관리자 작업), Client ID / Secret 을 입력하고 'Google 로그인'을 누르세요.\n"
            "일반 사용 환경에서는 위의 앱 비밀번호 방식이면 충분합니다."
        )
        guide.setObjectName("helpLabel")
        guide.setWordWrap(True)
        layout.addWidget(guide)

        cloud_link_button = QPushButton("Google Cloud Console (사용자 인증 정보) 열기")
        cloud_link_button.setObjectName("linkButton")
        cloud_link_button.clicked.connect(lambda: self._open_url(GOOGLE_CLOUD_CONSOLE_URL))
        layout.addWidget(cloud_link_button, alignment=Qt.AlignLeft)

        form = QFormLayout()
        self.oauth_client_id_edit = QLineEdit()
        self.oauth_client_id_edit.setPlaceholderText("예) 123456789-xxxxx.apps.googleusercontent.com")
        self.oauth_client_secret_edit = QLineEdit()
        self.oauth_client_secret_edit.setEchoMode(QLineEdit.Password)
        self.oauth_client_secret_edit.setPlaceholderText("Google Cloud 에서 발급된 Client Secret")
        form.addRow("OAuth Client ID", self.oauth_client_id_edit)
        form.addRow("OAuth Client Secret", self.oauth_client_secret_edit)
        layout.addLayout(form)

        self.oauth_status_label = QLabel("로그인 필요")
        self.oauth_status_label.setObjectName("helpLabel")

        button_row = QHBoxLayout()
        self.oauth_login_button = QPushButton("Google 로그인")
        self.oauth_login_button.clicked.connect(self.start_google_login)
        self.oauth_logout_button = QPushButton("연결 해제")
        self.oauth_logout_button.clicked.connect(self.logout_google)
        button_row.addWidget(self.oauth_status_label, stretch=1)
        button_row.addWidget(self.oauth_login_button)
        button_row.addWidget(self.oauth_logout_button)
        layout.addLayout(button_row)
        make_collapsible(box, expanded=False)
        return box

    def _build_template_settings_box(self) -> QWidget:
        box = QGroupBox("메일 템플릿 기본값")
        layout = QFormLayout(box)
        self.subject_template_edit = QLineEdit()
        self.body_template_edit = QTextEdit()
        self.body_template_edit.setMinimumHeight(200)
        layout.addRow("제목 템플릿", self.subject_template_edit)
        layout.addRow("본문 템플릿", self.body_template_edit)
        hint = QLabel(
            "치환 변수: {vendor_name}, {sender_name}, {quarter}, {period_year}, "
            "{period_start_month}, {period_end_month}, {report_date_kr}, {reply_due_kr}"
        )
        hint.setObjectName("helpLabel")
        hint.setWordWrap(True)
        layout.addRow("", hint)
        return box

    def _build_vendor_email_box(self) -> QWidget:
        box = QGroupBox("업체 이메일 관리")
        layout = QVBoxLayout(box)
        vendor_help = QLabel(
            "현재 불러온 거래처를 가져오면 자동으로 행이 생성됩니다. 이메일을 입력해 두면 메일 발송 시 자동으로 채워집니다."
        )
        vendor_help.setObjectName("helpLabel")
        vendor_help.setWordWrap(True)
        layout.addWidget(vendor_help)

        self.vendor_email_table = QTableWidget(0, 3)
        self.vendor_email_table.setHorizontalHeaderLabels(["업체명", "이메일", "상태"])
        self.vendor_email_table.verticalHeader().setVisible(False)
        self.vendor_email_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.vendor_email_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.vendor_email_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.vendor_email_table.itemChanged.connect(self._handle_vendor_email_changed)
        self.vendor_email_table.setMinimumHeight(240)
        layout.addWidget(self.vendor_email_table, stretch=1)
        return box

    def _build_settings_buttons(self) -> QHBoxLayout:
        row = QHBoxLayout()
        import_vendors_button = QPushButton("현재 거래처 가져오기")
        save_settings_button = QPushButton("설정 저장")
        save_settings_button.setObjectName("primaryButton")
        import_vendors_button.clicked.connect(self.import_current_vendors_to_settings)
        save_settings_button.clicked.connect(self.save_mail_settings)
        row.addStretch(1)
        row.addWidget(import_vendors_button)
        row.addWidget(save_settings_button)
        return row

    # ---------- Theme ----------

    def _apply_theme(self) -> None:
        self.setStyleSheet(
            """
            QWidget {
                background: #111315;
                color: #f2f2f2;
                font-family: 'Malgun Gothic';
                font-size: 11pt;
            }
            QGroupBox {
                border: 1px solid #2b2f36;
                border-radius: 12px;
                margin-top: 14px;
                padding-top: 14px;
                background: #171a1f;
            }
            QGroupBox::title {
                left: 12px;
                padding: 0 6px;
                color: #d7e2f2;
                font-weight: 600;
            }
            QPushButton {
                background: #2d3340;
                border: none;
                border-radius: 8px;
                padding: 8px 14px;
                font-weight: 600;
                min-height: 18px;
                color: #f2f2f2;
            }
            QPushButton:hover {
                background: #3a4254;
            }
            QPushButton:disabled {
                background: #1a1d22;
                color: #5b6371;
            }
            QPushButton#primaryButton {
                background: #2d6cdf;
            }
            QPushButton#primaryButton:hover {
                background: #3a79eb;
            }
            QPushButton#primaryButton:disabled {
                background: #1a2740;
                color: #5b6371;
            }
            QPushButton#linkButton {
                background: transparent;
                color: #6aa3ff;
                padding: 4px 0;
                text-align: left;
                text-decoration: underline;
            }
            QPushButton#linkButton:hover {
                color: #8cbcff;
            }
            QLineEdit, QListWidget, QTableWidget, QDateEdit, QTabWidget::pane, QTextEdit, QComboBox {
                background: #0f1114;
                border: 1px solid #30363d;
                border-radius: 8px;
                padding: 6px;
            }
            QHeaderView::section {
                background: #20242b;
                color: #f2f2f2;
                border: none;
                padding: 6px;
            }
            QListWidget::item {
                padding: 6px 4px;
            }
            QTabBar::tab {
                background: #20242b;
                border: 1px solid #30363d;
                border-bottom: none;
                padding: 8px 12px;
                margin-right: 4px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
            }
            QTabBar::tab:selected {
                background: #2d6cdf;
            }
            QLabel#helpLabel {
                color: #aab3bf;
                font-size: 10pt;
            }
            QLabel#flowLabel {
                background: #0f1114;
                border: 1px solid #2b2f36;
                border-radius: 10px;
                padding: 8px 12px;
                color: #9eb3ce;
                font-size: 10pt;
            }
            QLabel#summaryLabel {
                background: #171a1f;
                border: 1px solid #2b2f36;
                border-radius: 10px;
                padding: 10px 12px;
                color: #dbe7f5;
                font-weight: 600;
            }
            QScrollArea {
                background: transparent;
                border: none;
            }
            QScrollArea > QWidget > QWidget {
                background: transparent;
            }
            QScrollBar:vertical {
                background: #0f1114;
                width: 12px;
                margin: 2px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background: #3a4254;
                min-height: 36px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical:hover {
                background: #4a5468;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: transparent;
            }
            QScrollBar:horizontal {
                background: #0f1114;
                height: 12px;
                margin: 2px;
                border-radius: 6px;
            }
            QScrollBar::handle:horizontal {
                background: #3a4254;
                min-width: 36px;
                border-radius: 5px;
            }
            QScrollBar::handle:horizontal:hover {
                background: #4a5468;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0px;
            }
            """
        )

    # ---------- State helpers ----------

    def _restore_last_output_dir(self) -> None:
        last_dir = load_last_output_dir()
        if last_dir:
            self.output_dir_edit.setText(str(last_dir))

    def _load_mail_settings_into_form(self) -> None:
        settings = self.mail_settings
        self.sender_name_edit.setText(settings.sender_name)
        self._select_default_method(settings.default_method)
        self.smtp_host_edit.setText(settings.smtp_host or "smtp.gmail.com")
        self.smtp_port_edit.setText(str(settings.smtp_port or 587))
        self.smtp_username_edit.setText(settings.smtp_username)
        self.smtp_password_edit.setText(load_smtp_password())
        self.smtp_tls_checkbox.setChecked(settings.smtp_use_tls)
        self.oauth_client_id_edit.setText(settings.oauth_client_id)
        self.oauth_client_secret_edit.setText(load_oauth_client_secret())
        self.subject_template_edit.setText(settings.subject_template)
        self.body_template_edit.setPlainText(settings.body_template)
        self._refresh_oauth_status()
        # OAuth 정보가 이미 저장돼 있으면 OAuth 박스를 펼쳐서 상태를 보여준다.
        if settings.oauth_client_id or settings.oauth_account_email or settings.default_method == "gmail_oauth":
            self.oauth_group.setChecked(True)

    def _select_default_method(self, method: str) -> None:
        alias = {"gmail": "gmail_smtp"}.get(method, method)
        for index in range(self.default_method_combo.count()):
            if self.default_method_combo.itemData(index) == alias:
                self.default_method_combo.setCurrentIndex(index)
                return
        self.default_method_combo.setCurrentIndex(0)

    def _refresh_oauth_status(self) -> None:
        token = load_oauth_refresh_token()
        if token and self.mail_settings.oauth_account_email:
            self.oauth_status_label.setText(f"✅ 로그인됨: {self.mail_settings.oauth_account_email}")
            self.oauth_status_label.setStyleSheet("color: #6ad48b;")
            self.oauth_logout_button.setEnabled(True)
        else:
            self.oauth_status_label.setText("로그인 필요")
            self.oauth_status_label.setStyleSheet("")
            self.oauth_logout_button.setEnabled(False)

    def _load_vendor_emails_table(self) -> None:
        self.vendor_email_table.blockSignals(True)
        self.vendor_email_table.setRowCount(0)
        for item in self.mail_repository.load_vendor_emails():
            row = self.vendor_email_table.rowCount()
            self.vendor_email_table.insertRow(row)
            vendor_item = QTableWidgetItem(item.vendor_name)
            vendor_item.setFlags(vendor_item.flags() & ~Qt.ItemIsEditable)
            self.vendor_email_table.setItem(row, 0, vendor_item)
            self.vendor_email_table.setItem(row, 1, QTableWidgetItem(item.email))
            self.vendor_email_table.setItem(row, 2, self._build_email_status_item(item.email))
        self.vendor_email_table.blockSignals(False)

    def _build_email_status_item(self, email: str) -> QTableWidgetItem:
        email = (email or "").strip()
        if not email:
            text, color = "미입력", QColor("#7a4a00")
        elif is_valid_email(email):
            text, color = "정상", QColor("#1f6f3a")
        else:
            text, color = "형식 오류", QColor("#7a1a1a")
        item = QTableWidgetItem(text)
        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        item.setForeground(color)
        item.setTextAlignment(Qt.AlignCenter)
        return item

    def _handle_vendor_email_changed(self, item: QTableWidgetItem) -> None:
        if item.column() != 1:
            return
        email = item.text().strip()
        self.vendor_email_table.blockSignals(True)
        self.vendor_email_table.setItem(item.row(), 2, self._build_email_status_item(email))
        self.vendor_email_table.blockSignals(False)

    def _collect_vendor_email_rows(self) -> list[VendorEmail]:
        vendor_emails: list[VendorEmail] = []
        for row in range(self.vendor_email_table.rowCount()):
            vendor_item = self.vendor_email_table.item(row, 0)
            email_item = self.vendor_email_table.item(row, 1)
            if not vendor_item:
                continue
            vendor_emails.append(
                VendorEmail(
                    vendor_name=vendor_item.text().strip(),
                    email=(email_item.text().strip() if email_item else ""),
                )
            )
        return vendor_emails

    def _persist_mail_settings(self, show_message: bool) -> bool:
        try:
            smtp_port = int(self.smtp_port_edit.text().strip() or "587")
        except ValueError:
            QMessageBox.warning(self, "설정 오류", "SMTP 포트는 숫자로 입력해 주세요.")
            return False

        settings = MailSettings(
            sender_name=self.sender_name_edit.text().strip(),
            default_method=str(self.default_method_combo.currentData()),
            smtp_host=self.smtp_host_edit.text().strip() or "smtp.gmail.com",
            smtp_port=smtp_port,
            smtp_username=self.smtp_username_edit.text().strip(),
            smtp_use_tls=self.smtp_tls_checkbox.isChecked(),
            oauth_client_id=self.oauth_client_id_edit.text().strip(),
            oauth_account_email=self.mail_settings.oauth_account_email,
            subject_template=self.subject_template_edit.text().strip(),
            body_template=self.body_template_edit.toPlainText().strip(),
        )
        self.mail_repository.save_settings(settings)
        self.mail_repository.upsert_vendor_emails(self._collect_vendor_email_rows())
        save_smtp_password(self._sanitized_password())
        save_oauth_client_secret(self.oauth_client_secret_edit.text().strip())
        self.mail_settings = settings
        if show_message:
            QMessageBox.information(self, "설정 저장", "메일 설정과 업체 이메일을 저장했습니다.")
        return True

    def _sanitized_password(self) -> str:
        # Gmail 앱 비밀번호는 보통 'xxxx xxxx xxxx xxxx' 공백 포함 16자리 형식이 안내됨.
        return self.smtp_password_edit.text().replace(" ", "")

    def save_mail_settings(self) -> None:
        self._persist_mail_settings(show_message=True)

    def import_current_vendors_to_settings(self) -> None:
        if not self.parsed_sources:
            QMessageBox.warning(
                self,
                "거래처 없음",
                "먼저 '업무' 탭에서 소스 파일을 추가하고 '거래처 불러오기'를 실행해 주세요.",
            )
            return
        self.mail_repository.ensure_vendors(collect_vendor_names(self.parsed_sources))
        self._load_vendor_emails_table()
        QMessageBox.information(self, "거래처 반영", "현재 거래처 목록을 이메일 관리 표에 반영했습니다.")

    def _open_url(self, url: str) -> None:
        from PySide6.QtCore import QUrl

        QDesktopServices.openUrl(QUrl(url))

    def start_google_login(self) -> None:
        if not self._persist_mail_settings(show_message=False):
            return
        client_id = self.mail_settings.oauth_client_id
        client_secret = self.oauth_client_secret_edit.text().strip()
        if not client_id or not client_secret:
            QMessageBox.warning(
                self,
                "OAuth 정보 필요",
                "OAuth Client ID / Secret 을 먼저 입력해 주세요.\n"
                "Google Cloud Console 에서 '데스크톱 앱' OAuth 클라이언트를 등록해 발급받습니다.",
            )
            return

        self.oauth_status_label.setText("브라우저에서 로그인 진행 중...")
        self.oauth_status_label.setStyleSheet("color: #d7e2f2;")
        self.oauth_login_button.setEnabled(False)
        QApplication.processEvents()

        try:
            result = run_authorization_flow(client_id, client_secret)
        except Exception as exc:  # noqa: BLE001
            self.oauth_login_button.setEnabled(True)
            self.oauth_status_label.setText(f"❌ 로그인 실패: {exc}")
            self.oauth_status_label.setStyleSheet("color: #d96a6a;")
            QMessageBox.warning(self, "Google 로그인 실패", str(exc))
            return

        save_oauth_refresh_token(result.refresh_token)
        self.mail_settings.oauth_account_email = result.account_email
        self.mail_repository.save_settings(self.mail_settings)
        self.oauth_login_button.setEnabled(True)
        self._refresh_oauth_status()
        QMessageBox.information(
            self,
            "Google 로그인 완료",
            f"{result.account_email} 계정으로 연결되었습니다.\n이후 메일 발송 시 자동으로 사용됩니다.",
        )

    def logout_google(self) -> None:
        answer = QMessageBox.question(
            self,
            "Google 연결 해제",
            "저장된 Google 로그인 정보를 삭제합니다. 다시 발송하려면 재로그인해야 합니다.\n계속할까요?",
        )
        if answer != QMessageBox.Yes:
            return
        clear_oauth_refresh_token()
        self.mail_settings.oauth_account_email = ""
        self.mail_repository.save_settings(self.mail_settings)
        self._refresh_oauth_status()

    def test_gmail_connection(self) -> None:
        if not self._persist_mail_settings(show_message=False):
            return
        username = self.mail_settings.smtp_username
        password = self._sanitized_password()
        if not username or not password:
            self.smtp_status_label.setText("⚠️ Gmail 주소와 앱 비밀번호를 모두 입력해 주세요.")
            self.smtp_status_label.setStyleSheet("color: #d49a3a;")
            return
        try:
            verify_gmail_credentials(self.mail_settings, password)
        except Exception as exc:  # noqa: BLE001
            self.smtp_status_label.setText(f"❌ 인증 실패: {exc}")
            self.smtp_status_label.setStyleSheet("color: #d96a6a;")
            QMessageBox.warning(
                self,
                "Gmail 연결 실패",
                "Gmail 로그인에 실패했습니다.\n\n"
                "1) 2단계 인증이 켜져 있는지 확인\n"
                "2) 일반 비밀번호가 아닌 '앱 비밀번호'를 사용했는지 확인\n"
                "3) 입력한 Gmail 주소가 올바른지 확인\n\n"
                f"세부 오류: {exc}",
            )
            return
        self.smtp_status_label.setText("✅ Gmail 로그인 성공")
        self.smtp_status_label.setStyleSheet("color: #6ad48b;")

    def send_test_mail_to_self(self) -> None:
        from datetime import datetime

        from .mail_models import PreparedEmail
        from .mail_service import send_via_gmail

        if not self._persist_mail_settings(show_message=False):
            return
        username = self.mail_settings.smtp_username
        password = self._sanitized_password()
        if not username or not password:
            QMessageBox.warning(
                self,
                "정보 부족",
                "Gmail 주소와 앱 비밀번호를 먼저 입력해 주세요.",
            )
            return

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        test_email = PreparedEmail(
            vendor_name="(테스트)",
            recipient=username,
            subject="[테스트] 유상사급 타처보관 메일 발송 점검",
            body=(
                "이 메일은 메일 발송 기능 점검용 자동 발송 메일입니다.\n\n"
                f"- 발송 시각: {timestamp}\n"
                f"- 발신자 이름: {self.mail_settings.sender_name or '(미설정)'}\n"
                f"- 발신 계정: {username}\n\n"
                "이 메일이 정상적으로 도착했다면 실제 거래처 발송도 동일한 방식으로 처리됩니다.\n"
                "(이 메일은 자동 발송되었습니다. 회신은 필요 없습니다.)"
            ),
            attachments=[],
        )

        self.smtp_status_label.setText("테스트 메일 발송 중...")
        self.smtp_status_label.setStyleSheet("color: #d7e2f2;")
        QApplication.processEvents()

        try:
            send_via_gmail(self.mail_settings, password, test_email)
        except Exception as exc:  # noqa: BLE001
            self.smtp_status_label.setText(f"❌ 발송 실패: {exc}")
            self.smtp_status_label.setStyleSheet("color: #d96a6a;")
            QMessageBox.warning(self, "테스트 메일 발송 실패", f"발송에 실패했습니다.\n\n{exc}")
            return

        self.smtp_status_label.setText(f"✅ 테스트 메일 발송 완료 → {username}")
        self.smtp_status_label.setStyleSheet("color: #6ad48b;")
        QMessageBox.information(
            self,
            "테스트 메일 발송 완료",
            f"{username} 주소로 테스트 메일을 보냈습니다.\n수신함을 확인해 주세요.",
        )

    def _prompt_for_setup_if_missing(self) -> None:
        has_smtp_password = bool(self.mail_settings.smtp_username and load_smtp_password())
        has_oauth = bool(self.mail_settings.oauth_account_email and load_oauth_refresh_token())
        if not has_smtp_password and not has_oauth:
            self.main_tabs.setCurrentIndex(1)

    def _refresh_summary(self) -> None:
        source_count = self.source_table.rowCount()
        vendor_count = self.vendor_list.count()
        selected_count = len(self._selected_vendor_names(silent=True))
        generated_count = len(self.generated_documents)
        self.summary_label.setText(
            f"소스 파일 {source_count}건 · 거래처 {vendor_count}개 · 선택 {selected_count}개 · 산출물 {generated_count}건"
        )

    def _refresh_button_states(self) -> None:
        has_source = self.source_table.rowCount() > 0
        has_vendors_loaded = self.vendor_list.count() > 0
        has_selection = len(self._selected_vendor_names(silent=True)) > 0
        has_output_dir = bool(self.output_dir_edit.text().strip())
        has_generated = len(self.generated_documents) > 0

        self.load_vendors_button.setEnabled(has_source)
        self.refresh_preview_button.setEnabled(has_selection)
        self.preview_button.setEnabled(has_selection)
        self.generate_button.setEnabled(has_selection and has_output_dir)
        self.mail_button.setEnabled(has_selection and has_generated)

        if not has_source:
            tooltip = "먼저 소스 파일을 추가해 주세요."
            self.load_vendors_button.setToolTip(tooltip)
        else:
            self.load_vendors_button.setToolTip("")

        if not has_selection:
            self.generate_button.setToolTip("거래처를 1개 이상 선택해야 산출할 수 있습니다.")
        elif not has_output_dir:
            self.generate_button.setToolTip("저장 폴더를 먼저 선택해 주세요.")
        else:
            self.generate_button.setToolTip("")

        if not has_selection:
            self.mail_button.setToolTip("거래처를 1개 이상 선택해 주세요.")
        elif not has_generated:
            self.mail_button.setToolTip("먼저 ④ 산출 실행으로 첨부할 Excel/PDF를 생성해 주세요.")
        else:
            self.mail_button.setToolTip("")

    # ---------- Source files ----------

    def add_source_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "소스 파일 선택", "", "Excel Files (*.xls *.xlsx *.xlsm)")
        for file_path in files:
            path = Path(file_path)
            row = self.source_table.rowCount()
            self.source_table.insertRow(row)
            self.source_table.setItem(row, 0, QTableWidgetItem(str(path)))
            self.source_table.setItem(row, 1, QTableWidgetItem(guess_site_name(path)))
        self._refresh_summary()
        self._refresh_button_states()

    def remove_selected_sources(self) -> None:
        rows = sorted({item.row() for item in self.source_table.selectedItems()}, reverse=True)
        for row in rows:
            self.source_table.removeRow(row)
        self._refresh_summary()
        self._refresh_button_states()

    def _collect_sources(self) -> list[SourceFile]:
        sources: list[SourceFile] = []
        for row in range(self.source_table.rowCount()):
            path_item = self.source_table.item(row, 0)
            site_item = self.source_table.item(row, 1)
            if not path_item:
                continue
            path = Path(path_item.text().strip())
            if not path.exists():
                raise FileNotFoundError(f"소스 파일이 존재하지 않습니다: {path}")
            site_name = site_item.text().strip() if site_item and site_item.text().strip() else guess_site_name(path)
            sources.append(SourceFile(path=path, site_name=site_name))
        if not sources:
            raise ValueError("소스 파일을 1개 이상 추가해 주세요.")
        return sources

    def load_vendors(self) -> None:
        try:
            self.parsed_sources = [parse_source_file(source) for source in self._collect_sources()]
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "파일 읽기 오류", str(exc))
            return

        self.vendor_list.clear()
        vendor_names = collect_vendor_names(self.parsed_sources)
        for vendor_name in vendor_names:
            item = QListWidgetItem(vendor_name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            self.vendor_list.addItem(item)

        self.mail_repository.ensure_vendors(vendor_names)
        self._load_vendor_emails_table()
        self.vendor_count_label.setText(f"불러온 거래처 {len(vendor_names)}개 · 발송할 업체를 체크하세요.")
        self.select_all_checkbox.setChecked(False)
        self.preview_tabs.clear()
        self.generated_documents.clear()
        self._refresh_summary()
        self._refresh_button_states()
        QMessageBox.information(self, "거래처 불러오기", f"거래처 {len(vendor_names)}개를 불러왔습니다.")

    def filter_vendors(self, text: str) -> None:
        keyword = text.strip().lower()
        visible_count = 0
        for index in range(self.vendor_list.count()):
            item = self.vendor_list.item(index)
            hidden = keyword not in item.text().lower()
            item.setHidden(hidden)
            if not hidden:
                visible_count += 1
        if self.vendor_list.count() == 0:
            self.vendor_count_label.setText("거래처를 먼저 불러와 주세요.")
        else:
            self.vendor_count_label.setText(f"표시 중 {visible_count}개 / 전체 {self.vendor_list.count()}개")

    def toggle_visible_vendors(self, state: int) -> None:
        check_state = Qt.Checked if state == Qt.Checked.value else Qt.Unchecked
        for index in range(self.vendor_list.count()):
            item = self.vendor_list.item(index)
            if not item.isHidden():
                item.setCheckState(check_state)
        self._refresh_summary()
        self._refresh_button_states()

    def _selected_vendor_names(self, silent: bool = False) -> list[str]:
        names: list[str] = []
        if not hasattr(self, "vendor_list"):
            return names
        for index in range(self.vendor_list.count()):
            item = self.vendor_list.item(index)
            if item.checkState() == Qt.Checked:
                names.append(item.text())
        if not names and not silent:
            raise ValueError("거래처를 1개 이상 선택해 주세요.")
        return names

    def _selected_previews(self) -> list[VendorPreview]:
        if not self.parsed_sources:
            raise ValueError("먼저 '② 거래처 불러오기'를 실행해 주세요.")
        return [build_vendor_preview(self.parsed_sources, vendor_name) for vendor_name in self._selected_vendor_names()]

    def refresh_preview_tabs(self) -> None:
        try:
            previews = self._selected_previews()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "미리보기 오류", str(exc))
            return

        self.preview_tabs.clear()
        for preview in previews:
            self.preview_tabs.addTab(build_preview_table(preview.rows), f"{preview.vendor_name} ({len(preview.rows)})")
        self._refresh_summary()

    def show_preview_dialog(self) -> None:
        try:
            previews = self._selected_previews()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "미리보기 오류", str(exc))
            return

        if self.preview_tabs.count() == 0:
            self.refresh_preview_tabs()
        PreviewDialog(previews, self).exec()

    def select_output_dir(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "저장 폴더 선택", self.output_dir_edit.text().strip())
        if directory:
            self.output_dir_edit.setText(directory)
            save_last_output_dir(Path(directory))
            self._refresh_button_states()

    def generate_documents(self) -> None:
        try:
            previews = self._selected_previews()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "생성 오류", str(exc))
            return

        output_dir_text = self.output_dir_edit.text().strip()
        if not output_dir_text:
            QMessageBox.warning(self, "생성 오류", "저장 폴더를 먼저 선택해 주세요.")
            return

        output_root = Path(output_dir_text)
        output_root.mkdir(parents=True, exist_ok=True)
        save_last_output_dir(output_root)

        file_stamp = date.today().strftime("%y%m%d")
        colliding: list[Path] = []
        for preview in previews:
            vendor_safe = safe_vendor_name(preview.vendor_name)
            vendor_dir = output_root / vendor_safe
            colliding.extend(
                [
                    vendor_dir / f"{vendor_safe}_재고자산확인서_{file_stamp}.xlsx",
                    vendor_dir / f"{vendor_safe}_재고자산확인서_{file_stamp}.pdf",
                ]
            )
        colliding = [path for path in colliding if path.exists()]

        if colliding:
            answer = QMessageBox.question(
                self,
                "덮어쓰기 확인",
                f"같은 이름의 파일 {len(colliding)}건이 이미 있습니다.\n덮어쓰고 진행하시겠습니까?",
            )
            if answer != QMessageBox.Yes:
                return

        report_date = self.report_date_edit.date().toPython()
        self.generated_documents.clear()
        generated: list[str] = []
        try:
            for preview in previews:
                document = render_excel(preview, report_date, output_root)
                self.generated_documents[preview.vendor_name] = document
                generated.append(f"• {document.vendor_name}")
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "생성 오류", str(exc))
            return

        self._refresh_summary()
        self._refresh_button_states()
        QMessageBox.information(
            self,
            "생성 완료",
            f"총 {len(generated)}개 업체의 Excel/PDF를 생성했습니다.\n저장 위치: {output_root}\n\n" + "\n".join(generated[:20]) + (
                "\n…" if len(generated) > 20 else ""
            ) + "\n\n바로 '⑤ 메일 발송'을 진행할 수 있습니다.",
        )

    def _build_attachment_map(self, previews: list[VendorPreview]) -> dict[str, list[Path]]:
        attachment_map: dict[str, list[Path]] = {}
        output_root = Path(self.output_dir_edit.text().strip()) if self.output_dir_edit.text().strip() else None
        stamp = date.today().strftime("%y%m%d")

        for preview in previews:
            if preview.vendor_name in self.generated_documents:
                document = self.generated_documents[preview.vendor_name]
                attachment_map[preview.vendor_name] = [document.xlsx_path, document.pdf_path]
                continue

            if not output_root:
                attachment_map[preview.vendor_name] = []
                continue

            safe_name = safe_vendor_name(preview.vendor_name)
            vendor_dir = output_root / safe_name
            attachment_map[preview.vendor_name] = [
                vendor_dir / f"{safe_name}_재고자산확인서_{stamp}.xlsx",
                vendor_dir / f"{safe_name}_재고자산확인서_{stamp}.pdf",
            ]
        return attachment_map

    def open_mail_dialog(self) -> None:
        try:
            previews = self._selected_previews()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "메일 오류", str(exc))
            return

        if not self._persist_mail_settings(show_message=False):
            return

        dialog = MailSendDialog(
            settings=self.mail_settings,
            previews=previews,
            vendor_email_map=self.mail_repository.get_vendor_email_map(),
            report_date=self.report_date_edit.date().toPython(),
            attachment_map=self._build_attachment_map(previews),
            save_vendor_emails=self.mail_repository.upsert_vendor_emails,
            parent=self,
        )
        if dialog.exec() != MailSendDialog.Accepted:
            return

        self.mail_repository.upsert_vendor_emails(dialog.collect_vendor_emails())
        self._load_vendor_emails_table()

        prepared_emails, excluded = dialog.build_prepared_emails()
        if not prepared_emails:
            QMessageBox.warning(
                self,
                "발송 불가",
                "발송 가능한 메일이 없습니다.\n\n" + ("\n".join(excluded) if excluded else ""),
            )
            return

        method = dialog.selected_method()
        if not self._method_prerequisites_ok(method):
            return

        method_label = {
            "gmail_oauth": "Gmail (Google 로그인)",
            "gmail_smtp": "Gmail (앱 비밀번호)",
            "outlook": "Outlook",
            "eml": "EML 초안 파일 저장",
        }.get(method, method)
        action_label = "초안으로 저장" if method == "eml" else "발송"

        summary_lines = [
            f"방식: {method_label}",
            f"대상: {len(prepared_emails)}건",
        ]
        if excluded:
            summary_lines.append(f"제외: {len(excluded)}건")
            summary_lines.append("")
            summary_lines.extend(excluded[:10])
            if len(excluded) > 10:
                summary_lines.append(f"…외 {len(excluded) - 10}건")
        summary_lines.append("")
        summary_lines.append(f"이대로 {action_label}하시겠습니까?")
        answer = QMessageBox.question(self, f"{action_label} 확인", "\n".join(summary_lines))
        if answer != QMessageBox.Yes:
            return

        successes, failures, eml_folder = self._dispatch_send(method, prepared_emails)

        result_lines = [f"✅ 성공 {len(successes)}건"]
        if failures:
            result_lines.append(f"❌ 실패 {len(failures)}건")
        if excluded:
            result_lines.append(f"⏭️ 제외 {len(excluded)}건")
        if eml_folder:
            result_lines.append("")
            result_lines.append(f"📁 초안 저장 위치: {eml_folder}")
            result_lines.append("Outlook 등에서 .eml 파일을 더블클릭하면 첨부 포함된 메일이 열립니다.")
        if failures:
            result_lines.append("")
            result_lines.append("[실패 내역]")
            result_lines.extend(failures)
        if excluded:
            result_lines.append("")
            result_lines.append("[제외 내역]")
            result_lines.extend(excluded)

        if method == "eml" and eml_folder and successes:
            box = QMessageBox(self)
            box.setWindowTitle("EML 초안 저장 완료")
            box.setIcon(QMessageBox.Information)
            box.setText("\n".join(result_lines))
            open_button = box.addButton("저장 폴더 열기", QMessageBox.AcceptRole)
            box.addButton("닫기", QMessageBox.RejectRole)
            box.exec()
            if box.clickedButton() is open_button:
                self._open_url(eml_folder.as_uri())
        else:
            QMessageBox.information(self, "발송 결과", "\n".join(result_lines))

    def _method_prerequisites_ok(self, method: str) -> bool:
        if method == "gmail_oauth":
            if not (self.mail_settings.oauth_client_id and load_oauth_client_secret() and load_oauth_refresh_token()):
                answer = QMessageBox.question(
                    self,
                    "Google 로그인 필요",
                    "Gmail (OAuth) 발송을 위해 Google 로그인이 필요합니다.\n"
                    "메일 설정 탭으로 이동해서 'Google 로그인'을 진행하시겠습니까?",
                )
                if answer == QMessageBox.Yes:
                    self.main_tabs.setCurrentIndex(1)
                return False
        elif method == "gmail_smtp":
            if not (self.mail_settings.smtp_username and load_smtp_password()):
                answer = QMessageBox.question(
                    self,
                    "Gmail 앱 비밀번호 필요",
                    "Gmail (SMTP) 발송을 위해 주소와 앱 비밀번호가 필요합니다.\n"
                    "메일 설정 탭으로 이동해서 등록하시겠습니까?",
                )
                if answer == QMessageBox.Yes:
                    self.main_tabs.setCurrentIndex(1)
                return False
        # outlook, eml 은 추가 자격증명이 필요 없음
        return True

    def _dispatch_send(self, method: str, prepared_emails) -> tuple[list[str], list[str], Path | None]:
        successes: list[str] = []
        failures: list[str] = []
        eml_folder: Path | None = None

        if method == "gmail_oauth":
            client_id = self.mail_settings.oauth_client_id
            client_secret = load_oauth_client_secret()
            refresh_token = load_oauth_refresh_token()
            try:
                access_token = refresh_access_token(client_id, client_secret, refresh_token)
            except Exception as exc:  # noqa: BLE001
                QMessageBox.warning(
                    self,
                    "OAuth 토큰 만료",
                    "Google access token 발급에 실패했습니다.\n"
                    "메일 설정 탭에서 다시 'Google 로그인'을 진행해 주세요.\n\n"
                    f"세부 오류: {exc}",
                )
                return successes, failures, eml_folder

            for item in prepared_emails:
                try:
                    access_token = send_via_gmail_oauth(
                        self.mail_settings,
                        refresh_token,
                        client_id,
                        client_secret,
                        item,
                        access_token=access_token,
                    )
                    successes.append(item.vendor_name)
                except Exception as exc:  # noqa: BLE001
                    failures.append(f"{item.vendor_name}: {exc}")
            return successes, failures, eml_folder

        if method == "gmail_smtp":
            password = load_smtp_password()
            for item in prepared_emails:
                try:
                    send_via_gmail(self.mail_settings, password, item)
                    successes.append(item.vendor_name)
                except Exception as exc:  # noqa: BLE001
                    failures.append(f"{item.vendor_name}: {exc}")
            return successes, failures, eml_folder

        if method == "outlook":
            for item in prepared_emails:
                try:
                    send_via_outlook(item)
                    successes.append(item.vendor_name)
                except Exception as exc:  # noqa: BLE001
                    failures.append(f"{item.vendor_name}: {exc}")
            return successes, failures, eml_folder

        if method == "eml":
            output_root_text = self.output_dir_edit.text().strip()
            output_root = Path(output_root_text) if output_root_text else Path.cwd()
            eml_folder = output_root / EML_DRAFT_FOLDER_NAME
            stamp = date.today().strftime("%y%m%d")
            for item in prepared_emails:
                try:
                    safe_name = safe_vendor_name(item.vendor_name)
                    output_path = eml_folder / f"{safe_name}_재고자산확인서_{stamp}.eml"
                    save_as_eml(self.mail_settings, item, output_path)
                    successes.append(item.vendor_name)
                except Exception as exc:  # noqa: BLE001
                    failures.append(f"{item.vendor_name}: {exc}")
            return successes, failures, eml_folder

        failures.append(f"알 수 없는 발송 방식: {method}")
        return successes, failures, eml_folder


def build_application() -> QApplication:
    app = QApplication.instance() or QApplication([])
    app.setApplicationName("유상사급 타처보관 확인서")
    app.setOrganizationName("FURSYS")
    return app
