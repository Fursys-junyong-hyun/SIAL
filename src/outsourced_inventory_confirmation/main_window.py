from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from PySide6.QtCore import QDate, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QDateEdit,
    QDialog,
    QFileDialog,
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
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from .excel_parser import ParsedSource, build_vendor_preview, collect_vendor_names, parse_source_file
from .models import SourceFile, VendorPreview
from .rendering import render_excel, safe_vendor_name
from .settings import load_last_output_dir, save_last_output_dir


def guess_site_name(path: Path) -> str:
    stem = path.stem
    parts = [part.strip() for part in stem.split("_") if part.strip()]
    return parts[-1] if parts else stem


def default_report_date() -> date:
    today = date.today()
    first_of_month = today.replace(day=1)
    return first_of_month - timedelta(days=1)


class PreviewDialog(QDialog):
    def __init__(self, previews: list[VendorPreview], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("미리보기")
        self.resize(1080, 720)

        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        layout.addWidget(tabs)

        for preview in previews:
            table = build_preview_table(preview.rows)
            tabs.addTab(table, f"{preview.vendor_name} ({len(preview.rows)})")

        close_button = QPushButton("닫기")
        close_button.clicked.connect(self.accept)
        layout.addWidget(close_button, alignment=Qt.AlignRight)


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


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("유상사급 타처보관 확인서 생성기")
        self.resize(1440, 900)

        self.parsed_sources: list[ParsedSource] = []

        root = QWidget()
        self.setCentralWidget(root)
        page = QVBoxLayout(root)
        page.setContentsMargins(20, 20, 20, 20)
        page.setSpacing(14)

        page.addWidget(self._build_top_summary())
        page.addWidget(self._build_main_splitter(), stretch=1)
        page.addWidget(self._build_preview_group(), stretch=1)
        page.addLayout(self._build_actions())

        self._apply_theme()
        self._restore_last_output_dir()
        self._refresh_summary()

    def _build_top_summary(self) -> QWidget:
        panel = QWidget()
        layout = QHBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        self.summary_label = QLabel()
        self.summary_label.setObjectName("summaryLabel")
        layout.addWidget(self.summary_label)
        layout.addStretch(1)
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
        box = QGroupBox("소스 파일")
        layout = QVBoxLayout(box)

        help_label = QLabel("사업장별 자재유형별수불집계 파일을 추가한 뒤, 사업장명만 필요 시 수정합니다.")
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
        layout.addWidget(self.source_table, stretch=1)

        row = QHBoxLayout()
        add_button = QPushButton("파일 추가")
        remove_button = QPushButton("선택 제거")
        load_button = QPushButton("거래처 불러오기")
        add_button.clicked.connect(self.add_source_files)
        remove_button.clicked.connect(self.remove_selected_sources)
        load_button.clicked.connect(self.load_vendors)
        row.addWidget(add_button)
        row.addWidget(remove_button)
        row.addStretch(1)
        row.addWidget(load_button)
        layout.addLayout(row)
        return box

    def _build_vendor_group(self) -> QWidget:
        box = QGroupBox("거래처 선택")
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

        self.vendor_count_label = QLabel("0개 거래처")
        self.vendor_count_label.setObjectName("helpLabel")
        layout.addWidget(self.vendor_count_label)

        self.vendor_list = QListWidget()
        self.vendor_list.itemChanged.connect(lambda _item: self._refresh_summary())
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

        hint = QLabel("미리보기는 검토 전용입니다. 첫 탭은 즉시 확인용이며 전체 미리보기는 별도 창에서 열립니다.")
        hint.setObjectName("helpLabel")
        layout.addWidget(hint)

        self.preview_tabs = QTabWidget()
        layout.addWidget(self.preview_tabs, stretch=1)
        return box

    def _build_actions(self) -> QHBoxLayout:
        row = QHBoxLayout()
        refresh_button = QPushButton("선택 업체 반영")
        preview_button = QPushButton("전체 미리보기")
        generate_button = QPushButton("산출 실행")
        refresh_button.clicked.connect(self.refresh_preview_tabs)
        preview_button.clicked.connect(self.show_preview_dialog)
        generate_button.clicked.connect(self.generate_documents)
        row.addStretch(1)
        row.addWidget(refresh_button)
        row.addWidget(preview_button)
        row.addWidget(generate_button)
        return row

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
                margin-top: 12px;
                padding-top: 12px;
                background: #171a1f;
            }
            QGroupBox::title {
                left: 12px;
                padding: 0 6px;
                color: #d7e2f2;
            }
            QPushButton {
                background: #2d6cdf;
                border: none;
                border-radius: 8px;
                padding: 8px 14px;
                font-weight: 600;
                min-height: 18px;
            }
            QPushButton:hover {
                background: #3a79eb;
            }
            QLineEdit, QListWidget, QTableWidget, QDateEdit, QTabWidget::pane {
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
            QLabel#summaryLabel {
                background: #171a1f;
                border: 1px solid #2b2f36;
                border-radius: 10px;
                padding: 10px 12px;
                color: #dbe7f5;
                font-weight: 600;
            }
            """
        )

    def _restore_last_output_dir(self) -> None:
        last_dir = load_last_output_dir()
        if last_dir:
            self.output_dir_edit.setText(str(last_dir))

    def _refresh_summary(self) -> None:
        source_count = self.source_table.rowCount()
        vendor_count = self.vendor_list.count()
        selected_count = len(self._selected_vendor_names(silent=True))
        self.summary_label.setText(
            f"소스 파일 {source_count}건 / 거래처 {vendor_count}개 / 선택 업체 {selected_count}개"
        )

    def add_source_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "소스 파일 선택",
            "",
            "Excel Files (*.xls *.xlsx *.xlsm)",
        )
        for file_path in files:
            path = Path(file_path)
            row = self.source_table.rowCount()
            self.source_table.insertRow(row)
            self.source_table.setItem(row, 0, QTableWidgetItem(str(path)))
            self.source_table.setItem(row, 1, QTableWidgetItem(guess_site_name(path)))
        self._refresh_summary()

    def remove_selected_sources(self) -> None:
        rows = sorted({item.row() for item in self.source_table.selectedItems()}, reverse=True)
        for row in rows:
            self.source_table.removeRow(row)
        self._refresh_summary()

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
            raise ValueError("소스 파일을 1개 이상 선택해 주세요.")
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

        self.vendor_count_label.setText(f"{len(vendor_names)}개 거래처")
        self.select_all_checkbox.setChecked(False)
        self.preview_tabs.clear()
        self._refresh_summary()
        QMessageBox.information(self, "거래처 불러오기", "거래처 목록을 불러왔습니다.")

    def filter_vendors(self, text: str) -> None:
        keyword = text.strip().lower()
        visible_count = 0
        for index in range(self.vendor_list.count()):
            item = self.vendor_list.item(index)
            hidden = keyword not in item.text().lower()
            item.setHidden(hidden)
            if not hidden:
                visible_count += 1
        self.vendor_count_label.setText(f"표시 중 {visible_count}개 / 전체 {self.vendor_list.count()}개")

    def toggle_visible_vendors(self, state: int) -> None:
        check_state = Qt.Checked if state == Qt.Checked.value else Qt.Unchecked
        for index in range(self.vendor_list.count()):
            item = self.vendor_list.item(index)
            if not item.isHidden():
                item.setCheckState(check_state)
        self._refresh_summary()

    def _selected_vendor_names(self, silent: bool = False) -> list[str]:
        names: list[str] = []
        for index in range(self.vendor_list.count()):
            item = self.vendor_list.item(index)
            if item.checkState() == Qt.Checked:
                names.append(item.text())
        if not names and not silent:
            raise ValueError("거래처를 1개 이상 선택해 주세요.")
        return names

    def _selected_previews(self) -> list[VendorPreview]:
        if not self.parsed_sources:
            raise ValueError("먼저 거래처 불러오기를 실행해 주세요.")
        return [build_vendor_preview(self.parsed_sources, vendor_name) for vendor_name in self._selected_vendor_names()]

    def refresh_preview_tabs(self) -> None:
        try:
            previews = self._selected_previews()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "미리보기 오류", str(exc))
            return

        self.preview_tabs.clear()
        for preview in previews:
            table = build_preview_table(preview.rows)
            self.preview_tabs.addTab(table, f"{preview.vendor_name} ({len(preview.rows)})")
        self._refresh_summary()

    def show_preview_dialog(self) -> None:
        try:
            previews = self._selected_previews()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "미리보기 오류", str(exc))
            return

        if self.preview_tabs.count() == 0:
            self.refresh_preview_tabs()
        dialog = PreviewDialog(previews, self)
        dialog.exec()

    def select_output_dir(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "저장 폴더 선택", self.output_dir_edit.text().strip())
        if directory:
            self.output_dir_edit.setText(directory)
            save_last_output_dir(Path(directory))

    def generate_documents(self) -> None:
        try:
            previews = self._selected_previews()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "생성 오류", str(exc))
            return

        output_dir_text = self.output_dir_edit.text().strip()
        if not output_dir_text:
            QMessageBox.warning(self, "생성 오류", "저장 폴더를 선택해 주세요.")
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
                f"기존 파일 {len(colliding)}건이 있습니다. 덮어쓰시겠습니까?",
            )
            if answer != QMessageBox.Yes:
                return

        report_date = self.report_date_edit.date().toPython()
        generated: list[str] = []
        for preview in previews:
            document = render_excel(preview, report_date, output_root)
            generated.append(f"{document.vendor_name}: {document.output_dir}")

        QMessageBox.information(self, "생성 완료", "\n".join(generated))


def build_application() -> QApplication:
    app = QApplication.instance() or QApplication([])
    app.setApplicationName("유상사급 타처보관 확인서")
    app.setOrganizationName("FURSYS")
    return app
