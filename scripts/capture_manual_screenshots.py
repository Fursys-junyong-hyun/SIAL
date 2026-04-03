from __future__ import annotations

import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QMessageBox, QTableWidgetItem

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from outsourced_inventory_confirmation.main_window import MainWindow
from outsourced_inventory_confirmation.settings import build_settings


def populate_window(window: MainWindow) -> None:
    sample_dir = PROJECT_ROOT / "SAMPLE"
    source_paths = [
        next(sample_dir.glob("*FC1*.xls")),
        next(sample_dir.glob("*FC2*.xls")),
    ]
    for path in source_paths:
        row = window.source_table.rowCount()
        window.source_table.insertRow(row)
        window.source_table.setItem(row, 0, QTableWidgetItem(str(path)))
        window.source_table.setItem(row, 1, QTableWidgetItem(path.stem.split("_")[-1]))

    window.output_dir_edit.setText(str(Path.home() / "Documents" / "sample_output"))
    window.load_vendors()

    for index in range(min(6, window.vendor_list.count())):
        item = window.vendor_list.item(index)
        item.setCheckState(Qt.Checked)
    window.refresh_preview_tabs()

    window.main_tabs.setCurrentIndex(1)
    window.sender_name_edit.setText("홍길동")
    window.default_method_combo.setCurrentIndex(1)
    window.smtp_host_edit.setText("smtp.gmail.com")
    window.smtp_port_edit.setText("587")
    window.smtp_username_edit.setText("sample.user@gmail.com")
    window.smtp_password_edit.setText("app-password")
    window.smtp_tls_checkbox.setChecked(True)
    window.subject_template_edit.setText("[{vendor_name}] 재고자산확인서 송부 (기준일: {report_date_kr})")
    window.body_template_edit.setPlainText(window.mail_settings.body_template)

    for row in range(min(8, window.vendor_email_table.rowCount())):
        if not window.vendor_email_table.item(row, 1):
            window.vendor_email_table.setItem(row, 1, QTableWidgetItem(""))
        window.vendor_email_table.item(row, 1).setText(f"vendor{row + 1}@example.com")


def capture() -> None:
    app = QApplication.instance() or QApplication([])
    QMessageBox.information = staticmethod(lambda *args, **kwargs: QMessageBox.Ok)
    QMessageBox.warning = staticmethod(lambda *args, **kwargs: QMessageBox.Ok)
    QMessageBox.critical = staticmethod(lambda *args, **kwargs: QMessageBox.Ok)
    QMessageBox.question = staticmethod(lambda *args, **kwargs: QMessageBox.Yes)
    settings = build_settings()
    settings.setValue("paths/last_output_dir", str(Path.home() / "Documents" / "sample_output"))

    window = MainWindow()
    window.show()
    app.processEvents()
    populate_window(window)
    app.processEvents()

    output_dir = PROJECT_ROOT / "output" / "doc"
    output_dir.mkdir(parents=True, exist_ok=True)

    window.main_tabs.setCurrentIndex(0)
    app.processEvents()
    window.grab().save(str(output_dir / "manual_work_tab.png"))

    window.main_tabs.setCurrentIndex(1)
    app.processEvents()
    window.grab().save(str(output_dir / "manual_settings_tab.png"))

    window.close()
    app.processEvents()


if __name__ == "__main__":
    capture()
