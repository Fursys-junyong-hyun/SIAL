from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSettings


ORG_NAME = "FURSYS"
APP_NAME = "유상사급 타처보관 확인서"
KEY_LAST_OUTPUT_DIR = "paths/last_output_dir"


def build_settings() -> QSettings:
    return QSettings(ORG_NAME, APP_NAME)


def load_last_output_dir() -> Path | None:
    settings = build_settings()
    value = settings.value(KEY_LAST_OUTPUT_DIR, "", type=str)
    return Path(value) if value else None


def save_last_output_dir(path: Path) -> None:
    settings = build_settings()
    settings.setValue(KEY_LAST_OUTPUT_DIR, str(path))
