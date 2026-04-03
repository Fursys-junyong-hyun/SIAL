from __future__ import annotations

from pathlib import Path

from .settings import APP_NAME, ORG_NAME


def app_data_dir() -> Path:
    base = Path.home() / "AppData" / "Local" / ORG_NAME / APP_NAME
    base.mkdir(parents=True, exist_ok=True)
    return base


def sqlite_db_path() -> Path:
    return app_data_dir() / "app.db"
