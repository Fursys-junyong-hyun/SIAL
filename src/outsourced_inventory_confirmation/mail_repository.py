from __future__ import annotations

import sqlite3
from pathlib import Path

from .app_storage import sqlite_db_path
from .mail_models import MailSettings, VendorEmail


SETTINGS_KEYS = {
    "sender_name",
    "default_method",
    "smtp_host",
    "smtp_port",
    "smtp_username",
    "smtp_use_tls",
    "oauth_client_id",
    "oauth_account_email",
    "cc_list",
    "reply_to",
    "subject_template",
    "body_template",
}


class MailRepository:
    def __init__(self, db_path: Path | None = None) -> None:
        self.db_path = db_path or sqlite_db_path()
        self._init_db()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        return connection

    def _init_db(self) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS app_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS vendor_emails (
                    vendor_name TEXT PRIMARY KEY,
                    email TEXT NOT NULL DEFAULT ''
                )
                """
            )

    def load_settings(self) -> MailSettings:
        settings = MailSettings()
        with self._connect() as conn:
            rows = conn.execute("SELECT key, value FROM app_settings").fetchall()
        for row in rows:
            key = row["key"]
            value = row["value"]
            if key not in SETTINGS_KEYS:
                continue
            if key == "smtp_port":
                setattr(settings, key, int(value))
            elif key == "smtp_use_tls":
                setattr(settings, key, value == "1")
            elif key == "default_method" and value == "gmail":
                # 레거시 값: SMTP 앱 비밀번호 방식이지만 OAuth 우선 정책에 따라 가능한 경우 OAuth 로 이전
                setattr(settings, key, "gmail_smtp")
            else:
                setattr(settings, key, value)
        return settings

    def save_settings(self, settings: MailSettings) -> None:
        payload = {
            "sender_name": settings.sender_name,
            "default_method": settings.default_method,
            "smtp_host": settings.smtp_host,
            "smtp_port": str(settings.smtp_port),
            "smtp_username": settings.smtp_username,
            "smtp_use_tls": "1" if settings.smtp_use_tls else "0",
            "oauth_client_id": settings.oauth_client_id,
            "oauth_account_email": settings.oauth_account_email,
            "cc_list": settings.cc_list,
            "reply_to": settings.reply_to,
            "subject_template": settings.subject_template,
            "body_template": settings.body_template,
        }
        with self._connect() as conn:
            conn.executemany(
                "INSERT INTO app_settings(key, value) VALUES(?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                payload.items(),
            )

    def load_vendor_emails(self) -> list[VendorEmail]:
        with self._connect() as conn:
            rows = conn.execute("SELECT vendor_name, email FROM vendor_emails ORDER BY vendor_name").fetchall()
        return [VendorEmail(vendor_name=row["vendor_name"], email=row["email"]) for row in rows]

    def upsert_vendor_emails(self, vendor_emails: list[VendorEmail]) -> None:
        with self._connect() as conn:
            conn.executemany(
                "INSERT INTO vendor_emails(vendor_name, email) VALUES(?, ?) "
                "ON CONFLICT(vendor_name) DO UPDATE SET email=excluded.email",
                [(item.vendor_name, item.email) for item in vendor_emails],
            )

    def ensure_vendors(self, vendor_names: list[str]) -> None:
        with self._connect() as conn:
            conn.executemany(
                "INSERT INTO vendor_emails(vendor_name, email) VALUES(?, '') "
                "ON CONFLICT(vendor_name) DO NOTHING",
                [(vendor_name,) for vendor_name in vendor_names],
            )

    def get_vendor_email_map(self) -> dict[str, str]:
        return {item.vendor_name: item.email for item in self.load_vendor_emails()}
