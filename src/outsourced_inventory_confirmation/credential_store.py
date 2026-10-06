from __future__ import annotations

from pathlib import Path

import win32cred
import win32crypt

from .app_storage import app_data_dir


_SMTP_PASSWORD_TARGET = "outsourced_inventory_confirmation_gmail_smtp"
_OAUTH_REFRESH_TARGET = "outsourced_inventory_confirmation_gmail_oauth_refresh"
_OAUTH_CLIENT_SECRET_TARGET = "outsourced_inventory_confirmation_gmail_oauth_client_secret"

_SMTP_PASSWORD_FILE = "smtp_secret.bin"
_OAUTH_REFRESH_FILE = "oauth_refresh.bin"
_OAUTH_CLIENT_SECRET_FILE = "oauth_client_secret.bin"


class CredentialSaveError(RuntimeError):
    """Windows 자격 증명 저장소와 DPAPI 파일 저장이 모두 실패했을 때."""


# ---------- Windows Credential Manager ----------

def _cred_save(target_name: str, user_name: str, value: str) -> None:
    credential = {
        "Type": win32cred.CRED_TYPE_GENERIC,
        "TargetName": target_name,
        "UserName": user_name,
        "CredentialBlob": value,
        "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
    }
    win32cred.CredWrite(credential, 0)


def _cred_load(target_name: str) -> str:
    try:
        credential = win32cred.CredRead(TargetName=target_name, Type=win32cred.CRED_TYPE_GENERIC)
    except Exception:  # noqa: BLE001
        return ""
    blob = credential.get("CredentialBlob", b"")
    if isinstance(blob, bytes):
        return blob.decode("utf-16-le", errors="ignore")
    return str(blob)


def _cred_delete(target_name: str) -> None:
    try:
        win32cred.CredDelete(TargetName=target_name, Type=win32cred.CRED_TYPE_GENERIC)
    except Exception:  # noqa: BLE001
        pass


# ---------- DPAPI 파일 폴백 (사용자 단위 암호화) ----------

def _dpapi_path(filename: str) -> Path:
    return app_data_dir() / filename


def _dpapi_save(filename: str, value: str) -> None:
    encrypted = win32crypt.CryptProtectData(value.encode("utf-8"), filename, None, None, None, 0)
    _dpapi_path(filename).write_bytes(encrypted)


def _dpapi_load(filename: str) -> str:
    path = _dpapi_path(filename)
    if not path.exists():
        return ""
    try:
        _, decrypted = win32crypt.CryptUnprotectData(path.read_bytes(), None, None, None, 0)
        return decrypted.decode("utf-8")
    except Exception:  # noqa: BLE001
        return ""


def _dpapi_delete(filename: str) -> None:
    try:
        _dpapi_path(filename).unlink(missing_ok=True)
    except Exception:  # noqa: BLE001
        pass


# ---------- 공통 저장/로드: 저장소 우선, DPAPI 파일을 항상 미러 ----------

def _save_secret(target_name: str, user_name: str, filename: str, value: str) -> None:
    """자격 증명 저장소와 DPAPI 파일 양쪽에 저장한다.

    회사 보안 정책 등으로 Credential Manager 쓰기가 막힌 PC 에서도
    DPAPI 파일이 성공하면 정상 동작한다. 둘 다 실패하면 예외를 던져
    UI 에서 사용자에게 즉시 알릴 수 있게 한다.
    """
    if not value:
        _cred_delete(target_name)
        _dpapi_delete(filename)
        return

    cred_error: Exception | None = None
    dpapi_error: Exception | None = None
    try:
        _cred_save(target_name, user_name, value)
    except Exception as exc:  # noqa: BLE001
        cred_error = exc
    try:
        _dpapi_save(filename, value)
    except Exception as exc:  # noqa: BLE001
        dpapi_error = exc

    if cred_error is not None and dpapi_error is not None:
        raise CredentialSaveError(
            f"자격 증명 저장소 오류: {cred_error} / DPAPI 파일 오류: {dpapi_error}"
        )


def _load_secret(target_name: str, filename: str) -> str:
    value = _cred_load(target_name)
    if value:
        return value
    return _dpapi_load(filename)


# ---------- 공개 API ----------

def save_smtp_password(password: str) -> None:
    _save_secret(_SMTP_PASSWORD_TARGET, "gmail_smtp", _SMTP_PASSWORD_FILE, password)


def load_smtp_password() -> str:
    return _load_secret(_SMTP_PASSWORD_TARGET, _SMTP_PASSWORD_FILE)


def save_oauth_refresh_token(token: str) -> None:
    _save_secret(_OAUTH_REFRESH_TARGET, "gmail_oauth", _OAUTH_REFRESH_FILE, token)


def load_oauth_refresh_token() -> str:
    return _load_secret(_OAUTH_REFRESH_TARGET, _OAUTH_REFRESH_FILE)


def clear_oauth_refresh_token() -> None:
    _cred_delete(_OAUTH_REFRESH_TARGET)
    _dpapi_delete(_OAUTH_REFRESH_FILE)


def save_oauth_client_secret(secret: str) -> None:
    _save_secret(_OAUTH_CLIENT_SECRET_TARGET, "gmail_oauth_client_secret", _OAUTH_CLIENT_SECRET_FILE, secret)


def load_oauth_client_secret() -> str:
    return _load_secret(_OAUTH_CLIENT_SECRET_TARGET, _OAUTH_CLIENT_SECRET_FILE)
