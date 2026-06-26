from __future__ import annotations

import win32cred


_SMTP_PASSWORD_TARGET = "outsourced_inventory_confirmation_gmail_smtp"
_OAUTH_REFRESH_TARGET = "outsourced_inventory_confirmation_gmail_oauth_refresh"
_OAUTH_CLIENT_SECRET_TARGET = "outsourced_inventory_confirmation_gmail_oauth_client_secret"


def _save(target_name: str, user_name: str, value: str) -> None:
    credential = {
        "Type": win32cred.CRED_TYPE_GENERIC,
        "TargetName": target_name,
        "UserName": user_name,
        "CredentialBlob": value,
        "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
    }
    win32cred.CredWrite(credential, 0)


def _load(target_name: str) -> str:
    try:
        credential = win32cred.CredRead(TargetName=target_name, Type=win32cred.CRED_TYPE_GENERIC)
    except Exception:  # noqa: BLE001
        return ""
    blob = credential.get("CredentialBlob", b"")
    if isinstance(blob, bytes):
        return blob.decode("utf-16-le", errors="ignore")
    return str(blob)


def _delete(target_name: str) -> None:
    try:
        win32cred.CredDelete(TargetName=target_name, Type=win32cred.CRED_TYPE_GENERIC)
    except Exception:  # noqa: BLE001
        pass


def save_smtp_password(password: str) -> None:
    _save(_SMTP_PASSWORD_TARGET, "gmail_smtp", password)


def load_smtp_password() -> str:
    return _load(_SMTP_PASSWORD_TARGET)


def save_oauth_refresh_token(token: str) -> None:
    if not token:
        _delete(_OAUTH_REFRESH_TARGET)
        return
    _save(_OAUTH_REFRESH_TARGET, "gmail_oauth", token)


def load_oauth_refresh_token() -> str:
    return _load(_OAUTH_REFRESH_TARGET)


def clear_oauth_refresh_token() -> None:
    _delete(_OAUTH_REFRESH_TARGET)


def save_oauth_client_secret(secret: str) -> None:
    if not secret:
        _delete(_OAUTH_CLIENT_SECRET_TARGET)
        return
    _save(_OAUTH_CLIENT_SECRET_TARGET, "gmail_oauth_client_secret", secret)


def load_oauth_client_secret() -> str:
    return _load(_OAUTH_CLIENT_SECRET_TARGET)
