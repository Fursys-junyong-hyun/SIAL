from __future__ import annotations

import win32cred


TARGET_NAME = "outsourced_inventory_confirmation_gmail_smtp"


def save_smtp_password(password: str) -> None:
    credential = {
        "Type": win32cred.CRED_TYPE_GENERIC,
        "TargetName": TARGET_NAME,
        "UserName": "gmail_smtp",
        "CredentialBlob": password,
        "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
    }
    win32cred.CredWrite(credential, 0)


def load_smtp_password() -> str:
    try:
        credential = win32cred.CredRead(TargetName=TARGET_NAME, Type=win32cred.CRED_TYPE_GENERIC)
    except Exception:  # noqa: BLE001
        return ""
    blob = credential.get("CredentialBlob", b"")
    if isinstance(blob, bytes):
        return blob.decode("utf-16-le", errors="ignore")
    return str(blob)
