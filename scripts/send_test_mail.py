"""저장된 Gmail 설정으로 자체 발송 테스트.

저장된 SMTP 자격증명(Windows Credential Manager + SQLite app.db)을 그대로 사용해서
본인 주소(또는 인자로 지정한 주소)로 짧은 테스트 메일을 1통 보낸다.

실행:
    python scripts/send_test_mail.py            # 본인 주소로 발송
    python scripts/send_test_mail.py x@y.com    # 지정 주소로 발송
"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

# src 패키지 경로 추가
PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from outsourced_inventory_confirmation.credential_store import load_smtp_password  # noqa: E402
from outsourced_inventory_confirmation.mail_models import PreparedEmail  # noqa: E402
from outsourced_inventory_confirmation.mail_repository import MailRepository  # noqa: E402
from outsourced_inventory_confirmation.mail_service import send_via_gmail  # noqa: E402


def main() -> int:
    repo = MailRepository()
    settings = repo.load_settings()
    password = load_smtp_password()

    if not settings.smtp_username:
        print("❌ Gmail 주소가 저장돼 있지 않습니다. 프로그램의 메일 설정 탭에서 입력해 주세요.")
        return 1
    if not password:
        print("❌ 앱 비밀번호가 저장돼 있지 않습니다. 프로그램의 메일 설정 탭에서 입력 후 [설정 저장] 을 눌러 주세요.")
        return 1

    recipient = sys.argv[1] if len(sys.argv) > 1 else settings.smtp_username

    manual_pdf = PROJECT_ROOT / "output" / "doc" / "유상사급타처보관_사용자매뉴얼.pdf"
    attachments = [manual_pdf] if manual_pdf.exists() else []

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    test_email = PreparedEmail(
        vendor_name="(테스트)",
        recipient=recipient,
        subject="[테스트] 유상사급 타처보관 메일 발송 점검",
        body=(
            "이 메일은 유상사급 타처보관 확인서 생성 프로그램의\n"
            "메일 발송 기능을 점검하기 위해 자동 발송된 테스트 메일입니다.\n\n"
            f"- 발송 시각: {timestamp}\n"
            f"- 발신자 이름: {settings.sender_name or '(미설정)'}\n"
            f"- 발신 계정 (From): {settings.smtp_username}\n"
            f"- 수신 (To): {recipient}\n"
            f"- 첨부파일: {len(attachments)}개"
            + (f" ({manual_pdf.name})" if attachments else "")
            + "\n\n"
            "이 메일이 정상적으로 도착했다면 실제 거래처 발송도 동일한 방식으로 처리됩니다.\n"
            "- 첨부파일이 들어 있다면 첨부 전송 경로도 정상입니다.\n\n"
            "(이 메일은 자동 발송되었습니다. 회신은 필요 없습니다.)"
        ),
        attachments=attachments,
    )

    print(f"발송 시도: {settings.smtp_username} → {recipient}")
    print(f"첨부: {[p.name for p in attachments] or '없음'}")
    try:
        send_via_gmail(settings, password, test_email)
    except Exception as exc:  # noqa: BLE001
        print(f"❌ 발송 실패: {type(exc).__name__}: {exc}")
        return 2

    print("✅ 발송 성공. 수신함을 확인해 주세요.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
