from __future__ import annotations

import base64
import json
import secrets
import socket
import urllib.parse
import urllib.request
import webbrowser
from dataclasses import dataclass
from http.server import BaseHTTPRequestHandler, HTTPServer

GOOGLE_AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token"
GOOGLE_USERINFO_URL = "https://www.googleapis.com/oauth2/v2/userinfo"
SMTP_SEND_SCOPE = "https://mail.google.com/"
USERINFO_SCOPE = "https://www.googleapis.com/auth/userinfo.email"

SUCCESS_HTML = """<!doctype html><html lang='ko'><head><meta charset='utf-8'>
<title>로그인 완료</title></head><body style='font-family: sans-serif; padding: 40px; background:#0f1114; color:#f2f2f2;'>
<h1 style='color:#6ad48b;'>✅ Gmail 로그인이 완료되었습니다.</h1>
<p>이 창은 닫으셔도 됩니다. 프로그램으로 돌아가 발송을 진행해 주세요.</p>
</body></html>"""

ERROR_HTML_TEMPLATE = """<!doctype html><html lang='ko'><head><meta charset='utf-8'>
<title>로그인 오류</title></head><body style='font-family: sans-serif; padding: 40px; background:#0f1114; color:#f2f2f2;'>
<h1 style='color:#d96a6a;'>❌ 로그인에 실패했습니다.</h1>
<p>{message}</p>
<p>프로그램으로 돌아가 다시 시도해 주세요.</p>
</body></html>"""


@dataclass(slots=True)
class OAuthResult:
    refresh_token: str
    access_token: str
    account_email: str


class _CallbackHandler(BaseHTTPRequestHandler):
    received: dict[str, str] | None = None
    expected_state: str = ""

    def log_message(self, format: str, *args) -> None:  # noqa: A002
        return

    def do_GET(self) -> None:  # noqa: N802
        parsed = urllib.parse.urlparse(self.path)
        params = dict(urllib.parse.parse_qsl(parsed.query))
        _CallbackHandler.received = params

        if "error" in params:
            html = ERROR_HTML_TEMPLATE.format(message=params.get("error", "알 수 없는 오류"))
            status = 400
        elif params.get("state") != self.expected_state:
            html = ERROR_HTML_TEMPLATE.format(message="state 값이 일치하지 않습니다 (CSRF 보호)")
            status = 400
        elif "code" not in params:
            html = ERROR_HTML_TEMPLATE.format(message="인증 코드를 받지 못했습니다")
            status = 400
        else:
            html = SUCCESS_HTML
            status = 200

        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(html.encode("utf-8"))


def _find_free_port() -> int:
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def run_authorization_flow(client_id: str, client_secret: str) -> OAuthResult:
    """브라우저에서 Google 로그인을 진행해 refresh token 을 받아온다."""
    client_id = (client_id or "").strip()
    client_secret = (client_secret or "").strip()
    if not client_id or not client_secret:
        raise ValueError("OAuth 클라이언트 ID/Secret 이 필요합니다.")

    port = _find_free_port()
    redirect_uri = f"http://127.0.0.1:{port}"
    state = secrets.token_urlsafe(16)

    params = {
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": f"{SMTP_SEND_SCOPE} {USERINFO_SCOPE}",
        "access_type": "offline",
        "prompt": "consent",
        "state": state,
    }
    auth_url = f"{GOOGLE_AUTH_URL}?{urllib.parse.urlencode(params)}"

    _CallbackHandler.received = None
    _CallbackHandler.expected_state = state
    server = HTTPServer(("127.0.0.1", port), _CallbackHandler)
    server.timeout = 1

    if not webbrowser.open(auth_url):
        server.server_close()
        raise RuntimeError("기본 브라우저를 열 수 없습니다. 수동으로 다음 URL 을 열어주세요:\n" + auth_url)

    # 최대 3분 동안 콜백 대기
    deadline_loops = 180
    try:
        while _CallbackHandler.received is None and deadline_loops > 0:
            server.handle_request()
            deadline_loops -= 1
    finally:
        server.server_close()

    received = _CallbackHandler.received
    if received is None:
        raise RuntimeError("브라우저 로그인 응답을 받지 못했습니다 (시간 초과).")
    if "error" in received:
        raise RuntimeError(f"Google 인증 오류: {received['error']}")
    if received.get("state") != state:
        raise RuntimeError("OAuth state 불일치 — 보안상 중단합니다.")
    if "code" not in received:
        raise RuntimeError("인증 코드를 받지 못했습니다.")

    token_resp = _http_post(GOOGLE_TOKEN_URL, {
        "code": received["code"],
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
    })

    refresh_token = token_resp.get("refresh_token", "")
    access_token = token_resp.get("access_token", "")
    if not refresh_token or not access_token:
        raise RuntimeError(
            "Google 으로부터 refresh token 을 받지 못했습니다.\n"
            "OAuth 클라이언트 동의 화면 설정에서 access_type=offline 이 허용되는지 확인해 주세요.\n"
            f"응답: {json.dumps(token_resp, ensure_ascii=False)}"
        )

    email = _fetch_userinfo(access_token)
    return OAuthResult(refresh_token=refresh_token, access_token=access_token, account_email=email)


def refresh_access_token(client_id: str, client_secret: str, refresh_token: str) -> str:
    if not refresh_token:
        raise RuntimeError("저장된 refresh token 이 없습니다. 다시 로그인해 주세요.")
    if not client_id or not client_secret:
        raise RuntimeError("OAuth 클라이언트 ID/Secret 이 비어있습니다.")

    resp = _http_post(GOOGLE_TOKEN_URL, {
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
    })
    access_token = resp.get("access_token")
    if not access_token:
        raise RuntimeError(
            "Access token 발급에 실패했습니다. refresh token 이 만료되었거나 권한이 취소되었을 수 있습니다.\n"
            f"응답: {json.dumps(resp, ensure_ascii=False)}"
        )
    return access_token


def build_xoauth2_string(email: str, access_token: str) -> str:
    auth_string = f"user={email}\x01auth=Bearer {access_token}\x01\x01"
    return base64.b64encode(auth_string.encode()).decode()


def _http_post(url: str, data: dict) -> dict:
    encoded = urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(url, data=encoded, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as exc:  # noqa: F821 - exists at runtime
        body = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"HTTP {exc.code} : {body}") from exc


def _fetch_userinfo(access_token: str) -> str:
    req = urllib.request.Request(GOOGLE_USERINFO_URL)
    req.add_header("Authorization", f"Bearer {access_token}")
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read().decode())
    return data.get("email", "")
