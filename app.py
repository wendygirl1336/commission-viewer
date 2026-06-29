from __future__ import annotations

import json
import mimetypes
import hmac
import os
import re
import secrets
import socket
import time
from email import policy
from email.parser import BytesParser
from http import cookies
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent
DATA_FILE = ROOT / "current_upload.json"
ADMIN_ID = "admin"
ADMIN_PASSWORD = "1234"
VIEWER_ID = "company"
VIEWER_PASSWORD = "1234"
ADMIN_LOGIN_IDS = {ADMIN_ID, "관리자"}
VIEWER_LOGIN_IDS = {VIEWER_ID, "회사"}
SESSION_SECONDS = 60 * 60 * 4
MAX_LOGIN_FAILURES = 5
LOCK_SECONDS = 60 * 5
SESSIONS: dict[str, dict[str, Any]] = {}
LOGIN_FAILURES: dict[str, dict[str, Any]] = {}


def empty_upload(message: str = "", error: str = "") -> dict[str, Any]:
    return {"rows": [], "fileName": "", "message": message, "error": error}


def load_saved_upload() -> dict[str, Any]:
    if not DATA_FILE.exists():
        return empty_upload()
    try:
        data = json.loads(DATA_FILE.read_text(encoding="utf-8"))
        if isinstance(data, dict) and isinstance(data.get("rows"), list):
            return {
                "rows": data.get("rows", []),
                "fileName": data.get("fileName", ""),
                "message": data.get("message", ""),
                "error": data.get("error", ""),
            }
    except Exception:
        pass
    return empty_upload()


def save_upload(data: dict[str, Any]) -> None:
    DATA_FILE.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


CURRENT_UPLOAD: dict[str, Any] = load_saved_upload()


def clean(value: Any) -> str:
    return re.sub(r"\s+", " ", "" if value is None else str(value)).strip()


def compact(value: Any) -> str:
    return re.sub(r"\s+", "", clean(value))


def num(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return 0.0


def find_col(row: list[Any], candidates: list[str]) -> int:
    labels = [compact(cell) for cell in row]
    for idx, label in enumerate(labels):
        if any(candidate in label for candidate in candidates):
            return idx
    return -1


def find_year_col(rows: list[list[Any]], header_idx: int, candidates: list[str]) -> int:
    for offset in range(3):
        if header_idx + offset >= len(rows):
            continue
        for idx, cell in enumerate(rows[header_idx + offset]):
            label = compact(cell)
            if any(candidate in label for candidate in candidates):
                return idx
    return -1


def parse_life_sheet(rows: list[list[Any]], sheet_name: str) -> list[dict[str, Any]]:
    header_idx = -1
    company_col = -1
    product_col = -1

    for idx, row in enumerate(rows[:25]):
        c_col = find_col(row, ["제휴사", "보험사", "회사"])
        p_col = find_col(row, ["상품구분", "상품명", "상품"])
        if c_col >= 0 and p_col >= 0:
            header_idx = idx
            company_col = c_col
            product_col = p_col
            break

    if header_idx < 0:
        return []

    y1_col = find_year_col(rows, header_idx, ["1차년", "1차년도"])
    y2_col = find_year_col(rows, header_idx, ["2차년", "2차년도"])
    y3_col = find_year_col(rows, header_idx, ["3차년", "3차년도"])
    if min(y1_col, y2_col, y3_col) < 0:
        return []

    parsed: list[dict[str, Any]] = []
    current_company = ""
    for row in rows[header_idx + 2 :]:
        company_cell = clean(row[company_col] if company_col < len(row) else "")
        product = clean(row[product_col] if product_col < len(row) else "")
        if company_cell:
            current_company = company_cell
        company = company_cell or current_company
        if not company or not product:
            continue
        if any(label in compact(product) for label in ["상품구분", "상품명"]):
            continue
        year1 = num(row[y1_col] if y1_col < len(row) else 0)
        year2 = num(row[y2_col] if y2_col < len(row) else 0)
        year3 = num(row[y3_col] if y3_col < len(row) else 0)
        parsed.append(
            {
                "company": company,
                "product": product,
                "year1": year1,
                "year2": year2,
                "year3": year3,
                "total": year1 + year2 + year3,
                "source": sheet_name,
            }
        )
    return parsed


def parse_nonlife_sheet(rows: list[list[Any]], sheet_name: str) -> list[dict[str, Any]]:
    header_idx = -1
    for idx, row in enumerate(rows):
        labels = [compact(cell) for cell in row]
        if "보험사" in labels and "상품명" in labels and any("총환산" in label for label in labels):
            header_idx = idx
            break
    if header_idx < 0:
        return []

    header = rows[header_idx]
    company_col = find_col(header, ["보험사"])
    product_col = find_col(header, ["상품명"])
    y1_col = find_col(header, ["총환산"])
    y2_col = find_col(header, ["환산2차년도", "2차년도", "2차년"])
    y3_col = find_col(header, ["환산3차년도", "3차년도", "3차년"])
    if min(company_col, product_col, y1_col, y2_col, y3_col) < 0:
        return []

    grouped: dict[tuple[str, str], dict[str, Any]] = {}
    for row in rows[header_idx + 1 :]:
        company = clean(row[company_col] if company_col < len(row) else "")
        product = clean(row[product_col] if product_col < len(row) else "")
        if not company or not product:
            continue
        key = (company, product)
        item = grouped.setdefault(
            key,
            {"company": company, "product": product, "year1": 0.0, "year2": 0.0, "year3": 0.0, "source": f"{sheet_name} 합산"},
        )
        item["year1"] += num(row[y1_col] if y1_col < len(row) else 0)
        item["year2"] += num(row[y2_col] if y2_col < len(row) else 0)
        item["year3"] += num(row[y3_col] if y3_col < len(row) else 0)

    result = []
    for item in grouped.values():
        item["total"] = item["year1"] + item["year2"] + item["year3"]
        result.append(item)
    return result


def parse_workbook(file_bytes: bytes) -> list[dict[str, Any]]:
    if not file_bytes:
        raise ValueError("업로드된 파일이 비어 있습니다.")
    if not file_bytes.startswith(b"PK"):
        raise ValueError(".xlsx, .xlsm 등 최신 엑셀 파일만 업로드해 주세요. 구형 .xls 파일은 지원하지 않습니다.")
    workbook = load_workbook(BytesIO(file_bytes), data_only=True, read_only=True)
    rows: list[dict[str, Any]] = []
    for ws in workbook.worksheets:
        values = [list(row) for row in ws.iter_rows(values_only=True)]
        rows.extend(parse_life_sheet(values, ws.title))
        rows.extend(parse_nonlife_sheet(values, ws.title))
    return [row for row in rows if row["company"] and row["product"]]


def parse_upload(body: bytes, content_type: str) -> tuple[bytes, dict[str, str]]:
    message = BytesParser(policy=policy.default).parsebytes(
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
    )
    if not message.is_multipart():
        raise ValueError("multipart 형식이 아닙니다.")
    fields: dict[str, str] = {}
    file_bytes = b""
    for part in message.iter_parts():
        if part.get_filename():
            file_bytes = part.get_payload(decode=True) or b""
        else:
            name = part.get_param("name", header="content-disposition")
            if name:
                fields[str(name)] = clean(part.get_content())
    if not file_bytes:
        raise ValueError("업로드 파일을 찾지 못했습니다.")
    return file_bytes, fields


def session_from_cookie(cookie_header: str | None) -> str:
    if not cookie_header:
        return ""
    jar = cookies.SimpleCookie()
    jar.load(cookie_header)
    return jar.get("admin_session").value if "admin_session" in jar else ""


def make_session(role: str) -> tuple[str, dict[str, Any]]:
    token = secrets.token_urlsafe(32)
    session = {"expires": time.time() + SESSION_SECONDS, "csrf": secrets.token_urlsafe(32), "role": role}
    SESSIONS[token] = session
    return token, session


def valid_session(token: str) -> dict[str, Any] | None:
    if not token:
        return None
    session = SESSIONS.get(token)
    if not session:
        return None
    if session.get("expires", 0) < time.time():
        SESSIONS.pop(token, None)
        return None
    return session


def local_ip() -> str:
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.connect(("8.8.8.8", 80))
        ip = sock.getsockname()[0]
        sock.close()
        return ip
    except OSError:
        return "127.0.0.1"


class Handler(BaseHTTPRequestHandler):
    def end_headers(self) -> None:
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def send_json(self, status: int, payload: dict[str, Any], extra_headers: dict[str, str] | None = None) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        if extra_headers:
            for key, value in extra_headers.items():
                self.send_header(key, value)
        self.end_headers()
        self.wfile.write(data)

    def redirect_home(self) -> None:
        self.send_response(303)
        self.send_header("Location", "/")
        self.end_headers()

    def read_json(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def is_admin(self) -> bool:
        session = self.session()
        return bool(session and session.get("role") == "admin")

    def is_authenticated(self) -> bool:
        return bool(self.session())

    def session(self) -> dict[str, Any] | None:
        return valid_session(session_from_cookie(self.headers.get("Cookie")))

    def csrf_ok(self, submitted: str) -> bool:
        session = self.session()
        return bool(session and submitted and hmac.compare_digest(submitted, str(session.get("csrf", ""))))

    def client_ip(self) -> str:
        return self.client_address[0] if self.client_address else "unknown"

    def login_locked(self) -> bool:
        item = LOGIN_FAILURES.get(self.client_ip())
        return bool(item and item.get("count", 0) >= MAX_LOGIN_FAILURES and item.get("until", 0) > time.time())

    def record_login_failure(self) -> None:
        item = LOGIN_FAILURES.setdefault(self.client_ip(), {"count": 0, "until": 0})
        item["count"] = int(item.get("count", 0)) + 1
        if item["count"] >= MAX_LOGIN_FAILURES:
            item["until"] = time.time() + LOCK_SECONDS

    def clear_login_failure(self) -> None:
        LOGIN_FAILURES.pop(self.client_ip(), None)

    def do_GET(self) -> None:
        if self.path == "/api/session":
            session = self.session()
            self.send_json(
                200,
                {
                    "authenticated": bool(session),
                    "admin": bool(session and session.get("role") == "admin"),
                    "role": session.get("role", "") if session else "",
                    "csrf": session.get("csrf", "") if session else "",
                },
            )
            return

        if self.path.startswith("/assets/"):
            target = (ROOT / self.path.lstrip("/")).resolve()
            assets_root = (ROOT / "assets").resolve()
            if not str(target).startswith(str(assets_root)) or not target.exists() or not target.is_file():
                self.send_json(404, {"error": "Not found"})
                return
            data = target.read_bytes()
            content_type = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        html = (ROOT / "index.html").read_text(encoding="utf-8")
        visible_upload = CURRENT_UPLOAD if self.is_authenticated() else empty_upload(message="로그인 후 조회할 수 있습니다.")
        payload = json.dumps(visible_upload, ensure_ascii=False).replace("</", "<\\/")
        initial_data = f"<script>window.__INITIAL_UPLOAD__ = {payload};</script>\n"
        app_script_marker = "<script>\n    const state"
        if app_script_marker in html:
            html = html.replace(app_script_marker, initial_data + app_script_marker, 1)
        else:
            html = html.replace("</body>", initial_data + "</body>")
        data = html.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self) -> None:
        global CURRENT_UPLOAD

        if self.path == "/upload":
            session = self.session()
            if not session or session.get("role") != "admin":
                self.redirect_home()
                return
            try:
                length = int(self.headers.get("Content-Length", "0"))
                content_type = self.headers.get("Content-Type", "")
                file_bytes, fields = parse_upload(self.rfile.read(length), content_type)
                if not hmac.compare_digest(fields.get("csrf", ""), str(session.get("csrf", ""))):
                    self.redirect_home()
                    return
                rows = parse_workbook(file_bytes)
                CURRENT_UPLOAD = {
                    "rows": rows,
                    "fileName": "업로드된 엑셀 파일",
                    "message": "" if rows else "인식 가능한 수수료 표를 찾지 못했습니다.",
                    "error": "",
                }
                save_upload(CURRENT_UPLOAD)
            except Exception as exc:
                CURRENT_UPLOAD = {"rows": [], "fileName": "", "message": "", "error": f"엑셀 파일을 읽지 못했습니다: {exc}"}
            self.redirect_home()
            return

        if self.path == "/api/login":
            try:
                if self.login_locked():
                    self.send_json(429, {"error": "로그인 시도가 너무 많습니다. 잠시 후 다시 시도해 주세요."})
                    return
                payload = self.read_json()
                login_id = str(payload.get("id", "")).strip().lower()
                password = str(payload.get("password", ""))
                role = ""
                if login_id in ADMIN_LOGIN_IDS and hmac.compare_digest(password, ADMIN_PASSWORD):
                    role = "admin"
                elif login_id in VIEWER_LOGIN_IDS and hmac.compare_digest(password, VIEWER_PASSWORD):
                    role = "viewer"

                if role:
                    self.clear_login_failure()
                    token, session = make_session(role)
                    self.send_json(
                        200,
                        {"authenticated": True, "admin": role == "admin", "role": role, "csrf": session["csrf"]},
                        {"Set-Cookie": f"admin_session={token}; Path=/; Max-Age={SESSION_SECONDS}; HttpOnly; SameSite=Lax"},
                    )
                    return
                self.record_login_failure()
                self.send_json(401, {"error": "아이디 또는 비밀번호가 올바르지 않습니다."})
            except Exception as exc:
                self.send_json(400, {"error": str(exc)})
            return

        if self.path == "/api/logout":
            token = session_from_cookie(self.headers.get("Cookie"))
            SESSIONS.pop(token, None)
            self.send_json(200, {"admin": False}, {"Set-Cookie": "admin_session=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax"})
            return

        if self.path == "/api/switch-viewer":
            session = self.session()
            if not session or session.get("role") != "admin":
                self.send_json(403, {"error": "관리자 로그인 후 조회 모드로 전환할 수 있습니다."})
                return
            if not hmac.compare_digest(self.headers.get("X-CSRF-Token", ""), str(session.get("csrf", ""))):
                self.send_json(403, {"error": "보안 토큰이 올바르지 않습니다. 새로고침 후 다시 시도해 주세요."})
                return
            session["role"] = "viewer"
            session["csrf"] = secrets.token_urlsafe(32)
            self.send_json(200, {"authenticated": True, "admin": False, "role": "viewer", "csrf": session["csrf"]})
            return

        if self.path == "/api/upload":
            session = self.session()
            if not session or session.get("role") != "admin":
                self.send_json(403, {"error": "관리자 로그인 후 업로드해 주세요."})
                return
            if not hmac.compare_digest(self.headers.get("X-CSRF-Token", ""), str(session.get("csrf", ""))):
                self.send_json(403, {"error": "보안 토큰이 올바르지 않습니다. 새로고침 후 다시 시도해 주세요."})
                return
            try:
                length = int(self.headers.get("Content-Length", "0"))
                content_type = self.headers.get("Content-Type", "")
                file_bytes, _fields = parse_upload(self.rfile.read(length), content_type)
                rows = parse_workbook(file_bytes)
                message = "" if rows else "인식 가능한 수수료 표를 찾지 못했습니다."
                CURRENT_UPLOAD = {
                    "rows": rows,
                    "fileName": "업로드된 엑셀 파일",
                    "message": message,
                    "error": "",
                }
                save_upload(CURRENT_UPLOAD)
                self.send_json(200, {"rows": rows, "message": message})
            except Exception as exc:
                self.send_json(400, {"error": f"엑셀 파일을 읽지 못했습니다: {exc}"})
            return

        if self.path == "/api/clear-upload":
            session = self.session()
            if not session or session.get("role") != "admin":
                self.send_json(403, {"error": "관리자 로그인 후 삭제할 수 있습니다."})
                return
            if not hmac.compare_digest(self.headers.get("X-CSRF-Token", ""), str(session.get("csrf", ""))):
                self.send_json(403, {"error": "보안 토큰이 올바르지 않습니다. 새로고침 후 다시 시도해 주세요."})
                return
            CURRENT_UPLOAD = empty_upload("업로드된 엑셀 데이터가 삭제되었습니다.")
            save_upload(CURRENT_UPLOAD)
            self.send_json(200, CURRENT_UPLOAD)
            return

        self.send_json(404, {"error": "Not found"})

    def log_message(self, format: str, *args: Any) -> None:
        return


if __name__ == "__main__":
    host = "0.0.0.0"
    port = int(os.environ.get("PORT", "8766"))
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"Local:   http://127.0.0.1:{port}/")
    print(f"Network: http://{local_ip()}:{port}/")
    server.serve_forever()
