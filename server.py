#!/usr/bin/env python3
import hmac
import io
import json
import os
import re
import secrets
import time
import zipfile
from http.cookies import SimpleCookie
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parent
PASS_FILE = ROOT / "pass.txt"
WORKBOOK = ROOT / "structure.xlsx"
SESSION_COOKIE = "phones_session"
SESSION_TTL = 8 * 60 * 60
MAX_LOGIN_BODY = 4096
MAX_ATTEMPTS = 5
ATTEMPT_WINDOW = 5 * 60
XML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def load_accounts(path):
    accounts = []
    for line_number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        fields = line.split()
        if not fields:
            continue
        if len(fields) != 2:
            raise RuntimeError(f"pass.txt:{line_number} must contain a username and password separated by whitespace")
        accounts.append((fields[0], fields[1]))
    if not accounts:
        raise RuntimeError("pass.txt contains no login accounts")
    return accounts


def make_private_workbook(source):
    result = io.BytesIO()
    with zipfile.ZipFile(source, "r") as original, zipfile.ZipFile(result, "w") as sanitized:
        shared_indices_to_redact = set()
        sheet_xml = ET.fromstring(original.read("xl/worksheets/sheet1.xml"))
        for row in sheet_xml.iter(f"{{{XML_NS}}}row"):
            for cell in row:
                if cell.tag == f"{{{XML_NS}}}c" and re.fullmatch(r"F\d+", cell.get("r", "")) and cell.get("t") == "s":
                    value = cell.find(f"{{{XML_NS}}}v")
                    if value is not None and value.text is not None:
                        shared_indices_to_redact.add(int(value.text))

        for item in original.infolist():
            contents = original.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                for row in sheet_xml.iter(f"{{{XML_NS}}}row"):
                    for cell in list(row):
                        if cell.tag == f"{{{XML_NS}}}c" and re.fullmatch(r"F\d+", cell.get("r", "")):
                            row.remove(cell)
                contents = ET.tostring(sheet_xml, encoding="utf-8", xml_declaration=True)
            elif item.filename == "xl/sharedStrings.xml" and shared_indices_to_redact:
                strings_xml = ET.fromstring(contents)
                shared_strings = list(strings_xml.iter(f"{{{XML_NS}}}si"))
                for index in shared_indices_to_redact:
                    if index < len(shared_strings):
                        item_element = shared_strings[index]
                        item_element.clear()
                        ET.SubElement(item_element, f"{{{XML_NS}}}t").text = ""
                contents = ET.tostring(strings_xml, encoding="utf-8", xml_declaration=True)
            sanitized.writestr(item, contents)
    return result.getvalue()


class Handler(SimpleHTTPRequestHandler):
    server_version = "PhonesDirectory"
    sys_version = ""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def end_headers(self):
        self.send_header("X-Content-Type-Options", "nosniff")
        super().end_headers()

    def send_json(self, status, payload, extra_headers=()):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for name, value in extra_headers:
            self.send_header(name, value)
        self.end_headers()
        self.wfile.write(body)

    def session_token(self):
        cookie = SimpleCookie()
        try:
            cookie.load(self.headers.get("Cookie", ""))
        except (ValueError, KeyError):
            return None
        morsel = cookie.get(SESSION_COOKIE)
        return morsel.value if morsel else None

    def authenticated(self):
        token = self.session_token()
        if not token:
            return False
        expires = self.server.sessions.get(token)
        if expires is None:
            return False
        if expires <= time.monotonic():
            self.server.sessions.pop(token, None)
            return False
        self.server.sessions[token] = time.monotonic() + SESSION_TTL
        return True

    def do_GET(self):
        path = urlsplit(self.path).path
        if path == "/api/session":
            self.send_json(200, {"authenticated": self.authenticated()})
            return
        if path == "/structure.xlsx":
            if self.authenticated():
                body = WORKBOOK.read_bytes()
            else:
                body = make_private_workbook(WORKBOOK)
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("Vary", "Cookie")
            self.end_headers()
            self.wfile.write(body)
            return
        if path == "/pass.txt":
            self.send_error(404)
            return
        if path.startswith("/api/"):
            self.send_json(404, {"error": "Not found"})
            return
        super().do_GET()

    def do_POST(self):
        path = urlsplit(self.path).path
        if path == "/api/logout":
            token = self.session_token()
            if token:
                self.server.sessions.pop(token, None)
            self.send_json(200, {"authenticated": False}, (
                ("Set-Cookie", f"{SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0"),
            ))
            return
        if path != "/api/login":
            self.send_json(404, {"error": "Not found"})
            return

        client_ip = self.client_address[0]
        now = time.monotonic()
        attempts = [attempt for attempt in self.server.login_attempts.get(client_ip, []) if now - attempt < ATTEMPT_WINDOW]
        self.server.login_attempts[client_ip] = attempts
        if len(attempts) >= MAX_ATTEMPTS:
            self.send_json(429, {"error": "Слишком много попыток. Попробуйте позже."})
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            self.send_json(400, {"error": "Некорректный запрос."})
            return
        if content_length <= 0 or content_length > MAX_LOGIN_BODY:
            self.send_json(400, {"error": "Некорректный запрос."})
            return
        try:
            payload = json.loads(self.rfile.read(content_length))
            username = payload.get("username", "")
            password = payload.get("password", "")
            if not isinstance(username, str) or not isinstance(password, str):
                raise ValueError("Invalid credentials format")
        except (json.JSONDecodeError, UnicodeDecodeError, AttributeError, ValueError):
            self.send_json(400, {"error": "Некорректный запрос."})
            return

        username_bytes = username.encode("utf-8")
        password_bytes = password.encode("utf-8")
        account = None
        for stored_username, stored_password in self.server.accounts:
            username_matches = hmac.compare_digest(username_bytes, stored_username.encode("utf-8"))
            password_matches = hmac.compare_digest(password_bytes, stored_password.encode("utf-8"))
            if username_matches & password_matches:
                account = stored_username
        if account is None:
            attempts.append(now)
            self.server.login_attempts[client_ip] = attempts
            self.send_json(401, {"error": "Неверный логин или пароль."})
            return

        token = secrets.token_urlsafe(32)
        self.server.sessions[token] = now + SESSION_TTL
        self.send_json(200, {"authenticated": True}, (
            ("Set-Cookie", f"{SESSION_COOKIE}={token}; Path=/; HttpOnly; SameSite=Lax; Max-Age={SESSION_TTL}"),
        ))

    def translate_path(self, path):
        parsed = urlsplit(path)
        if parsed.path == "/pass.txt":
            return str(ROOT / "__not_found__")
        return super().translate_path(path)


def main():
    accounts = load_accounts(PASS_FILE)
    if not WORKBOOK.is_file():
        raise RuntimeError("structure.xlsx is missing")
    server = ThreadingHTTPServer(("0.0.0.0", int(os.environ.get("APP_PORT", "8080"))), Handler)
    server.accounts = accounts
    server.sessions = {}
    server.login_attempts = {}
    print("Serving directory on 0.0.0.0:8080", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
