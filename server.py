#!/usr/bin/env python3
import hmac
import io
import json
import os
import posixpath
import re
import secrets
import tempfile
import threading
import time
import zipfile
from datetime import datetime, timezone
from http.cookies import SimpleCookie
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parent
PASS_FILE = ROOT / "pass.txt"
WORKBOOK = Path(os.environ.get("WORKBOOK_PATH", ROOT / "structure.xlsx"))
CHANGES_LOG = Path(os.environ.get("CHANGES_LOG_PATH", ROOT / "data" / "changes.log"))
SESSION_COOKIE = "phones_session"
SESSION_TTL = 8 * 60 * 60
MAX_LOGIN_BODY = 4096
MAX_EDIT_BODY = 32768
MAX_EDIT_VALUE = 10000
MAX_ATTEMPTS = 5
ATTEMPT_WINDOW = 5 * 60
XML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
ET.register_namespace("", XML_NS)


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


def normalized_request_path(path):
    decoded = unquote(urlsplit(path).path).replace("\\", "/")
    return "/" + posixpath.normpath(decoded).lstrip("/")


def first_worksheet_path(workbook):
    workbook_xml = ET.fromstring(workbook.read("xl/workbook.xml"))
    sheet = workbook_xml.find(f".//{{{XML_NS}}}sheet")
    if sheet is None:
        raise ValueError("В книге Excel нет листов.")
    relation_id = sheet.get(f"{{{REL_NS}}}id")
    relations_xml = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
    relation = next(
        (item for item in relations_xml.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
         if item.get("Id") == relation_id),
        None,
    )
    if relation is None:
        raise ValueError("Не удалось определить первый лист книги Excel.")
    target = relation.get("Target", "")
    sheet_path = posixpath.normpath(
        target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target)
    )
    if sheet_path not in workbook.namelist():
        raise ValueError("Первый лист книги Excel отсутствует.")
    return sheet_path


def read_shared_strings(workbook):
    if "xl/sharedStrings.xml" not in workbook.namelist():
        return []
    strings_xml = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
    return [
        "".join(text.text or "" for text in item.iter(f"{{{XML_NS}}}t"))
        for item in strings_xml.iter(f"{{{XML_NS}}}si")
    ]


def column_index(cell_reference):
    letters = re.match(r"[A-Z]+", cell_reference)
    if not letters:
        raise ValueError("Некорректный адрес ячейки в книге Excel.")
    result = 0
    for letter in letters.group(0):
        result = result * 26 + ord(letter) - ord("A") + 1
    return result - 1


def column_name(index):
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(ord("A") + remainder) + name
    return name


def cell_text(cell, shared_strings):
    cell_type = cell.get("t")
    if cell_type == "s":
        value = cell.find(f"{{{XML_NS}}}v")
        if value is None or value.text is None:
            return ""
        try:
            return shared_strings[int(value.text)]
        except (ValueError, IndexError):
            raise ValueError("Некорректная ссылка на строку в книге Excel.") from None
    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.iter(f"{{{XML_NS}}}t"))
    value = cell.find(f"{{{XML_NS}}}v")
    return value.text if value is not None and value.text is not None else ""


def update_workbook(source, row_number, changes):
    result = io.BytesIO()
    with zipfile.ZipFile(source, "r") as original:
        worksheet_path = first_worksheet_path(original)
        worksheet = ET.fromstring(original.read(worksheet_path))
        shared_strings = read_shared_strings(original)
        sheet_data = worksheet.find(f".//{{{XML_NS}}}sheetData")
        rows = sheet_data.findall(f"{{{XML_NS}}}row") if sheet_data is not None else []
        if not rows:
            raise ValueError("В книге Excel нет строк с данными.")

        header_row = rows[0]
        headers = {
            column_index(cell.get("r", "")): cell_text(cell, shared_strings).strip()
            for cell in header_row.findall(f"{{{XML_NS}}}c")
        }
        target_row = next((row for row in rows if row.get("r") == str(row_number)), None)
        if target_row is None or target_row is header_row:
            raise ValueError("Строка карточки не найдена в книге Excel.")

        cells = {
            column_index(cell.get("r", "")): cell
            for cell in target_row.findall(f"{{{XML_NS}}}c")
        }
        changes_to_apply = []
        for index, value in changes.items():
            header = headers.get(index, "")
            if not header:
                raise ValueError("Нельзя изменить столбец без заголовка.")
            cell = cells.get(index)
            old_value = cell_text(cell, shared_strings) if cell is not None else ""
            if old_value != value:
                changes_to_apply.append((index, header, old_value, value))

        if not changes_to_apply:
            raise ValueError("Нет изменений для сохранения.")

        for index, _, _, value in changes_to_apply:
            cell = cells.get(index)
            if cell is None:
                cell = ET.Element(f"{{{XML_NS}}}c", {"r": f"{column_name(index + 1)}{row_number}"})
                next_cell = next(
                    (item for item in target_row.findall(f"{{{XML_NS}}}c")
                     if column_index(item.get("r", "")) > index),
                    None,
                )
                if next_cell is None:
                    target_row.append(cell)
                else:
                    target_row.insert(list(target_row).index(next_cell), cell)
                cells[index] = cell

            for child in list(cell):
                cell.remove(child)
            if value:
                cell.set("t", "inlineStr")
                inline_string = ET.SubElement(cell, f"{{{XML_NS}}}is")
                text = ET.SubElement(inline_string, f"{{{XML_NS}}}t")
                text.set(XML_SPACE, "preserve")
                text.text = value
            else:
                cell.attrib.pop("t", None)

        worksheet_contents = ET.tostring(worksheet, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(result, "w") as updated:
            updated.comment = original.comment
            for item in original.infolist():
                contents = (
                    worksheet_contents
                    if item.filename == worksheet_path
                    else original.read(item.filename)
                )
                updated.writestr(item, contents)

    return result.getvalue(), changes_to_apply


def safe_log_value(value):
    return value.replace("\\", "\\\\").replace("\r", "\\r").replace("\n", "\\n").replace("|", "\\|")


def make_private_workbook(source):
    result = io.BytesIO()
    with zipfile.ZipFile(source, "r") as original, zipfile.ZipFile(result, "w") as sanitized:
        worksheet_path = first_worksheet_path(original)
        shared_indices_to_redact = set()
        sheet_xml = ET.fromstring(original.read(worksheet_path))
        for row in sheet_xml.iter(f"{{{XML_NS}}}row"):
            for cell in row:
                if cell.tag == f"{{{XML_NS}}}c" and re.fullmatch(r"[FG]\d+", cell.get("r", "")) and cell.get("t") == "s":
                    value = cell.find(f"{{{XML_NS}}}v")
                    if value is not None and value.text is not None:
                        shared_indices_to_redact.add(int(value.text))

        for item in original.infolist():
            contents = original.read(item.filename)
            if item.filename == worksheet_path:
                for row in sheet_xml.iter(f"{{{XML_NS}}}row"):
                    for cell in list(row):
                        if cell.tag == f"{{{XML_NS}}}c" and re.fullmatch(r"[FG]\d+", cell.get("r", "")):
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
        return self.session_username() is not None

    def session_username(self):
        token = self.session_token()
        if not token:
            return None
        session = self.server.sessions.get(token)
        if session is None:
            return None
        if session["expires"] <= time.monotonic():
            self.server.sessions.pop(token, None)
            return None
        session["expires"] = time.monotonic() + SESSION_TTL
        return session["username"]

    def do_GET(self):
        path = normalized_request_path(self.path)
        if path == "/api/session":
            username = self.session_username()
            self.send_json(200, {"authenticated": username is not None, "admin": username == "admin"})
            return
        if path == "/structure.xlsx":
            with self.server.workbook_lock:
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
        if path == "/data" or path.startswith("/data/"):
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
        self.server.sessions[token] = {"username": account, "expires": now + SESSION_TTL}
        self.send_json(200, {"authenticated": True, "admin": account == "admin"}, (
            ("Set-Cookie", f"{SESSION_COOKIE}={token}; Path=/; HttpOnly; SameSite=Lax; Max-Age={SESSION_TTL}"),
        ))

    def do_PUT(self):
        if urlsplit(self.path).path != "/api/employee":
            self.send_json(404, {"error": "Not found"})
            return
        username = self.session_username()
        if username is None:
            self.send_json(401, {"error": "Необходима авторизация."})
            return
        if username != "admin":
            self.send_json(403, {"error": "Редактирование доступно только пользователю admin."})
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            self.send_json(400, {"error": "Некорректный запрос."})
            return
        if content_length <= 0 or content_length > MAX_EDIT_BODY:
            self.send_json(400, {"error": "Некорректный запрос."})
            return
        try:
            payload = json.loads(self.rfile.read(content_length))
            row_number = payload.get("row")
            raw_changes = payload.get("changes")
            if (not isinstance(row_number, int) or isinstance(row_number, bool)
                    or not isinstance(raw_changes, dict) or not raw_changes):
                raise ValueError("Некорректный запрос.")
            changes = {}
            for raw_index, value in raw_changes.items():
                if not isinstance(raw_index, str) or not raw_index.isdecimal():
                    raise ValueError("Некорректный номер столбца.")
                if not isinstance(value, str) or len(value) > MAX_EDIT_VALUE:
                    raise ValueError("Некорректное значение поля.")
                if any(ord(character) < 32 and character not in "\t\n\r" for character in value):
                    raise ValueError("Поле содержит недопустимые символы.")
                changes[int(raw_index)] = value
        except (json.JSONDecodeError, UnicodeDecodeError, AttributeError, ValueError) as error:
            self.send_json(400, {"error": str(error) or "Некорректный запрос."})
            return

        try:
            with self.server.workbook_lock:
                updated_bytes, events = update_workbook(WORKBOOK, row_number, changes)
                with tempfile.NamedTemporaryFile(dir=WORKBOOK.parent, prefix=".structure-", delete=False) as temporary:
                    temporary_path = Path(temporary.name)
                    temporary.write(updated_bytes)
                    temporary.flush()
                    os.fsync(temporary.fileno())
                try:
                    os.replace(temporary_path, WORKBOOK)
                finally:
                    temporary_path.unlink(missing_ok=True)

                timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
                log_lines = [
                    f"{timestamp} | {username} | строка {row_number} | {safe_log_value(label)}: "
                    f"{safe_log_value(old_value)} -> {safe_log_value(new_value)}\n"
                    for _, label, old_value, new_value in events
                ]
                try:
                    CHANGES_LOG.parent.mkdir(parents=True, exist_ok=True)
                    with CHANGES_LOG.open("a", encoding="utf-8") as log_file:
                        log_file.writelines(log_lines)
                        log_file.flush()
                        os.fsync(log_file.fileno())
                except OSError as error:
                    self.send_json(500, {
                        "error": f"Книга сохранена, но не удалось записать журнал изменений: {error}",
                        "saved": True,
                    })
                    return
        except ValueError as error:
            self.send_json(400, {"error": str(error)})
            return
        except (OSError, zipfile.BadZipFile, ET.ParseError) as error:
            self.send_json(500, {"error": f"Не удалось сохранить изменения: {error}"})
            return

        self.send_json(200, {"saved": True, "changedFields": len(events)})

    def translate_path(self, path):
        request_path = normalized_request_path(path)
        if request_path == "/pass.txt" or request_path == "/data" or request_path.startswith("/data/"):
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
    server.workbook_lock = threading.Lock()
    print("Serving directory on 0.0.0.0:8080", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
