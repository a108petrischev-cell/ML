#!/usr/bin/env python3
"""Local, single-user span annotation. No remote services or browser dependencies."""
from __future__ import annotations

import argparse
import io
import json
import secrets
import sqlite3
import sys
import threading
import webbrowser
import zipfile
from contextlib import contextmanager
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "vendor"))
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

LABELS = {"PER", "ORG"}
STATUSES = {"pending", "draft", "done", "review"}
STATUS_NAMES = {"pending": "Не проверено", "draft": "Черновик", "done": "Проверено", "review": "Вопрос"}


class Problem(Exception):
    def __init__(self, message, status=400):
        super().__init__(message)
        self.status = status


def now():
    return datetime.now(timezone.utc).isoformat()


class Store:
    def __init__(self, directory):
        self.directory = Path(directory)
        self.directory.mkdir(parents=True, exist_ok=True, mode=0o700)
        self.uploads = self.directory / "uploads"
        self.uploads.mkdir(exist_ok=True, mode=0o700)
        self.db_path = self.directory / "annotations.sqlite3"
        with self.connect() as db:
            db.executescript("""
                PRAGMA journal_mode=WAL;
                CREATE TABLE IF NOT EXISTS projects (
                    id TEXT PRIMARY KEY, name TEXT NOT NULL, filename TEXT NOT NULL,
                    sheet TEXT NOT NULL, column_number INTEGER NOT NULL,
                    header_row INTEGER NOT NULL, created_at TEXT NOT NULL,
                    last_item INTEGER
                );
                CREATE TABLE IF NOT EXISTS items (
                    id INTEGER PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(id),
                    excel_row INTEGER NOT NULL, text TEXT NOT NULL,
                    status TEXT NOT NULL DEFAULT 'pending',
                    entities TEXT NOT NULL DEFAULT '[]', note TEXT NOT NULL DEFAULT '',
                    revision INTEGER NOT NULL DEFAULT 0, updated_at TEXT,
                    UNIQUE(project_id, excel_row)
                );
                CREATE INDEX IF NOT EXISTS items_queue ON items(project_id, status, excel_row);
            """)
        self.db_path.chmod(0o600)

    @contextmanager
    def connect(self):
        db = sqlite3.connect(self.db_path, timeout=15)
        db.row_factory = sqlite3.Row
        db.execute("PRAGMA foreign_keys=ON")
        try:
            with db:
                yield db
        finally:
            db.close()

    def upload(self, content, filename):
        if not filename.lower().endswith(".xlsx"):
            raise Problem("Нужен файл .xlsx. Сохраните старый .xls или .xlsm как .xlsx.")
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                if sum(x.file_size for x in archive.infolist()) > 500 * 1024 * 1024:
                    raise Problem("Распакованный файл слишком большой (больше 500 МБ).")
            book = load_workbook(io.BytesIO(content), read_only=True, keep_links=False)
            sheets = book.sheetnames
            book.close()
        except Problem:
            raise
        except Exception:
            raise Problem("Не удалось открыть Excel. Нужен обычный .xlsx без пароля.") from None
        token = secrets.token_hex(16)
        path = self.uploads / (token + ".xlsx")
        path.write_bytes(content)
        path.chmod(0o600)
        return {"upload": token, "filename": Path(filename).name, "sheets": sheets}

    def upload_path(self, token):
        if not isinstance(token, str) or len(token) != 32 or any(c not in "0123456789abcdef" for c in token):
            raise Problem("Некорректный идентификатор файла.")
        path = self.uploads / (token + ".xlsx")
        if not path.is_file():
            raise Problem("Файл не найден. Загрузите его заново.", 404)
        return path

    def preview(self, token, sheet, header_row):
        path = self.upload_path(token)
        header_row = integer(header_row, 1, 1000)
        book = load_workbook(path, read_only=True, keep_links=False)
        try:
            if sheet not in book.sheetnames:
                raise Problem("Лист не найден.")
            ws = book[sheet]
            # Reset unreliable dimensions written by some Excel exporters.
            ws.reset_dimensions()
            rows = list(ws.iter_rows(min_row=header_row, max_row=header_row + 4, values_only=True))
            width = min(300, max((len(row) for row in rows), default=0))
            headers = rows[0] if rows else []
            columns = [{"number": i + 1, "letter": get_column_letter(i + 1),
                        "name": str(headers[i]) if i < len(headers) and headers[i] is not None else "Без заголовка"}
                       for i in range(width)]
            samples = [[str(x)[:300] if x is not None else "" for x in row[:width]] for row in rows[1:]]
            return {"columns": columns, "samples": samples}
        finally:
            book.close()

    def create_project(self, payload):
        path = self.upload_path(payload.get("upload"))
        header = integer(payload.get("header_row"), 1, 1000)
        column = integer(payload.get("column"), 1, 300)
        sheet = payload.get("sheet")
        book = load_workbook(path, read_only=True, keep_links=False)
        records = []
        skipped = 0
        try:
            if sheet not in book.sheetnames:
                raise Problem("Лист не найден.")
            ws = book[sheet]
            ws.reset_dimensions()
            for row_number, (cell,) in enumerate(ws.iter_rows(min_row=header + 1, min_col=column, max_col=column), header + 1):
                if row_number > 200_000:
                    raise Problem("Поддерживается до 200 000 строк. Разделите книгу на части.")
                if cell.data_type == "f":
                    raise Problem(f"В строке {row_number} комментарий задан формулой. Скопируйте колонку и вставьте как значения перед загрузкой.")
                if cell.value is None or not str(cell.value).strip():
                    skipped += 1
                    continue
                records.append((row_number, str(cell.value)))
        finally:
            book.close()
        if not records:
            raise Problem("В выбранной колонке нет непустых комментариев.")
        project_id = secrets.token_hex(12)
        filename = str(payload.get("filename", "Таблица.xlsx"))[:250]
        name = str(payload.get("name") or Path(filename).stem).strip()[:150] or "Разметка"
        with self.connect() as db:
            db.execute("INSERT INTO projects VALUES (?,?,?,?,?,?,?,NULL)",
                       (project_id, name, filename, sheet, column, header, now(),))
            db.executemany("INSERT INTO items(project_id,excel_row,text) VALUES (?,?,?)",
                           ((project_id, row, text) for row, text in records))
        return {"id": project_id, "count": len(records), "skipped": skipped}

    def projects(self):
        with self.connect() as db:
            return [dict(row) for row in db.execute("""
                SELECT p.*, count(i.id) AS total,
                  sum(CASE WHEN i.status='done' THEN 1 ELSE 0 END) AS done
                FROM projects p LEFT JOIN items i ON i.project_id=p.id
                GROUP BY p.id ORDER BY p.created_at DESC
            """)]

    def queue(self, project_id, status="all", search=""):
        if status not in STATUSES | {"all", "unfinished"}:
            raise Problem("Неизвестный фильтр.")
        with self.connect() as db:
            project = db.execute("SELECT * FROM projects WHERE id=?", (project_id,)).fetchone()
            if not project:
                raise Problem("Проект не найден.", 404)
            stats = {s: 0 for s in STATUSES}
            stats.update({r["status"]: r["n"] for r in db.execute(
                "SELECT status,count(*) n FROM items WHERE project_id=? GROUP BY status", (project_id,))})
            # Python casefold handles Cyrillic, unlike SQLite's default LIKE.
            needle = search.casefold()
            items = []
            for row in db.execute("SELECT id,excel_row,text,status FROM items WHERE project_id=? ORDER BY excel_row", (project_id,)):
                if status == "unfinished" and row["status"] == "done":
                    continue
                if status in STATUSES and row["status"] != status:
                    continue
                if needle and needle not in row["text"].casefold():
                    continue
                items.append({"id": row["id"], "row": row["excel_row"], "status": row["status"],
                              "preview": row["text"][:120]})
            return {"project": dict(project), "stats": stats, "items": items}

    def item(self, item_id):
        with self.connect() as db:
            row = db.execute("SELECT * FROM items WHERE id=?", (item_id,)).fetchone()
            if not row:
                raise Problem("Комментарий не найден.", 404)
            result = dict(row)
            result["entities"] = json.loads(result["entities"])
            return result

    def save(self, item_id, payload):
        with self.connect() as db:
            db.execute("BEGIN IMMEDIATE")
            row = db.execute("SELECT * FROM items WHERE id=?", (item_id,)).fetchone()
            if not row:
                raise Problem("Комментарий не найден.", 404)
            if payload.get("revision") != row["revision"]:
                raise Problem("Эта строка изменена в другой вкладке. Обновите страницу, прежде чем продолжить.", 409)
            status = payload.get("status")
            if status not in STATUSES:
                raise Problem("Неизвестный статус.")
            entities = validate_entities(row["text"], payload.get("entities"))
            if status == "pending" and entities:
                status = "draft"
            note = payload.get("note", "")
            if not isinstance(note, str) or len(note) > 5000:
                raise Problem("Заметка должна содержать не больше 5000 символов.")
            db.execute("UPDATE items SET entities=?,status=?,note=?,revision=revision+1,updated_at=? WHERE id=?",
                       (json.dumps(entities, ensure_ascii=False), status, note, now(), item_id))
            db.execute("UPDATE projects SET last_item=? WHERE id=?", (item_id, row["project_id"]))
        return self.item(item_id)

    def resume(self, project_id, item_id):
        with self.connect() as db:
            if not db.execute("SELECT 1 FROM items WHERE id=? AND project_id=?", (item_id, project_id)).fetchone():
                raise Problem("Комментарий не найден.", 404)
            db.execute("UPDATE projects SET last_item=? WHERE id=?", (item_id, project_id))
        return {"ok": True}

    def export(self, project_id, kind):
        with self.connect() as db:
            project = db.execute("SELECT * FROM projects WHERE id=?", (project_id,)).fetchone()
            if not project:
                raise Problem("Проект не найден.", 404)
            rows = list(db.execute("SELECT * FROM items WHERE project_id=? ORDER BY excel_row", (project_id,)))
        if kind == "jsonl":
            lines = []
            for row in rows:
                if row["status"] != "done":
                    continue
                entities = validate_entities(row["text"], json.loads(row["entities"]))
                lines.append(json.dumps({"id": row["id"], "text": row["text"], "entities": entities,
                    "source": {"file": project["filename"], "sheet": project["sheet"], "row": row["excel_row"]}}, ensure_ascii=False))
            if not lines:
                raise Problem("Пока нет проверенных строк. Нажмите «Готово» хотя бы для одного комментария.")
            return ("\n".join(lines) + "\n").encode("utf-8"), "application/x-ndjson; charset=utf-8", "training.jsonl"
        if kind != "xlsx":
            raise Problem("Неизвестный формат.")
        # Separate review workbook: never open/save over the source workbook.
        book = Workbook()
        ws = book.active
        ws.title = "Разметка"
        ws.append(["Строка Excel", "Комментарий", "Статус", "ФИО", "Организации", "Разметка JSON", "Заметка"])
        spans = book.create_sheet("Сущности")
        spans.append(["Строка Excel", "Тип", "Текст", "Начало", "Конец"])
        for row in rows:
            entities = json.loads(row["entities"])
            serialized = json.dumps(entities, ensure_ascii=False)
            if len(serialized) > 32767:
                serialized = "Разметка длиннее лимита ячейки Excel. Все сущности — на листе «Сущности» и в JSONL."
            ws.append([row["excel_row"], row["text"], STATUS_NAMES[row["status"]],
                       "\n".join(e["text"] for e in entities if e["label"] == "PER"),
                       "\n".join(e["text"] for e in entities if e["label"] == "ORG"),
                       serialized, row["note"]])
            for entity in entities:
                spans.append([row["excel_row"], entity["label"], entity["text"], entity["start"], entity["end"]])
        for sheet in book:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="273D38")
            for row in sheet.iter_rows(min_row=2):
                for cell in row:
                    # Treat user text as literal, including strings beginning '='.
                    if isinstance(cell.value, str):
                        cell.data_type = "s"
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
            for i in range(1, sheet.max_column + 1):
                sheet.column_dimensions[get_column_letter(i)].width = 18 if i == 1 else 35
        ws.column_dimensions["B"].width = 80
        ws.column_dimensions["F"].width = 60
        output = io.BytesIO()
        book.save(output)
        return output.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "annotations.xlsx"


def integer(value, low, high):
    try:
        number = int(value)
    except (ValueError, TypeError):
        raise Problem("Неверный номер строки или колонки.") from None
    if not low <= number <= high:
        raise Problem(f"Номер должен быть от {low} до {high}.")
    return number


def validate_entities(text, entities):
    if not isinstance(entities, list):
        raise Problem("Некорректная разметка.")
    result = []
    for entity in entities:
        if not isinstance(entity, dict):
            raise Problem("Некорректная сущность.")
        start, end, label = entity.get("start"), entity.get("end"), entity.get("label")
        if type(start) is not int or type(end) is not int or not 0 <= start < end <= len(text) or label not in LABELS:
            raise Problem("Неверные границы или тип сущности.")
        fragment = text[start:end]
        if not fragment.strip() or fragment != fragment.strip():
            raise Problem("Выделение не должно содержать пробелы по краям.")
        if entity.get("text", fragment) != fragment:
            raise Problem("Текст сущности не совпадает с её границами.")
        result.append({"start": start, "end": end, "label": label, "text": fragment})
    result.sort(key=lambda e: (e["start"], e["end"]))
    if any(a["end"] > b["start"] for a, b in zip(result, result[1:])):
        raise Problem("Выделения пересекаются. Удалите старое выделение перед заменой.")
    return result


class Handler(BaseHTTPRequestHandler):
    server_version = "LocalAnnotator"

    def log_message(self, *_):
        pass  # Do not log comments, filenames, or request parameters.

    def send(self, body, content_type="application/json; charset=utf-8", status=200, filename=None):
        if not isinstance(body, bytes):
            body = json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Content-Security-Policy", "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'")
        if filename:
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        self.handle_request(False)

    def do_POST(self):
        self.handle_request(True)

    def handle_request(self, mutation):
        try:
            expected = f"127.0.0.1:{self.server.server_port}"
            if self.headers.get("Host") != expected:
                raise Problem("Используйте локальный адрес 127.0.0.1.", 403)
            if self.headers.get("Sec-Fetch-Site") == "cross-site":
                raise Problem("Запрос с другого сайта запрещён.", 403)
            if mutation:
                if self.headers.get("Origin") not in (None, f"http://{expected}") or self.headers.get("X-Local-Token") != self.server.token:
                    raise Problem("Доступ запрещён. Обновите страницу программы.", 403)
            parsed = urlparse(self.path)
            path = parsed.path
            query = {k: v[0] for k, v in parse_qs(parsed.query).items()}
            store = self.server.store
            if not mutation:
                static = {"/": ("index.html", "text/html; charset=utf-8"), "/app.js": ("app.js", "text/javascript; charset=utf-8"),
                          "/style.css": ("style.css", "text/css; charset=utf-8")}
                if path in static:
                    file, mime = static[path]
                    return self.send((ROOT / "static" / file).read_bytes(), mime)
                if path == "/api/bootstrap":
                    return self.send({"token": self.server.token, "projects": store.projects()})
                if path == "/api/preview":
                    return self.send(store.preview(query.get("upload"), query.get("sheet"), query.get("header_row")))
                if path == "/api/queue":
                    return self.send(store.queue(query.get("project"), query.get("status", "all"), query.get("search", "")))
                if path == "/api/item":
                    return self.send(store.item(integer(query.get("id"), 1, 2**53)))
                if path == "/api/export":
                    body, mime, filename = store.export(query.get("project"), query.get("kind"))
                    return self.send(body, mime, filename=filename)
            else:
                size = integer(self.headers.get("Content-Length"), 1, 100 * 1024 * 1024)
                raw = self.rfile.read(size)
                if path == "/api/upload":
                    return self.send(store.upload(raw, unquote(self.headers.get("X-Filename", "table.xlsx"))))
                try:
                    payload = json.loads(raw)
                except (ValueError, UnicodeDecodeError):
                    raise Problem("Некорректный запрос.") from None
                if not isinstance(payload, dict):
                    raise Problem("Некорректный запрос.")
                if path == "/api/projects":
                    return self.send(store.create_project(payload))
                if path == "/api/save":
                    return self.send(store.save(integer(payload.get("id"), 1, 2**53), payload))
                if path == "/api/resume":
                    return self.send(store.resume(payload.get("project"), integer(payload.get("id"), 1, 2**53)))
            raise Problem("Страница не найдена.", 404)
        except Problem as exc:
            self.send({"error": str(exc)}, status=exc.status)
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception:
            self.send({"error": "Не удалось выполнить действие. Данные разметки сохранены до последнего успешного изменения."}, status=500)


def make_server(directory, port=0):
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    server.daemon_threads = True
    server.store = Store(directory)
    server.token = secrets.token_urlsafe(32)
    return server


def main():
    parser = argparse.ArgumentParser(description="Локальная разметка ФИО и организаций")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--data-dir", type=Path, default=ROOT / "data")
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()
    try:
        server = make_server(args.data_dir, args.port)
    except OSError:
        print(f"Порт {args.port} занят. Закройте другой экземпляр или запустите с --port 8766.")
        return 1
    url = f"http://127.0.0.1:{server.server_port}"
    print(f"Разметка работает: {url}\nДанные: {args.data_dir.resolve()}\nДля остановки нажмите Ctrl+C.", flush=True)
    if not args.no_browser:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nПрограмма остановлена. Разметка сохранена.")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
