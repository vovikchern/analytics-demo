#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
demo_server.py — автономный сервер демо-версии (репозиторий analytics-demo).

Заменяет продовый стек (n8n + NocoDB + 1С) на:
  - синтетические данные (data/demo_records.json — события в формате выгрузки 1С);
  - справочник стоимостей (data/demo_terms.json);
  - in-memory авторизацию (демо-пользователи, пароль у всех Demo1234!).

Запуск:  python demo_server.py  →  http://127.0.0.1:8000
Логин:   admin@demo.ru / buh@demo.ru / manager@demo.ru, пароль Demo1234!

Агрегация событий по датам и логика дашборда — те же модули, что в проде
(scripts/transform_reo_v1.0.py, report_reo.py, generate_excel_reo_v1.0.py),
поэтому демо честно показывает рабочую логику, а не её имитацию.
"""
import base64
import importlib.util
import io
import json
import os
import sys
import uuid
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(HERE, "scripts")
DATA = os.path.join(HERE, "data")
sys.path.insert(0, SCRIPTS)


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


transform = load_module("reo_transform", os.path.join(SCRIPTS, "transform_reo_v1.0.py"))
report_reo = load_module("reo_report", os.path.join(SCRIPTS, "report_reo.py"))
gen_excel = load_module("reo_gen_excel", os.path.join(SCRIPTS, "generate_excel_reo_v1.0.py"))

import openpyxl  # noqa: E402

# ---------------------------------------------------------------- данные
with open(os.path.join(DATA, "demo_records.json"), encoding="utf-8") as f:
    RECORDS = json.load(f)
with open(os.path.join(DATA, "demo_terms.json"), encoding="utf-8") as f:
    TERMS_LIST = json.load(f)
with open(os.path.join(DATA, "demo_meta.json"), encoding="utf-8") as f:
    META = json.load(f)

SNAP_DATES = META["snapshot_dates"]            # плановые даты (scheduled)
MAX_ALLOWED = META["max_allowed_date"]

RULES = [{"contract_no": t["contract_no"], "keyword": t["ng_keyword"],
          "object_name": t["object_name"]} for t in TERMS_LIST]
NG_ROWS = [{"object_name": t["object_name"], "site": t["site"]} for t in TERMS_LIST]
TERMS = {t["object_name"]: t for t in TERMS_LIST}

# предрасчёт снапшотов по всем плановым датам (триггер-источник scheduled)
HISTORY = {}
for d in SNAP_DATES:
    rows, _stats, _samples = transform.aggregate(RECORDS, d, RULES)
    HISTORY[d] = {"scheduled": rows}


def pick_snapshot(requested):
    """последний снапшот <= даты (как в проде: дата = дате данных)."""
    older = [d for d in SNAP_DATES if d <= requested]
    if older:
        return older[-1], False
    return SNAP_DATES[0], True


def terms_get():
    return [{"object_name": t["object_name"], "contract_no": t["contract_no"],
             "contract_cost": t["contract_cost"], "advance_limit": t["advance_limit"],
             "active": True, "completed": bool(t.get("completed"))}
            for t in sorted(TERMS_LIST, key=lambda x: x["object_name"])]


def make_excel(actual_date):
    """Заполнение template.xlsx снапшотом на дату (логика generate_excel)."""
    rows = HISTORY[actual_date]["scheduled"]
    by_obj_period, bg_by_contract = {}, {}
    for f in rows:
        if f.get("period") == "contract":
            bg_by_contract[f.get("contract_no")] = f.get("bg_retention") or 0
        else:
            by_obj_period[(f.get("object_name"), f.get("period"))] = f

    def db_value(obj, tpl_period, field):
        if tpl_period == "2025":
            return sum((by_obj_period.get((obj, p), {}).get(field) or 0)
                       for p in ("2025H1", "2025H2"))
        return by_obj_period.get((obj, tpl_period), {}).get(field) or 0

    wb = openpyxl.load_workbook(os.path.join(HERE, "template.xlsx"), data_only=False)
    ws = wb.worksheets[0]
    ws["B1"] = datetime.strptime(actual_date, "%Y-%m-%d").date()
    ws["B1"].number_format = "DD.MM.YYYY"

    # бухгалтерский формат шаблона; ячейки, пустые в шаблоне (например L5/N5),
    # стиля не имеют — иначе число отображается как General (1417714,5)
    ACC_FMT = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-'

    def put(row, col, value):
        c = ws.cell(row=row, column=col)
        c.value = value
        if isinstance(value, (int, float)) and c.number_format in (None, "General"):
            c.number_format = ACC_FMT

    for obj, periods in gen_excel.SECTIONS.items():
        t = TERMS.get(obj) or {}
        for tpl_period, row in periods.items():
            label = gen_excel.LABELS.get(tpl_period, tpl_period)
            ws.cell(row=row, column=4).value = int(label) if label.isdigit() else label
            for field, col in gen_excel.VALUE_COLS.items():
                put(row, col, round(db_value(obj, tpl_period, field), 2))
        ws.cell(row=gen_excel.I_ROW_OF[obj], column=4).value = "итого"
        if t:
            put(gen_excel.B_ROWS[obj], 2, round(t["contract_cost"], 2))
            ws.cell(row=gen_excel.I_ROW_OF[obj], column=9).value = t["advance_limit"]
        contract_no = t.get("contract_no")
        if obj in gen_excel.C_TOP and contract_no:
            put(gen_excel.C_TOP[obj], 3, round(bg_by_contract.get(contract_no, 0.0), 2))
    for r, obj in gen_excel.CONTRACT_I_ROWS.items():
        t = TERMS.get(obj) or {}
        if t:
            ws.cell(row=r, column=9).value = t["advance_limit"]

    buf = io.BytesIO()
    wb.save(buf)
    return base64.b64encode(buf.getvalue()).decode("ascii")


# ---------------------------------------------------------------- auth (in-memory)
USERS = {
    u["email"]: u for u in [
        {"email": "admin@demo.ru", "display_name": "Администратор", "role": "admin"},
        {"email": "buh@demo.ru", "display_name": "Главный бухгалтер", "role": "accounting"},
        {"email": "manager@demo.ru", "display_name": "Менеджер", "role": "management"},
    ]
}
PASSWORD = "Demo1234!"
SESSIONS = {}          # token -> email


def auth_handle(payload):
    action = payload.get("action")
    token = payload.get("token") or ""
    if action == "login":
        email = (payload.get("email") or "").strip().lower()
        u = USERS.get(email)
        if not u or payload.get("password") != PASSWORD:
            return {"status": "error", "message": "Неверный email или пароль", "code": 401}
        token = str(uuid.uuid4())
        SESSIONS[token] = email
        return {"status": "success", "data": {
            "token": token, "role": u["role"], "display_name": u["display_name"],
            "email": u["email"], "must_change_password": False}}
    if action == "check":
        email = SESSIONS.get(token)
        if not email:
            return {"status": "error", "message": "Сессия недействительна", "code": 401}
        u = USERS[email]
        return {"status": "success", "data": {
            "role": u["role"], "display_name": u["display_name"],
            "email": u["email"], "must_change_password": False}}
    if action == "logout":
        SESSIONS.pop(token, None)
        return {"status": "success", "data": {"logged_out": True}}
    if action == "users_list":
        email = SESSIONS.get(token)
        if not email or USERS[email]["role"] != "admin":
            return {"status": "error", "message": "Только администратор", "code": 403}
        return {"status": "success", "data": {"users": [
            {"email": e, "display_name": u["display_name"], "role": u["role"],
             "active": True, "must_change_password": False, "has_password": True}
            for e, u in sorted(USERS.items())]}}
    return {"status": "error", "message": f"Неизвестное действие: {action}", "code": 400}


# ---------------------------------------------------------------- report-reo
def reo_handle(payload):
    action = payload.get("action") or ""
    token = payload.get("token") or ""
    email = SESSIONS.get(token)
    if not email:
        return {"status": "error", "message": "Сессия недействительна — войдите заново", "code": 401}
    role = USERS[email]["role"]
    if role not in report_reo.REO_ROLES:
        return {"status": "error", "message": "Нет доступа к отчёту", "code": 403}

    if action == "terms_get":
        return {"status": "success", "data": {"terms": terms_get()}}
    if action == "terms_set":
        if role not in ("admin", "accounting"):
            return {"status": "error",
                    "message": "Редактирование стоимостей доступно только бухгалтерии и администратору",
                    "code": 403}
        changed = 0
        for u in payload.get("updates") or []:
            t = TERMS.get((u.get("object_name") or "").strip())
            if not t:
                continue
            if "contract_cost" in u and isinstance(u["contract_cost"], (int, float)):
                t["contract_cost"] = round(float(u["contract_cost"]), 2)
                changed += 1
            if "advance_limit" in u and isinstance(u["advance_limit"], (int, float)):
                t["advance_limit"] = round(float(u["advance_limit"]), 2)
                changed += 1
            if "completed" in u:
                t["completed"] = bool(u["completed"])
                changed += 1
        with open(os.path.join(DATA, "demo_terms.json"), "w", encoding="utf-8") as f:
            json.dump(TERMS_LIST, f, ensure_ascii=False, indent=1)
        return {"status": "success", "data": {"changed": changed, "terms": terms_get()}}
    if action == "fetch":
        return {"status": "error",
                "message": "В демо-версии подгрузка из 1С недоступна — выберите дату из истории",
                "code": 400}

    fmt = (payload.get("format") or "all").lower()
    requested = (payload.get("date") or "").strip() or MAX_ALLOWED
    actual_date, fallback = pick_snapshot(requested)
    snap_rows = HISTORY[actual_date]["scheduled"]
    dashboard = report_reo.build_dashboard(snap_rows, HISTORY, actual_date, TERMS, NG_ROWS)
    meta = {"requested_date": requested, "actual_date": actual_date,
            "fallback": fallback, "available_dates": SNAP_DATES,
            "scheduled_dates": SNAP_DATES, "max_allowed_date": MAX_ALLOWED,
            "trigger_source": "scheduled"}
    out = {"meta": meta, "snapshot": {"report_date": actual_date, "rows": snap_rows}}
    if fmt in ("dashboard", "all"):
        out["dashboard"] = dashboard
    if fmt in ("excel", "all"):
        b64 = make_excel(actual_date)
        out["excel"] = {"file_name": f"Demo_Report_{actual_date}.xlsx",
                        "file_data": b64, "report_date": actual_date,
                        "requested_date": requested, "cached": False}
    return {"status": "success", "data": out}


# ---------------------------------------------------------------- http
class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        data = body if isinstance(body, bytes) else json.dumps(
            body, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        path = urlparse(self.path).path
        if path in ("/", "/index.html"):
            return self._static("index.html", "text/html; charset=utf-8")
        if path == "/logo.svg":
            return self._static("logo.svg", "image/svg+xml")
        return self._send(404, {"error": "not found"})

    def _static(self, name, ctype):
        try:
            with open(os.path.join(HERE, name), "rb") as f:
                return self._send(200, f.read(), ctype)
        except FileNotFoundError:
            return self._send(404, {"error": "not found"})

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            length = int(self.headers.get("Content-Length") or 0)
            payload = json.loads(self.rfile.read(length).decode("utf-8") or "{}")
        except ValueError:
            return self._send(400, {"status": "error", "message": "Некорректный JSON"})
        if path == "/webhook/auth":
            return self._send(200, auth_handle(payload))
        if path == "/webhook/report-reo":
            return self._send(200, reo_handle(payload))
        return self._send(200, {"status": "error",
                                "message": "В демо-версии этот раздел недоступен", "code": 400})

    def log_message(self, *args):   # тише в консоли
        pass


if __name__ == "__main__":
    import threading
    import webbrowser
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    port = int(args[0]) if args else 8000
    url = f"http://127.0.0.1:{port}"
    print(f"Demo server: {url}  (Ctrl+C — выход)")
    print("Логины: admin@demo.ru / buh@demo.ru / manager@demo.ru, пароль Demo1234!")
    if "--no-browser" not in sys.argv:
        # открыть браузер после подъёма сервера (запуск двойным кликом)
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        ThreadingHTTPServer(("127.0.0.1", port), Handler).serve_forever()
    except KeyboardInterrupt:
        print("\nСервер остановлен")
