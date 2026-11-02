# -*- coding: utf-8 -*-
"""
auth.py — авторизация портала аналитики СтройИнжиниринг (этап 4а).

Запускается в python_worker (docker exec ... python /data/scripts/auth.py '<json>').
n8n-воркфлоу [DEV]_IN_Auth вызывает через SSH-ноду и возвращает ответ фронтенду.

Вход:  JSON (argv[1] или stdin):
  {"action": "login",            "email": "...", "password": "..."}
  {"action": "check",            "token": "..."}
  {"action": "change_password",  "token": "...", "old_password": "...", "new_password": "..."}
  {"action": "set_password",     "token": "...", "email": "...", "new_password": "..."}   (только admin)
  {"action": "logout",           "token": "..."}

Выход: одна строка JSON:
  {"status": "success", "data": {...}}
  {"status": "error",   "message": "...", "code": 401|403|400|500}

env: NOCODB_URL, NOCODB_TOKEN, NOCODB_BASE (как в transform_reo_v1.0.py),
     PORTAL_USERS_TABLE (по умолчанию "portal_users").

Хеш: sha256(salt + password). Токен: uuid4, хранится в portal_sessions;
скользящий TTL 8 часов (каждый проверенный запрос продлевает сессию).
"""

import hashlib
import json
import os
import secrets
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone

USERS_TABLE = os.environ.get("PORTAL_USERS_TABLE", "portal_users")
SESSIONS_TABLE = os.environ.get("PORTAL_SESSIONS_TABLE", "portal_sessions")
BASE_TITLE = os.environ.get("NOCODB_BASE", "Analytics-Portal")
SESSION_TTL_HOURS = 8           # скользящий TTL сессии (решение 2026-11-02, п.8)
MIN_PASSWORD_LEN = 6

ROLES = ("admin", "accounting", "management", "procurement", "manager")

# алфавит генератора паролей — без похожих символов (0/O, 1/l/I)
PWD_ALPHABET = "abcdefghjkmnpqrstuvwxyzABCDEFGHJKMNPQRSTUVWXYZ23456789"


def gen_password(length=10):
    return "".join(secrets.choice(PWD_ALPHABET) for _ in range(length))


# ---------------------------------------------------------------- NocoDB (v2/v1 fallback)
class Noco:
    def __init__(self):
        self.base_url = os.environ.get("NOCODB_URL", "").rstrip("/")
        self.token = os.environ.get("NOCODB_TOKEN", "")
        self.base_id = None
        self.tables = {}

    def request(self, method, path, params=None, body=None):
        url = self.base_url + path
        if params:
            url += "?" + urllib.parse.urlencode(params)
        data = json.dumps(body).encode("utf-8") if body is not None else None
        req = urllib.request.Request(url, data=data, method=method)
        req.add_header("xc-token", self.token)
        req.add_header("Content-Type", "application/json")
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                raw = resp.read().decode("utf-8")
                return json.loads(raw) if raw else {}
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", "replace")[:300]
            raise RuntimeError(f"NocoDB {method} {path} -> HTTP {e.code}: {detail}")

    def connect(self):
        bases = self.request("GET", "/api/v2/meta/bases")
        found = None
        for b in bases.get("list", []):
            if (b.get("title") or "").lower() == BASE_TITLE.lower():
                found = b
                break
        if not found:
            titles = [b.get("title") for b in bases.get("list", [])]
            raise RuntimeError(f"Base '{BASE_TITLE}' not found. Bases: {titles}")
        self.base_id = found["id"]
        meta = self.request("GET", f"/api/v2/meta/bases/{self.base_id}/tables")
        for t in meta.get("list", []):
            self.tables[t["title"]] = t["id"]
        if USERS_TABLE not in self.tables:
            raise RuntimeError(f"Table '{USERS_TABLE}' not found in base '{BASE_TITLE}'")
        if SESSIONS_TABLE not in self.tables:
            raise RuntimeError(f"Table '{SESSIONS_TABLE}' not found in base '{BASE_TITLE}'")

    def _try(self, variants):
        last_err = None
        for method, path, params, body in variants:
            try:
                return self.request(method, path, params=params, body=body)
            except RuntimeError as e:
                last_err = e
                if "HTTP 404" in str(e) or "HTTP 405" in str(e):
                    continue
                raise
        raise last_err

    def list_rows(self, table_title):
        tid = self.tables[table_title]
        rows, offset = [], 0
        while True:
            params = {"limit": 500, "offset": offset}
            resp = self._try([
                ("GET", f"/api/v2/tables/{tid}/rows", params, None),
                ("POST", f"/api/v2/tables/{tid}/rows/list", None, params),
                ("GET", f"/api/v1/db/data/v1/{self.base_id}/{tid}", params, None),
            ])
            if isinstance(resp, list):
                chunk, page_info = resp, {}
            else:
                chunk = resp.get("list", [])
                page_info = resp.get("pageInfo", {})
            rows.extend(chunk)
            is_last = page_info.get("isLastPage", len(chunk) < 500)
            if is_last or not chunk:
                break
            offset = page_info.get("offset", offset + 500)
        return rows

    def update_row(self, table_title, row_id, fields):
        tid = self.tables[table_title]
        self._try([
            ("PATCH", f"/api/v2/tables/{tid}/rows/{row_id}", None, fields),
            ("PATCH", f"/api/v1/db/data/v1/{self.base_id}/{tid}/{row_id}", None, fields),
        ])

    def delete_row(self, table_title, row_id):
        tid = self.tables[table_title]
        self._try([
            ("DELETE", f"/api/v2/tables/{tid}/rows/{row_id}", None, None),
            ("DELETE", f"/api/v1/db/data/v1/{self.base_id}/{tid}/{row_id}", None, None),
        ])

    def insert_row(self, table_title, fields):
        tid = self.tables[table_title]
        self._try([
            ("POST", f"/api/v2/tables/{tid}/rows", None, fields),
            ("POST", f"/api/v1/db/data/v1/{self.base_id}/{tid}", None, fields),
        ])


def fields_of(row):
    """Строки NocoDB: v2 — плоские поля, v1 — обёртка {"Id":..,"fields":{...}}."""
    f = row.get("fields", row)
    return {k: v for k, v in f.items() if not k.startswith("__")}


def row_id_of(row):
    return row.get("Id") if isinstance(row, dict) else None


# ---------------------------------------------------------------- хеши/токены
def make_salt():
    return secrets.token_hex(16)


def hash_password(salt, password):
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


def make_token():
    return secrets.token_hex(32)


def now_utc():
    return datetime.now(timezone.utc)


def fmt_dt(dt):
    # Явное смещение +00:00 ОБЯЗАТЕЛЬНО: NocoDB трактует наивную строку как
    # локальное время своего контейнера (MSK, +3) и хранит на 3 часа раньше.
    # С явным смещением значение сохраняется как есть (проверено probe_tz.py,
    # 2026-11-02). parse_dt() хвост "+00:00" отрезает — совместимо.
    return dt.strftime("%Y-%m-%d %H:%M:%S+00:00")


def parse_dt(s):
    if not s:
        return None
    s = str(s).strip().replace("T", " ")
    if s.endswith("Z"):
        s = s[:-1]
    if "+" in s:            # хвост "+00:00" (NocoDB нормализует dateTime в UTC)
        s = s.split("+")[0]
    s = s.split(".")[0].strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).replace(tzinfo=timezone.utc)
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------- действия
def find_user(nc, email=None, token=None):
    rows = nc.list_rows(USERS_TABLE)
    for row in rows:
        rid = row_id_of(row)
        f = fields_of(row)
        if email is not None and (f.get("email") or "").strip().lower() == email.strip().lower():
            return rid, f
        if token is not None and (f.get("token") or "") == token:
            return rid, f
    return None, None


def action_login(nc, p):
    email = (p.get("email") or "").strip()
    password = p.get("password") or ""
    if not email or not password:
        return err(400, "Укажите email и пароль")
    rid, f = find_user(nc, email=email)
    if rid is None:
        return err(401, "Неверный email или пароль")
    if not f.get("active", True):
        return err(403, "Учётная запись отключена")
    stored = f.get("password_hash") or ""
    salt = f.get("salt") or ""
    if not stored or hash_password(salt, password) != stored:
        return err(401, "Неверный email или пароль")
    cleanup_expired_sessions(nc)
    token = make_token()
    expires = now_utc() + timedelta(hours=SESSION_TTL_HOURS)
    # сессия — отдельная строка в portal_sessions: несколько входов
    # (браузер/CLI/телефон) живут параллельно (решение 2026-11-01)
    nc.insert_row(SESSIONS_TABLE, {
        "token": token,
        "email": f.get("email") or "",
        "expires_at": fmt_dt(expires),
    })
    return ok({
        "token": token,
        "expires": fmt_dt(expires),
        "role": f.get("role") or "",
        "display_name": f.get("display_name") or "",
        "email": f.get("email") or "",
        "must_change_password": bool(f.get("must_change_password")),
    })


def cleanup_expired_sessions(nc):
    """Удаление просроченных сессий (вызывается при каждом входе)."""
    now = now_utc()
    for row in nc.list_rows(SESSIONS_TABLE):
        f = fields_of(row)
        exp = parse_dt(f.get("expires_at"))
        if exp is None or exp < now:
            nc.delete_row(SESSIONS_TABLE, row_id_of(row))


def action_check(nc, p):
    rid, f, = _by_token(nc, p.get("token"))
    if rid is None:
        return err(401, "Сессия недействительна")
    return ok({
        "role": f.get("role") or "",
        "display_name": f.get("display_name") or "",
        "email": f.get("email") or "",
        "must_change_password": bool(f.get("must_change_password")),
    })


def action_change_password(nc, p):
    rid, f = _by_token(nc, p.get("token"))
    if rid is None:
        return err(401, "Сессия недействительна")
    old_password = p.get("old_password") or ""
    new_password = p.get("new_password") or ""
    if len(new_password) < MIN_PASSWORD_LEN:
        return err(400, f"Пароль должен быть не короче {MIN_PASSWORD_LEN} символов")
    stored = f.get("password_hash") or ""
    if hash_password(f.get("salt") or "", old_password) != stored:
        return err(401, "Текущий пароль указан неверно")
    new_salt = make_salt()
    nc.update_row(USERS_TABLE, rid, {
        "salt": new_salt,
        "password_hash": hash_password(new_salt, new_password),
        "must_change_password": False,
    })
    return ok({"changed": True, "email": f.get("email") or ""})


def action_users_list(nc, p):
    """Список пользователей для админ-панели (без хешей)."""
    rid, f = _by_token(nc, p.get("token"))
    if rid is None:
        return err(401, "Сессия недействительна")
    if (f.get("role") or "") != "admin":
        return err(403, "Только администратор")
    users = []
    for row in nc.list_rows(USERS_TABLE):
        uf = fields_of(row)
        users.append({
            "email": (uf.get("email") or "").strip(),
            "display_name": uf.get("display_name") or "",
            "role": uf.get("role") or "",
            "active": bool(uf.get("active", True)),
            "must_change_password": bool(uf.get("must_change_password")),
            "has_password": bool((uf.get("password_hash") or "").strip()),
        })
    users.sort(key=lambda u: u["email"])
    return ok({"users": users})


def action_set_password(nc, p):
    rid, f = _by_token(nc, p.get("token"))
    if rid is None:
        return err(401, "Сессия недействительна")
    if (f.get("role") or "") != "admin":
        return err(403, "Только администратор может устанавливать пароли")
    email = (p.get("email") or "").strip()
    new_password = p.get("new_password") or ""
    generated = False
    if not new_password:
        new_password = gen_password()      # пароль не передан — генерируем на сервере
        generated = True
    if not email:
        return err(400, "Укажите email пользователя")
    if len(new_password) < MIN_PASSWORD_LEN:
        return err(400, f"Пароль должен быть не короче {MIN_PASSWORD_LEN} символов")
    tid_rid, _ = find_user(nc, email=email)
    if tid_rid is None:
        return err(404, f"Пользователь {email} не найден")
    new_salt = make_salt()
    nc.update_row(USERS_TABLE, tid_rid, {
        "salt": new_salt,
        "password_hash": hash_password(new_salt, new_password),
        "must_change_password": True,
    })
    # принудительный разлогон пользователя: все его сессии удаляются
    for row in nc.list_rows(SESSIONS_TABLE):
        sf = fields_of(row)
        if (sf.get("email") or "").strip().lower() == email.strip().lower():
            nc.delete_row(SESSIONS_TABLE, row_id_of(row))
    return ok({
        "email": email,
        "must_change_password": True,
        "generated": generated,
        # пароль возвращается ТОЛЬКО при серверной генерации и только один раз
        **({"password": new_password} if generated else {}),
    })


def action_logout(nc, p):
    token = p.get("token") or ""
    deleted = 0
    if token:
        for row in nc.list_rows(SESSIONS_TABLE):
            if (fields_of(row).get("token") or "") == token:
                nc.delete_row(SESSIONS_TABLE, row_id_of(row))
                deleted += 1
    return ok({"logged_out": True, "sessions_deleted": deleted})


def _by_token(nc, token, check_expiry=True):
    """-> (user_row_id, user_fields) по токену из portal_sessions.
    Возвращает ИД строки ПОЛЬЗОВАТЕЛЯ (не сессии) — вызывающий код
    обновляет пользователя (смена пароля) по этому ИД.

    Скользящий TTL (решение 2026-11-02, п.8): каждый проверенный запрос
    продлевает сессию до SESSION_TTL_HOURS с момента последнего запроса.
    Продление пишется в NocoDB не чаще раза в час (экономия записей):
    если сессии осталось больше SESSION_TTL_HOURS - 1 ч — не трогаем.
    Закрытая вкладка: сессия живёт до 8 ч после последнего запроса."""
    if not token:
        return None, None
    email = None
    session_row = None
    for row in nc.list_rows(SESSIONS_TABLE):
        sf = fields_of(row)
        if (sf.get("token") or "") != token:
            continue
        if check_expiry:
            exp = parse_dt(sf.get("expires_at"))
            if exp is None or exp < now_utc():
                return None, None
        session_row = row
        email = (sf.get("email") or "").strip()
        break
    if not email:
        return None, None
    rid, f = find_user(nc, email=email)
    if rid is None or not f.get("active", True):
        return None, None
    if session_row is not None:
        exp = parse_dt(fields_of(session_row).get("expires_at"))
        now = now_utc()
        if exp is None or (exp - now) < timedelta(hours=SESSION_TTL_HOURS - 1):
            nc.update_row(SESSIONS_TABLE, row_id_of(session_row),
                          {"expires_at": fmt_dt(now + timedelta(hours=SESSION_TTL_HOURS))})
    return rid, f


# ---------------------------------------------------------------- каркас
def ok(data):
    return {"status": "success", "data": data}


def err(code, message):
    return {"status": "error", "message": message, "code": code}


ACTIONS = {
    "login": action_login,
    "check": action_check,
    "change_password": action_change_password,
    "set_password": action_set_password,
    "logout": action_logout,
    "users_list": action_users_list,
}


def main():
    try:
        raw = sys.argv[1] if len(sys.argv) > 1 else sys.stdin.read()
        payload = json.loads(raw)
        action = payload.get("action")
        fn = ACTIONS.get(action)
        if fn is None:
            print(json.dumps(err(400, f"Неизвестное действие: {action}"), ensure_ascii=False))
            return
        nc = Noco()
        nc.connect()
        result = fn(nc, payload)
        print(json.dumps(result, ensure_ascii=False))
    except Exception as e:  # noqa: BLE001 — наружу всегда JSON
        print(json.dumps(err(500, f"{type(e).__name__}: {e}"), ensure_ascii=False))


if __name__ == "__main__":
    main()
