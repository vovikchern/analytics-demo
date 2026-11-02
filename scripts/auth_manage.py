# -*- coding: utf-8 -*-
"""
auth_manage.py — управление пользователями portal_users (этап 4а).

Запускается в python_worker: docker exec python_worker python /data/scripts/auth_manage.py <команда>

Команды:
  init                          — для всех строк без password_hash: сгенерировать salt и
                                  случайный пароль, записать хеш, вывести "email  пароль"
                                  (список показывается ОДИН раз — раздать сотрудникам)
  set-password <email> [пароль] — принудительно установить пароль (must_change_password=1,
                                  активные сессии сбрасываются); без пароля — сгенерировать
  list                          — сводка по пользователям (email, роль, active, хеш есть/нет)

env: NOCODB_URL, NOCODB_TOKEN, NOCODB_BASE, PORTAL_USERS_TABLE (см. auth.py).
"""

import json
import secrets
import string
import sys

from auth import (Noco, USERS_TABLE, MIN_PASSWORD_LEN, fields_of, row_id_of,
                  make_salt, hash_password)

ALPHABET = "abcdefghjkmnpqrstuvwxyzABCDEFGHJKMNPQRSTUVWXYZ23456789"


def gen_password(length=10):
    return "".join(secrets.choice(ALPHABET) for _ in range(length))


def load_users(nc):
    users = []
    for row in nc.list_rows(USERS_TABLE):
        users.append((row_id_of(row), fields_of(row)))
    return users


def cmd_init(nc):
    users = load_users(nc)
    issued = []
    for rid, f in users:
        if (f.get("password_hash") or "").strip():
            continue
        email = (f.get("email") or "").strip()
        if not email:
            print("SKIP: строка без email (Id=%s)" % rid)
            continue
        pwd = gen_password()
        salt = make_salt()
        nc.update_row(USERS_TABLE, rid, {
            "salt": salt,
            "password_hash": hash_password(salt, pwd),
            "must_change_password": True,
        })
        issued.append((email, pwd))
    print("=== Выданные пароли (показать один раз, не сохранять в чате/файлах) ===")
    for email, pwd in issued:
        print(f"{email}\t{pwd}")
    print(f"=== Итого выдано: {len(issued)} из {len(users)} пользователей ===")


def cmd_set_password(nc, email, password=None):
    if not email:
        print("ERROR: укажите email")
        return 1
    pwd = password or gen_password()
    if len(pwd) < MIN_PASSWORD_LEN:
        print(f"ERROR: пароль короче {MIN_PASSWORD_LEN} символов")
        return 1
    target = None
    for rid, f in load_users(nc):
        if (f.get("email") or "").strip().lower() == email.strip().lower():
            target = (rid, f)
            break
    if target is None:
        print(f"ERROR: пользователь {email} не найден")
        return 1
    salt = make_salt()
    nc.update_row(USERS_TABLE, target[0], {
        "salt": salt,
        "password_hash": hash_password(salt, pwd),
        "must_change_password": True,
        "token": "",
        "token_expires": "",
    })
    print(f"OK: пароль установлен для {email} (must_change_password=1)")
    if password is None:
        print(f"{email}\t{pwd}")
    return 0


def cmd_list(nc):
    for rid, f in load_users(nc):
        print("%-35s %-12s active=%-5s hash=%s mcp=%s" % (
            f.get("email") or "-",
            f.get("role") or "-",
            "yes" if f.get("active", True) else "no",
            "yes" if (f.get("password_hash") or "").strip() else "NO",
            "yes" if f.get("must_change_password") else "no",
        ))


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return 1
    cmd = args[0]
    nc = Noco()
    nc.connect()
    if cmd == "init":
        cmd_init(nc)
        return 0
    if cmd == "set-password":
        return cmd_set_password(nc, args[1] if len(args) > 1 else "",
                                args[2] if len(args) > 2 else None)
    if cmd == "list":
        cmd_list(nc)
        return 0
    print(f"ERROR: неизвестная команда: {cmd}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
