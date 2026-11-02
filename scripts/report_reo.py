# -*- coding: utf-8 -*-
"""
report_reo.py — фасад данных REO для дашборда (этап 4б).

Запускается в python_worker (n8n [DEV]_IN_REO_Report вызывает через SSH-ноду):
  echo '<base64-json>' | base64 -d | docker exec -i python_worker python /data/scripts/report_reo.py

Вход:  JSON {"token": "...", "format": "json|dashboard|excel|all", "date": "ГГГГ-ММ-ДД"}
Выход: {"status": "success", "data": {meta, snapshot, dashboard, excel}}

Логика (согласована 2026-11-01):
  - дата = дате данных: снапшот = последний report_date <= запрошенной даты
    (приоритет scheduled); отчёт/дашборд датируются ФАКТИЧЕСКОЙ датой снапшота;
  - графики динамики — только trigger_source=scheduled (ключевое решение ROADMAP 9.2);
  - excel — generate_excel_reo_v1.0.py --fallback --reuse-existing (быстрый путь);
  - доступ: роли admin/accounting/management (проверка токена на сервере).

env: те же NOCODB_* / PORTAL_USERS_TABLE, что в auth.py.
"""

import base64
import importlib.util
import json
import os
import subprocess
import sys
from datetime import date

import auth
from auth import Noco, ok, err, fields_of, row_id_of

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
GENERATE_SCRIPT = os.path.join(SCRIPT_DIR, "generate_excel_reo_v1.0.py")
TRANSFORM_SCRIPT = os.path.join(SCRIPT_DIR, "transform_reo_v1.0.py")

TABLE_TITLES = {
    "terms": "[DEV]_reo_contract_terms",
    "statement": "[DEV]_reo_statement",
    "ng": "[DEV]_reo_ng_reference",
}

REO_ROLES = ("admin", "accounting", "management")

# Площадка (верхний уровень drill-down). Явный маппинг по договору —
# в духе правила «явные маппинги»; когда в reo_ng_reference появится
# колонка site, она будет иметь приоритет (см. site_of()).
SITE_BY_CONTRACT = {
    "дог.№101": "Сосновый",
    "дог.№102": "Центральный",
    "дог.№103": "Северный",
    "дог.№104": "Северный",
}

# Удержание по БГ (bg_retention) проводится по договору обеспечения, но
# удерживается из оплат КОНКРЕТНОГО объекта (решение 2026-11-02: «получили
# деньги по объекту, а заплатили по всему договору» — делить не нужно,
# вся сумма относится к объекту-источнику оплат). Явный маппинг
# «договор -> объект», одобрен пользователем 2026-11-02.
BG_OBJECT_BY_CONTRACT = {
    "дог.№101": "ЖК Сосновый, корпус 2 (дог 101)",
    "дог.№102": "Административный центр (дог 102)",
    "дог.№103": "Спорткомплекс (дог 103)",
    "дог.№104": "Инженерная сеть (дог 104)",
}

METRIC_FIELDS = ("advance_money", "advance_cession", "works_paid",
                 "closed_advance", "realization", "bg_retention")


def norm(s):
    return (s or "").lower().replace(".", "").replace(" ", "").replace("\u00a0", "")


BG_OBJECT_BY_NORM = {norm(k): v for k, v in BG_OBJECT_BY_CONTRACT.items()}


def num(v):
    return v if isinstance(v, (int, float)) else 0.0


def max_allowed_report_date():
    """Лаг выписки (та же логика, что в transform_reo_v1.0.py): сегодня - 2 раб. дня."""
    spec = importlib.util.spec_from_file_location("reo_transform", TRANSFORM_SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod.max_allowed_report_date().isoformat()


# ---------------------------------------------------------------- данные
def load_history(nc):
    """report_date -> {trigger_source: [строки]} по всей таблице снапшотов."""
    history = {}
    for r in nc.list_rows(TABLE_TITLES["statement"]):
        f = fields_of(r)
        rd = (f.get("report_date") or "")[:10]
        if rd:
            history.setdefault(rd, {}).setdefault(
                f.get("trigger_source") or "manual", []).append(f)
    return history


def pick_snapshot(history, requested_date):
    """-> (actual_date, rows, source, fallback). Снапшот <= даты.
    Приоритет источников: scheduled > backfill > manual (любой годится для
    просмотра; в график динамики попадают только scheduled+backfill)."""
    def for_date(rd):
        for src in ("scheduled", "backfill", "manual"):
            rows = history.get(rd, {}).get(src, [])
            if rows:
                return rows, src
        return [], None

    rows, src = for_date(requested_date)
    if rows:
        return requested_date, rows, src, False
    older = sorted(rd for rd in history if rd <= requested_date)
    if older:
        rd = older[-1]
        rows, src = for_date(rd)
        return rd, rows, src, True
    return requested_date, [], None, False


def load_terms(nc):
    terms = {}
    for r in nc.list_rows(TABLE_TITLES["terms"]):
        f = fields_of(r)
        if f.get("active", True):
            terms[(f.get("object_name") or "").strip()] = f
    return terms


def site_of(obj_name, contract_no, ng_rows):
    # приоритет: явная колонка site в справочнике (на будущее)
    for f in ng_rows:
        if (f.get("object_name") or "").strip() == obj_name and (f.get("site") or "").strip():
            return f["site"].strip()
    n = norm(contract_no)
    for contract, site in SITE_BY_CONTRACT.items():
        if norm(contract) == n:
            return site
    return contract_no or "Прочее"


# ---------------------------------------------------------------- агрегаты
def bg_by_contract_of(rows):
    """norm(contract_no) -> bg_retention из contract-строк снапшота."""
    d = {}
    for f in rows:
        if f.get("period") == "contract":
            d[norm(f.get("contract_no") or "")] = num(f.get("bg_retention"))
    return d


def object_metrics(snap_rows, obj_name, contract_no="", bg_contracts=None):
    """Суммы по всем периодам объекта в снапшоте (кумулятивно на дату снапшота).
    bg_retention — удержание по БГ договора объекта, если оно удерживалось
    из оплат этого объекта (BG_OBJECT_BY_CONTRACT)."""
    m = {f: 0.0 for f in METRIC_FIELDS}
    for f in snap_rows:
        if f.get("period") != "contract" and (f.get("object_name") or "").strip() == obj_name:
            for k in METRIC_FIELDS:
                m[k] += num(f.get(k))
    n = norm(contract_no or "")
    if BG_OBJECT_BY_NORM.get(n) == obj_name:
        m["bg_retention"] += num((bg_contracts or {}).get(n, 0.0))
    # «Получено оплат» = деньги (авансы + оплата работ) + цессии + удержание по БГ
    m["received"] = (m["advance_money"] + m["advance_cession"]
                     + m["works_paid"] + m["bg_retention"])
    # остаток аванса НЕ включает bg_retention: зачет на БГ идёт через Кт 62.01
    # (закрывает дебиторку за работы, а не аванс) — решение 2026-11-02
    m["advance_balance"] = m["advance_money"] + m["advance_cession"] - m["closed_advance"]
    return m


def round_metrics(m):
    return {k: round(v, 2) for k, v in m.items()}


def build_tree(snap_rows, terms, ng_rows):
    """Портфель -> Площадка -> Договор -> Объект (для Махачкалы/Центральныйа
    площадка = договор, уровень договора не рисуем)."""
    objects = sorted({(f.get("object_name") or "").strip() for f in snap_rows
                      if f.get("period") != "contract" and f.get("object_name")})
    obj_contract = {}
    for o in objects:
        t = terms.get(o) or {}
        obj_contract[o] = (t.get("contract_no") or "").strip()

    sites = {}   # site -> contract -> [objects]
    for o in objects:
        site = site_of(o, obj_contract[o], ng_rows)
        sites.setdefault(site, {}).setdefault(obj_contract[o], []).append(o)

    portfolio = {k: 0.0 for k in METRIC_FIELDS}
    bg_contracts = bg_by_contract_of(snap_rows)
    tree = []
    for site in sorted(sites):
        site_m = {k: 0.0 for k in METRIC_FIELDS}
        site_node = {"name": site, "metrics": None, "children": []}
        contracts = sites[site]
        single = len(contracts) == 1
        for contract_no in sorted(contracts):
            contract_m = {k: 0.0 for k in METRIC_FIELDS}
            contract_node = {"name": contract_no, "metrics": None, "children": []}
            for o in contracts[contract_no]:
                m = object_metrics(snap_rows, o, obj_contract[o], bg_contracts)
                t = terms.get(o) or {}
                cost = num(t.get("contract_cost"))
                exec_pct = (m["realization"] / cost * 100) if cost else None
                obj_node = {
                    "name": o,
                    "contract_no": contract_no,
                    "cost": round(cost, 2),
                    "execution_pct": round(exec_pct, 1) if exec_pct is not None else None,
                    # флаг «объект завершён» из справочника (вариант 3, 2026-11-02):
                    # процент по букве договора может быть < 100, показываем «✓ завершён»
                    "completed": bool(t.get("completed")),
                    "metrics": round_metrics(m),
                    "children": [],
                }
                contract_node["children"].append(obj_node)
                for k in METRIC_FIELDS:
                    contract_m[k] += m[k]
            contract_m["received"] = (contract_m["advance_money"]
                                      + contract_m["advance_cession"]
                                      + contract_m["works_paid"]
                                      + contract_m["bg_retention"])
            contract_m["advance_balance"] = (contract_m["advance_money"]
                                             + contract_m["advance_cession"]
                                             - contract_m["closed_advance"])
            contract_node["metrics"] = round_metrics(contract_m)
            if single:
                # площадка = договор: объекты сразу в площадку
                site_node["children"].extend(contract_node["children"])
            else:
                site_node["children"].append(contract_node)
            for k in METRIC_FIELDS:
                site_m[k] += contract_m[k]
        site_m["received"] = (site_m["advance_money"] + site_m["advance_cession"]
                              + site_m["works_paid"] + site_m["bg_retention"])
        site_m["advance_balance"] = (site_m["advance_money"] + site_m["advance_cession"]
                                     - site_m["closed_advance"])
        site_node["metrics"] = round_metrics(site_m)
        tree.append(site_node)
        for k in METRIC_FIELDS:
            portfolio[k] += site_m[k]
    portfolio["received"] = (portfolio["advance_money"] + portfolio["advance_cession"]
                             + portfolio["works_paid"] + portfolio["bg_retention"])
    portfolio["advance_balance"] = (portfolio["advance_money"] + portfolio["advance_cession"]
                                    - portfolio["closed_advance"])
    return tree, round_metrics(portfolio)


def build_dashboard(snap_rows, history, actual_date, terms, ng_rows):
    tree, portfolio = build_tree(snap_rows, terms, ng_rows)

    # динамика по истории scheduled + backfill до actual_date включительно
    # (решение 2026-11-01, вариант A: восполненная история — отдельный источник,
    #  случайные ручные прогоны по-прежнему в график не попадают)
    def series_rows(rd):
        srcs = history.get(rd, {})
        return srcs.get("scheduled") or srcs.get("backfill") or []

    dates = sorted(rd for rd, srcs in history.items()
                   if ("scheduled" in srcs or "backfill" in srcs) and rd <= actual_date)
    objects = sorted({(f.get("object_name") or "").strip() for f in snap_rows
                      if f.get("period") != "contract" and f.get("object_name")})
    obj_contract = {o: ((terms.get(o) or {}).get("contract_no") or "").strip()
                    for o in objects}
    by_object = {o: [] for o in objects}
    by_type = {f: [] for f in METRIC_FIELDS}
    for rd in dates:
        rows = series_rows(rd)
        bg_contracts = bg_by_contract_of(rows)
        totals = {k: 0.0 for k in METRIC_FIELDS}
        for o in objects:
            m = object_metrics(rows, o, obj_contract[o], bg_contracts)
            by_object[o].append(round(m["realization"], 2))
            for k in METRIC_FIELDS:
                totals[k] += m[k]
        for k in METRIC_FIELDS:
            by_type[k].append(round(totals[k], 2))

    kpi = {
        "data_date": actual_date,
        "portfolio_cost": round(sum(num((terms.get(o) or {}).get("contract_cost"))
                                    for o in objects), 2),
        "received": portfolio["received"],
        "realization": portfolio["realization"],
        "advance_balance": portfolio["advance_balance"],
        # полные метрики портфеля — для расшифровки карточки «Получено оплат»
        "metrics": portfolio,
    }
    return {"kpi": kpi, "tree": tree, "series": {
        "dates": dates, "by_object": by_object, "by_type": by_type}}


# ---------------------------------------------------------------- excel
def fetch_from_1c_action(fetch_date, max_allowed):
    """Подгрузка данных из 1С на произвольную дату (схема A+C: пользователь
    отказался от ближайшего снапшота). transform пишет снапшот с
    trigger_source=manual — в графики динамики (scheduled) он не попадает."""
    if not fetch_date:
        raise RuntimeError("Не указана дата для подгрузки из 1С")
    if fetch_date > max_allowed:
        raise RuntimeError(
            f"Дата {fetch_date} свежее разрешённой {max_allowed} "
            f"(лаг выписки ~2 рабочих дня) — данные в 1С ещё не проведены")
    proc = subprocess.run(
        [sys.executable, TRANSFORM_SCRIPT, "--fetch",
         "--report-date", fetch_date, "--trigger-source", "manual"],
        capture_output=True, text=True, timeout=300)
    tail = ((proc.stdout or "") + (proc.stderr or "")).strip()[-400:]
    if proc.returncode != 0:
        # человекочитаемые причины (фронт показывает этот текст напрямую)
        if "HTTP 500" in tail or "Internal Server Error" in tail:
            raise RuntimeError(
                "Сервис 1С недоступен (ошибка на сервере 1С, HTTP 500). "
                "Попробуйте позже или обратитесь к администратору 1С — "
                "данные на портале (история) при этом доступны")
        if "свежее разрешённой" in tail:
            raise RuntimeError(
                "Данные за эту дату ещё не проведены в 1С (лаг выписки "
                "~2 рабочих дня). Выберите более раннюю дату")
        raise RuntimeError(f"Подгрузка из 1С не удалась: {tail}")
    return tail


def make_excel(requested_date):
    proc = subprocess.run(
        [sys.executable, GENERATE_SCRIPT,
         "--report-date", requested_date, "--fallback", "--reuse-existing"],
        capture_output=True, text=True, timeout=120)
    out = (proc.stdout or "").strip()
    try:
        info = json.loads(out)
    except ValueError:
        raise RuntimeError(f"generate_excel вернул не-JSON: {(proc.stderr or out)[:300]}")
    if info.get("status") != "ok":
        raise RuntimeError(f"generate_excel: {info}")
    path = info["output"]
    with open(path, "rb") as fh:
        b64 = base64.b64encode(fh.read()).decode("ascii")
    return {
        "file_name": os.path.basename(path),
        "file_data": b64,
        "report_date": info.get("report_date"),
        "requested_date": info.get("requested_date"),
        "cached": bool(info.get("cached")),
    }


# ---------------------------------------------------------------- справочник стоимостей
def terms_get_action(nc):
    """Список строк [DEV]_reo_contract_terms для редактора (все REO-роли)."""
    out = []
    for r in nc.list_rows(TABLE_TITLES["terms"]):
        f = fields_of(r)
        out.append({
            "id": row_id_of(r),
            "object_name": (f.get("object_name") or "").strip(),
            "contract_no": (f.get("contract_no") or "").strip(),
            "contract_cost": num(f.get("contract_cost")),
            "advance_limit": num(f.get("advance_limit")),
            "active": bool(f.get("active", True)),
            # объект завершён, стоимость договора не выбрана полностью
            # (остаток ждёт дс/перенос) — решение 2026-11-02, вариант 3
            "completed": bool(f.get("completed")),
        })
    out.sort(key=lambda x: x["object_name"])
    return out


def terms_set_action(nc, uf, updates):
    """Обновление стоимостей/лимитов. Право: только accounting/admin
    (проверка на сервере; скрытие кнопки во фронте — не защита)."""
    if (uf.get("role") or "") not in ("admin", "accounting"):
        raise PermissionError("Редактирование стоимостей доступно только бухгалтерии и администратору")
    if not isinstance(updates, list) or not updates:
        raise ValueError("Пустой список изменений")
    # индекс строк по object_name
    by_name = {}
    for r in nc.list_rows(TABLE_TITLES["terms"]):
        f = fields_of(r)
        by_name[(f.get("object_name") or "").strip()] = row_id_of(r)
    changed = 0
    for u in updates:
        name = (u.get("object_name") or "").strip()
        rid = by_name.get(name)
        if rid is None:
            raise ValueError(f"Объект '{name}' не найден в [DEV]_reo_contract_terms")
        fields = {}
        if "contract_cost" in u:
            v = u["contract_cost"]
            if v is not None and (not isinstance(v, (int, float)) or v < 0):
                raise ValueError(f"Некорректная стоимость у '{name}'")
            fields["contract_cost"] = round(float(v or 0), 2)
        if "advance_limit" in u:
            v = u["advance_limit"]
            if v is not None and (not isinstance(v, (int, float)) or v < 0):
                raise ValueError(f"Некорректный лимит у '{name}'")
            fields["advance_limit"] = round(float(v or 0), 2)
        if "completed" in u:
            fields["completed"] = bool(u["completed"])
        if fields:
            nc.update_row(TABLE_TITLES["terms"], rid, fields)
            changed += 1
    return {"changed": changed, "terms": terms_get_action(nc)}


# ---------------------------------------------------------------- main
def main():
    try:
        raw = sys.argv[1] if len(sys.argv) > 1 else sys.stdin.read()
        payload = json.loads(raw)
        token = payload.get("token") or ""
        fmt = (payload.get("format") or "all").lower()
        requested = (payload.get("date") or "").strip()

        nc = Noco()
        nc.connect()

        # --- авторизация: токен + роль (проверка НА СЕРВЕРЕ) ---
        rid, uf = auth._by_token(nc, token)
        if rid is None:
            print(json.dumps(err(401, "Сессия недействительна — войдите заново"),
                             ensure_ascii=False))
            return
        if (uf.get("role") or "") not in REO_ROLES:
            print(json.dumps(err(403, "Нет доступа к отчёту Демо-аналитика"),
                             ensure_ascii=False))
            return

        history = load_history(nc)
        scheduled_dates = sorted(rd for rd, srcs in history.items() if "scheduled" in srcs)
        all_dates = sorted(history)
        max_allowed = max_allowed_report_date()

        # --- действия справочника стоимостей (4г) ---
        if payload.get("action") == "terms_get":
            print(json.dumps(ok({"terms": terms_get_action(nc)}), ensure_ascii=False))
            return
        if payload.get("action") == "terms_set":
            try:
                result = terms_set_action(nc, uf, payload.get("updates") or [])
            except PermissionError as e:
                print(json.dumps(err(403, str(e)), ensure_ascii=False))
                return
            except (ValueError, RuntimeError) as e:
                print(json.dumps(err(400, str(e)), ensure_ascii=False))
                return
            print(json.dumps(ok(result), ensure_ascii=False))
            return

        # --- действие «подгрузить из 1С на дату» (схема A+C, ответ «Нет») ---
        if payload.get("action") == "fetch":
            fetch_date = (payload.get("date") or "").strip()
            fetch_from_1c_action(fetch_date, max_allowed)
            history = load_history(nc)
            scheduled_dates = sorted(rd for rd, srcs in history.items()
                                     if "scheduled" in srcs)
            all_dates = sorted(history)
            print(json.dumps(ok({
                "fetched_date": fetch_date,
                "available_dates": all_dates,
                "scheduled_dates": scheduled_dates,
                "max_allowed_date": max_allowed,
            }), ensure_ascii=False))
            return

        if not requested:
            requested = scheduled_dates[-1] if scheduled_dates else date.today().isoformat()
        actual_date, snap_rows, used_source, fallback = pick_snapshot(history, requested)
        if not snap_rows:
            print(json.dumps(err(404, f"На дату {requested} данных нет (история пуста)"),
                             ensure_ascii=False))
            return

        meta = {
            "requested_date": requested,
            "actual_date": actual_date,
            "fallback": fallback,
            "trigger_source": used_source,
            "snapshot_rows": len(snap_rows),
            "available_dates": all_dates,
            "scheduled_dates": scheduled_dates,
            "max_allowed_date": max_allowed,
            "role": uf.get("role") or "",
        }

        data = {"meta": meta}
        want_json = fmt in ("json", "all")
        want_dash = fmt in ("dashboard", "all")
        want_excel = fmt in ("excel", "all")

        if want_json or want_dash:
            terms = load_terms(nc)
            ng_rows = [fields_of(r) for r in nc.list_rows(TABLE_TITLES["ng"])]
            if want_json:
                data["snapshot"] = {
                    "rows": snap_rows,
                    "terms": [{k: t.get(k) for k in
                               ("object_name", "contract_no", "contract_cost", "advance_limit")}
                              for t in terms.values()],
                }
            if want_dash:
                data["dashboard"] = build_dashboard(snap_rows, history, actual_date,
                                                    terms, ng_rows)
        if want_excel:
            data["excel"] = make_excel(requested)

        print(json.dumps(ok(data), ensure_ascii=False))
    except Exception as e:  # noqa: BLE001 — наружу всегда JSON
        print(json.dumps(err(500, f"{type(e).__name__}: {e}"), ensure_ascii=False))


if __name__ == "__main__":
    main()
