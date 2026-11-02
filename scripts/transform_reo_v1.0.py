#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
transform_reo_v1.0.py — ETL: выгрузка reo_statement (JSON из 1С) -> NocoDB [DEV]_reo_statement

Использование (python_worker / локально):
  python transform_reo_v1.0.py <input.json> [--report-date ГГГГ-ММ-ДД] \
      [--trigger-source scheduled|manual] [--dry-run]
  python transform_reo_v1.0.py --fetch [--report-date ГГГГ-ММ-ДД] [...]

  report_date: по умолчанию — последний полностью проведённый день
  (сегодня минус 2 рабочих дня, Пн-Пт; праздники сознательно не учитываются).
  Дата свежее разрешённой -> отказ (обход --force).
  trigger_source: по умолчанию manual

Конфигурация NocoDB (переменные окружения контейнера python_worker, docker-compose):
  NOCODB_URL   — напр. http://nocodb:8080 (внутри automation_net)
  NOCODB_TOKEN — API-токен NocoDB (xc-token)
  NOCODB_BASE  — название базы (по умолчанию "Demo-Portal")

Логика (порт проверенного scripts/_fill_template.py, этап 0):
  - сопоставление: НГ содержит contract_no И keyword (регистронезависимо,
    без точек/пробелов; "дог19" и "дог.103" нормализуются в одно) -> object_name
  - периоды: годовые; 2025 делится на 2025H1/2025H2 (единая схема для всех
    объектов; генератор Excel сам соберёт годовые строки для секций без деления)
  - классификация: payment/advance -> advance_money; payment/cession ->
    advance_cession; payment/works -> works_paid; closed_advance ->
    closed_advance; realization -> realization;
    91.02->62.0x (корректировка дебиторки) -> works_paid (решение 2026-11-02)
  - обеспечение-договоры (cession, Дт 76.09 Кт 62.01) -> bg_retention,
    contract-строка (period="contract", object_name="-")
  - прочее (other, needs_review, unmapped) в суммы не входит;
    needs_review_count/needs_review_sum — статистика исключённого по (объект, период)
  - записи с датой > report_date игнорируются (страховка)
  - снапшот идемпотентен: строки с тем же report_date+trigger_source удаляются
    и пишутся заново; снапшоты других дат и других trigger_source не трогаются

Вывод: JSON-сводка в stdout (ASCII) — её парсит n8n для etl_log.
"""
import argparse
import base64
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone

# ---------------------------------------------------------------- константы
TYPE_COL = {
    "payment/advance": "advance_money",
    "payment/cession": "advance_cession",
    "payment/works": "works_paid",
    "closed_advance": "closed_advance",
    "realization": "realization",
}
NUMERIC_FIELDS = ["advance_money", "advance_cession", "works_paid",
                  "closed_advance", "realization", "bg_retention",
                  "needs_review_sum"]
TABLE_TITLES = {
    "ng": "[DEV]_reo_ng_reference",
    "terms": "[DEV]_reo_contract_terms",
    "statement": "[DEV]_reo_statement",
}
PAGE_LIMIT = 500

# ---------------------------------------------------------------- NocoDB API
class Noco:
    def __init__(self, base_url, token, base_title):
        self.base_url = base_url.rstrip("/")
        self.token = token
        self.base_title = base_title
        self.base_id = None
        self.tables = {}  # title -> table id

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
            if (b.get("title") or "").lower() == self.base_title.lower():
                found = b
                break
        if not found:
            titles = [b.get("title") for b in bases.get("list", [])]
            raise RuntimeError(f"Base '{self.base_title}' not found. Bases: {titles}")
        base_id = found["id"]
        self.base_id = base_id
        meta = self.request("GET", f"/api/v2/meta/bases/{base_id}/tables")
        for t in meta.get("list", []):
            self.tables[t["title"]] = t["id"]
        missing = [t for t in TABLE_TITLES.values() if t not in self.tables]
        if missing:
            raise RuntimeError(f"Tables not found in base: {missing}")

    def _try(self, variants):
        """Пробует варианты эндпоинтов по очереди (разные версии NocoDB).
        variants: [(method, path, params, body), ...]; 404/405 -> следующий вариант."""
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

    def list_rows(self, table_title, where=None):
        tid = self.tables[table_title]
        rows, offset = [], 0
        while True:
            params = {"limit": PAGE_LIMIT, "offset": offset}
            if where:
                params["where"] = where
            list_body = {"limit": PAGE_LIMIT, "offset": offset}
            if where:
                list_body["where"] = where
            resp = self._try([
                ("GET", f"/api/v2/tables/{tid}/rows", params, None),
                ("POST", f"/api/v2/tables/{tid}/rows/list", None, list_body),
                ("GET", f"/api/v1/db/data/v1/{self.base_id}/{tid}", params, None),
            ])
            if isinstance(resp, list):          # старый v1: голый массив
                chunk, page_info = resp, {}
            else:
                chunk = resp.get("list", [])
                page_info = resp.get("pageInfo", {})
            rows.extend(chunk)
            is_last = page_info.get("isLastPage", len(chunk) < PAGE_LIMIT)
            if is_last or not chunk:
                break
            offset = page_info.get("offset", offset + PAGE_LIMIT)
        return rows

    def delete_row(self, table_title, row_id):
        tid = self.tables[table_title]
        self._try([
            ("DELETE", f"/api/v2/tables/{tid}/rows/{row_id}", None, None),
            ("DELETE", f"/api/v2/tables/{tid}/rows", None, {"Id": [row_id]}),
            ("DELETE", f"/api/v1/db/data/v1/{self.base_id}/{tid}/{row_id}", None, None),
        ])

    def insert_rows(self, table_title, rows):
        if not rows:
            return 0
        tid = self.tables[table_title]
        written = 0
        for i in range(0, len(rows), 100):
            batch = rows[i:i + 100]
            try:
                self.request("POST", f"/api/v2/tables/{tid}/rows", body=batch)
            except RuntimeError as e:
                if "HTTP 404" in str(e) or "HTTP 405" in str(e):
                    for r in batch:   # v1: по одной строке
                        self.request("POST",
                                     f"/api/v1/db/data/v1/{self.base_id}/{tid}", body=r)
                else:
                    raise
            written += len(batch)
        return written

# ---------------------------------------------------------------- нормализация
def norm(s):
    """нижний регистр, без точек и пробелов: 'дог.№101' -> 'дог№9', 'дог.103' -> 'дог19'"""
    return (s or "").lower().replace(".", "").replace(" ", "").replace("\u00a0", "")

# ---------------------------------------------------------------- сопоставление
def build_rules(ng_rows):
    rules = []
    for r in ng_rows:
        f = r.get("fields", r)
        if not f.get("active", True):
            continue
        rules.append({
            "contract_no": (f.get("contract_no") or "").strip(),
            "keyword": (f.get("keyword") or "").strip().lower(),
            "object_name": (f.get("object_name") or "").strip(),
        })
    return rules

def object_of(ng, rules):
    """-> object_name | None | 'AMBIGUOUS'"""
    n = norm(ng)
    if not n:
        return None
    hits = set()
    for rule in rules:
        if norm(rule["contract_no"]) in n and rule["keyword"] in n:
            hits.add(rule["object_name"])
    if len(hits) == 1:
        return hits.pop()
    if len(hits) > 1:
        return "AMBIGUOUS"
    return None

def period_key(iso_date):
    year = iso_date[:4]
    if year == "2025":
        return "2025H1" if iso_date <= "2025-06-30" else "2025H2"
    return year

# ---------------------------------------------------------------- агрегация
def aggregate(records, report_date, rules):
    """-> (object_rows, contract_rows, stats)"""
    obj = defaultdict(lambda: defaultdict(float))   # (object, period) -> field -> sum
    nr_cnt = defaultdict(int)                       # (object, period) -> count
    nr_sum = defaultdict(float)
    bg = defaultdict(float)                         # contract_no -> sum
    obj_contract = {r["object_name"]: r["contract_no"] for r in rules}
    stats = {"total": len(records), "used": 0, "after_cutoff": 0,
             "bg_records": 0, "excluded": 0, "excluded_sum": 0.0,
             "technical": 0, "technical_sum": 0.0,
             "works_reclass": 0, "works_reclass_sum": 0.0,
             "unmapped": 0, "unmapped_sum": 0.0, "ambiguous": 0}
    unmapped_samples = []

    for rec in records:
        if rec["Дата"] > report_date:
            stats["after_cutoff"] += 1
            continue
        dog = rec.get("Договор") or ""
        # известные технические проводки — исключаются МОЛЧА (решено на этапе 1):
        #   Дт 46.02 Кт 90.01.* — закрытие этапов строительства (годовое/квартальное) + реверсы.
        # Они НЕ попадают в needs_review: паттерны известны и проверены сверкой.
        if rec.get("СчетДт") == "46.02":
            stats["technical"] += 1
            stats["technical_sum"] += rec["Сумма"]
            continue
        # Дт 91.02 Кт 62.0x — корректировка дебиторки за выполненные работы
        # (задолженность погашена списанием на прочие расходы, деньги не приходили).
        # Решение пользователя 2026-11-02 (п.4 файла правок): учитывать КАК оплату
        # выполненных работ (works_paid). Единственный известный случай — Сосновый-2,
        # 24.03.2025, корректировка актов от 10.12.2024, 767 203,16.
        if (rec.get("СчетДт") or "").startswith("91.02"):
            if (rec.get("СчетКт") or "").startswith("62."):
                rec = dict(rec, record_type="payment", payment_type="works",
                           needs_review=False)
                stats["works_reclass"] += 1
                stats["works_reclass_sum"] += rec["Сумма"]
            else:
                stats["technical"] += 1
                stats["technical_sum"] += rec["Сумма"]
                continue
        # обеспечение-договоры -> bg_retention (cession) / забаланс (other, мимо)
        if dog.lower().startswith("обеспечение"):
            m = re.search(r"дог\.№(\d+)", dog)
            if m and rec["record_type"] == "payment" and rec.get("payment_type") == "cession":
                bg["дог.№" + m.group(1)] += rec["Сумма"]
                stats["bg_records"] += 1
                continue
            # забалансовый учёт гарантии Дт 009.01 (other) — техническая запись:
            # на суммы/needs_review снапшота не влияет (решение 2026-11-01,
            # раньше падала в unmapped-счётчик — безвредно, но чище так)
            if (rec.get("СчетДт") or "").startswith("009."):
                stats["technical"] += 1
                stats["technical_sum"] += rec["Сумма"]
                continue
        obj_name = object_of(rec.get("НоменклатурнаяГруппа"), rules)
        if obj_name in (None, "AMBIGUOUS"):
            stats["unmapped"] += 1
            stats["unmapped_sum"] += rec["Сумма"]
            if obj_name == "AMBIGUOUS":
                stats["ambiguous"] += 1
            if len(unmapped_samples) < 5:
                unmapped_samples.append({"Дата": rec["Дата"],
                                         "НГ": (rec.get("НоменклатурнаяГруппа") or "")[:60],
                                         "Сумма": rec["Сумма"]})
            continue
        pk = period_key(rec["Дата"])
        excluded = bool(rec.get("needs_review")) or rec["record_type"] == "other"
        if excluded:
            stats["excluded"] += 1
            stats["excluded_sum"] += rec["Сумма"]
            nr_cnt[(obj_name, pk)] += 1
            nr_sum[(obj_name, pk)] += rec["Сумма"]
            continue
        field = TYPE_COL.get(rec["record_type"] + "/" + (rec.get("payment_type") or "")
                             if rec.get("payment_type") else rec["record_type"])
        if field is None:
            stats["unmapped"] += 1
            stats["unmapped_sum"] += rec["Сумма"]
            continue
        obj[(obj_name, pk)][field] += rec["Сумма"]
        stats["used"] += 1

    rows = []
    for (o, pk) in sorted(obj):
        row = {"object_name": o, "period": pk,
               "contract_no": obj_contract.get(o, "")}
        for f in TYPE_COL.values():
            row[f] = round(obj[(o, pk)].get(f, 0.0), 2)
        row["needs_review_count"] = nr_cnt.get((o, pk), 0)
        row["needs_review_sum"] = round(nr_sum.get((o, pk), 0.0), 2)
        rows.append(row)
    for c in sorted(bg):
        rows.append({"object_name": "-", "period": "contract",
                     "contract_no": c, "bg_retention": round(bg[c], 2),
                     "needs_review_count": 0, "needs_review_sum": 0.0})
    return rows, stats, unmapped_samples

# ---------------------------------------------------------------- запись снапшота
def write_snapshot(nc, report_date, trigger_source, rows, dry_run):
    # where-синтаксис v1 API капризный — читаем всё и фильтруем в Python
    # (таблица небольшая: ~40 строк на снапшот)
    all_rows = nc.list_rows(TABLE_TITLES["statement"])
    old = []
    for r in all_rows:
        f = r.get("fields", r)
        if ((f.get("report_date") or "")[:10] == report_date
                and (f.get("trigger_source") or "") == trigger_source):
            old.append(r)
    deleted = 0
    if not dry_run:
        for r in old:
            nc.delete_row(TABLE_TITLES["statement"], r["Id"])
            deleted += 1
        payload = []
        for r in rows:
            rec = {"report_date": report_date, "trigger_source": trigger_source}
            rec.update(r)
            key = "|".join([rec.get("contract_no") or "-", rec["object_name"],
                            rec["period"], report_date])
            rec["unique_key"] = key
            payload.append(rec)
        nc.insert_rows(TABLE_TITLES["statement"], payload)
    return len(old), deleted

# ---------------------------------------------------------------- даты и лаг выписки
def max_allowed_report_date(today=None):
    """Самая свежая дата, за которую 1С уже полностью проведена.

    Лаг: платёж в день D -> выписка банка на след. день -> проведение вечером
    этого дня. Полностью проведённый день = 2 РАБОЧИХ дня назад
    (Пн->Чт, Вт->Пт, Ср->Пн, Чт->Вт, Пт->Ср, Сб/Вс->Чт).
    Рабочие дни = Пн-Пт (праздники/переносы НЕ учитываются сознательно:
    отчёты кумулятивные, движения праздничных недель просто попадают
    в следующий снапшот — решение 2026-11-01).
    Фикс 2026-11-01: раньше считалось «-2 календарных дня с откатом от
    выходных» — в субботу/воскресенье/понедельник это разрешало ПЯТНИЦУ,
    хотя её выписка проводится только вечером понедельника (неполные данные).
    """
    d = today or date.today()
    workdays = 0
    while workdays < 2:
        d -= timedelta(days=1)
        if d.weekday() < 5:          # Sat=5, Sun=6
            workdays += 1
    return d

# ---------------------------------------------------------------- fetch из 1С
def fetch_from_1c(report_date):
    """Скачивает выгрузку reo_statement из 1С (Basic-auth), сохраняет в /data/input/.
    Креды и параметры — env контейнера:
      REO_1C_URL       — база публикации, напр. http://10.0.0.2/demo_demo-api
      REO_1C_USER      — пользователь 1С
      REO_1C_PASSWORD  — пароль
      REO_DATE_FROM    — начало периода (по умолчанию 2023-12-01)
      REO_CONTRACTS    — список договоров (по умолчанию дог.№101;дог.№102;дог.№103;дог.№104)
      REO_COUNTERPARTY — контрагент (по умолчанию ИнвестСтрой)
    """
    base = os.environ.get("REO_1C_URL", "").rstrip("/")
    user = os.environ.get("REO_1C_USER", "")
    pwd = os.environ.get("REO_1C_PASSWORD", "")
    if not base or not user or not pwd:
        raise SystemExit("ERROR: env REO_1C_URL / REO_1C_USER / REO_1C_PASSWORD not set")
    params = urllib.parse.urlencode({
        "date_from": os.environ.get("REO_DATE_FROM", "2023-12-01"),
        "date_to": report_date,
        "contracts": os.environ.get("REO_CONTRACTS", "дог.№101;дог.№102;дог.№103;дог.№104"),
        "counterparty": os.environ.get("REO_COUNTERPARTY", "ИнвестСтрой"),
    })
    url = f"{base}/hs/demo-api/reo_statement?{params}"
    req = urllib.request.Request(url)
    req.add_header("Authorization", "Basic " +
                   base64.b64encode(f"{user}:{pwd}".encode()).decode())
    try:
        with urllib.request.urlopen(req, timeout=300) as resp:
            raw = resp.read()
    except urllib.error.HTTPError as e:
        raise SystemExit(f"ERROR: 1C fetch -> HTTP {e.code}: "
                         f"{e.read().decode('utf-8', 'replace')[:300]}")
    records = json.loads(raw.decode("utf-8-sig"))
    if not isinstance(records, list):
        raise SystemExit("ERROR: 1C response is not a JSON array")
    out = f"/data/input/reo_statement_{report_date}.json"
    with open(out, "wb") as f:
        f.write(raw)
    return out, len(records)

# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input", nargs="?", default=None,
                    help="путь к JSON-выгрузке (не нужен при --fetch)")
    ap.add_argument("--fetch", action="store_true",
                    help="скачать выгрузку из 1С (env REO_1C_*)")
    ap.add_argument("--report-date", default=None,
                    help="по умолчанию — последний проведённый день (сегодня-2 раб. дня)")
    ap.add_argument("--trigger-source", default="manual")
    ap.add_argument("--force", action="store_true",
                    help="разрешить дату свежее лага выписки (неполные данные!)")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    max_allowed = max_allowed_report_date()
    if args.report_date:
        report_date = args.report_date
        if not args.force and report_date > max_allowed.isoformat():
            raise SystemExit(
                f"ERROR: report_date {report_date} свежее разрешённой "
                f"{max_allowed.isoformat()} (лаг выписки ~2 раб. дня: платёж D -> "
                f"выписка D+1 -> проведение вечером D+1). Данные будут неполными. "
                f"Повторите без даты, с датой <= {max_allowed.isoformat()} или с --force.")
    else:
        report_date = max_allowed.isoformat()
    datetime.strptime(report_date, "%Y-%m-%d")

    if args.fetch:
        input_path, fetched = fetch_from_1c(report_date)
        print(f"FETCHED {fetched} records -> {input_path}")
    else:
        if not args.input:
            raise SystemExit("ERROR: input file or --fetch required")
        input_path = args.input

    nc = Noco(os.environ.get("NOCODB_URL", ""),
              os.environ.get("NOCODB_TOKEN", ""),
              os.environ.get("NOCODB_BASE", "Demo-Portal"))
    if not nc.base_url or not nc.token:
        raise SystemExit("ERROR: env NOCODB_URL / NOCODB_TOKEN not set")

    with open(input_path, encoding="utf-8-sig") as f:
        records = json.load(f)

    nc.connect()
    rules = build_rules(nc.list_rows(TABLE_TITLES["ng"]))
    if not rules:
        raise SystemExit("ERROR: [DEV]_reo_ng_reference is empty")
    terms = nc.list_rows(TABLE_TITLES["terms"])
    term_objects = {(r.get("fields", r).get("object_name") or "").strip() for r in terms}

    rows, stats, unmapped_samples = aggregate(records, report_date, rules)

    # объекты без договорных условий — предупреждение
    no_terms = sorted({r["object_name"] for r in rows
                       if r["period"] != "contract" and r["object_name"] not in term_objects})

    old_count, deleted = write_snapshot(nc, report_date, args.trigger_source,
                                        rows, args.dry_run)

    summary = {
        "status": "dry-run" if args.dry_run else "ok",
        "report_date": report_date,
        "max_allowed_date": max_allowed.isoformat(),
        "trigger_source": args.trigger_source,
        "records_total": stats["total"],
        "records_used": stats["used"],
        "records_after_cutoff": stats["after_cutoff"],
        "records_bg": stats["bg_records"],
        "records_excluded": stats["excluded"],
        "excluded_sum": round(stats["excluded_sum"], 2),
        "records_technical": stats["technical"],
        "technical_sum": round(stats["technical_sum"], 2),
        "records_works_reclass": stats["works_reclass"],
        "works_reclass_sum": round(stats["works_reclass_sum"], 2),
        "records_unmapped": stats["unmapped"],
        "unmapped_sum": round(stats["unmapped_sum"], 2),
        "ambiguous_ng": stats["ambiguous"],
        "rows_written": 0 if args.dry_run else len(rows),
        "rows_deleted_before": old_count,
        "objects_without_terms": no_terms,
        "unmapped_samples": unmapped_samples,
    }
    print(json.dumps(summary, ensure_ascii=True, indent=1))

if __name__ == "__main__":
    main()
