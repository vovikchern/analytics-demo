#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_excel_reo_v1.0.py — отчёт Demo_Report: NocoDB снапшот -> xlsx по шаблону.

Использование (python_worker):
  python /data/scripts/generate_excel_reo_v1.0.py --report-date ГГГГ-ММ-ДД \
      [--trigger-source scheduled|manual] [--output /data/output/имя.xlsx]

  trigger_source: по умолчанию сначала ищется снапшот scheduled, если нет — manual.
  output: по умолчанию /data/output/Demo_Report_<report_date>.xlsx

Конфигурация NocoDB — те же env, что у transform_reo_v1.0.py:
  NOCODB_URL, NOCODB_TOKEN, NOCODB_BASE

Шаблон: /data/templates/reo_template.xlsx (этап 0; формулы S/C/итого считаются Excel).
Логика раскладки — порт scripts/_fill_template.py (этап 0, сверка с ручной сошлась):
  - секции шаблона: у Махачкалы/Урады/Центральныйов 2025 разбит на H1/H2,
    у Северныйов годовые строки (2025 = 2025H1 + 2025H2 из БД);
  - C (bg_retention) — из contract-строк (period="contract"), пишется в верх
    объединения C5/C16/C27/C37;
  - B (стоимость) и I (лимит аванса) — из [DEV]_reo_contract_terms;
  - B1 = report_date («Отчёт актуален на»).
"""
import argparse
import json
import os
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime

import openpyxl

PAGE_LIMIT = 500
TABLE_TITLES = {
    "terms": "[DEV]_reo_contract_terms",
    "statement": "[DEV]_reo_statement",
}
TEMPLATE_PATH = "/data/templates/reo_template.xlsx"
OUTPUT_DIR = "/data/output/reo"   # подпапка только для REO-отчётов (раздача 8081)

# Секции шаблона: объект -> {период шаблона: строка}
SECTIONS = {
    "ЖК Сосновый (дог 101)":     {"2024": 5, "2025H1": 6, "2025H2": 7, "2026": 8},
    "ЖК Сосновый, корпус 2 (дог 101)":             {"2024": 10, "2025H1": 11, "2025H2": 12, "2026": 13},
    "Административный центр (дог 102)":      {"2024": 16, "2025H1": 17, "2025H2": 18, "2026": 19},
    "Пристрой к школе (дог 102)":  {"2024": 21, "2025H1": 22, "2025H2": 23, "2026": 24},
    "Спорткомплекс (дог 103)":     {"2024": 27, "2025": 28, "2026": 29},
    "Плавательный бассейн (дог 103)": {"2023": 31, "2024": 32, "2025": 33, "2026": 34},
    "Инженерная сеть (дог 104)":     {"2024": 37, "2025": 38, "2026": 39},
    "Котельная (дог 104)": {"2023": 41, "2024": 42, "2025": 43, "2026": 44},
}
# подписи периодов в колонке D
LABELS = {"2025H1": "2025 (01.01-30.06)", "2025H2": "2025 (01.07-31.12)"}
SECTION_ITOGO = {9: (5, 8), 14: (10, 13), 20: (16, 19), 25: (21, 24),
                 30: (27, 29), 35: (31, 34), 40: (37, 39), 45: (41, 44)}
C_TOP = {"ЖК Сосновый (дог 101)": 5, "Административный центр (дог 102)": 16,
         "Спорткомплекс (дог 103)": 27, "Инженерная сеть (дог 104)": 37}
B_ROWS = {"ЖК Сосновый (дог 101)": 5, "ЖК Сосновый, корпус 2 (дог 101)": 10,
          "Административный центр (дог 102)": 16, "Пристрой к школе (дог 102)": 21,
          "Спорткомплекс (дог 103)": 27, "Плавательный бассейн (дог 103)": 31,
          "Инженерная сеть (дог 104)": 37, "Котельная (дог 104)": 41}
I_ROW_OF = {"ЖК Сосновый (дог 101)": 9, "ЖК Сосновый, корпус 2 (дог 101)": 14,
            "Административный центр (дог 102)": 20, "Пристрой к школе (дог 102)": 25,
            "Спорткомплекс (дог 103)": 30, "Плавательный бассейн (дог 103)": 35,
            "Инженерная сеть (дог 104)": 40, "Котельная (дог 104)": 45}
# договорные ИТОГО-строки I (15/26/36) — лимит договора; I46 — формула, не трогаем
CONTRACT_I_ROWS = {15: "ЖК Сосновый (дог 101)", 26: "Административный центр (дог 102)",
                   36: "Спорткомплекс (дог 103)"}
VALUE_COLS = {"advance_money": 5, "advance_cession": 6,   # E, F
              "works_paid": 10, "closed_advance": 12,     # J, L
              "realization": 14}                          # N

# ---------------------------------------------------------------- NocoDB API
class Noco:
    def __init__(self, base_url, token, base_title):
        self.base_url = base_url.rstrip("/")
        self.token = token
        self.base_title = base_title
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
            if (b.get("title") or "").lower() == self.base_title.lower():
                found = b
                break
        if not found:
            titles = [b.get("title") for b in bases.get("list", [])]
            raise RuntimeError(f"Base '{self.base_title}' not found. Bases: {titles}")
        self.base_id = found["id"]
        meta = self.request("GET", f"/api/v2/meta/bases/{self.base_id}/tables")
        for t in meta.get("list", []):
            self.tables[t["title"]] = t["id"]
        missing = [t for t in TABLE_TITLES.values() if t not in self.tables]
        if missing:
            raise RuntimeError(f"Tables not found in base: {missing}")

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
            params = {"limit": PAGE_LIMIT, "offset": offset}
            resp = self._try([
                ("GET", f"/api/v2/tables/{tid}/rows", params, None),
                ("POST", f"/api/v2/tables/{tid}/rows/list", None,
                 {"limit": PAGE_LIMIT, "offset": offset}),
                ("GET", f"/api/v1/db/data/v1/{self.base_id}/{tid}", params, None),
            ])
            if isinstance(resp, list):
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

# ---------------------------------------------------------------- раскладка
def flat(row):
    return row.get("fields", row)

def num(v):
    return v if isinstance(v, (int, float)) else 0.0

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--report-date", default=None,
                    help="по умолчанию — сегодня (дата запуска)")
    ap.add_argument("--trigger-source", default=None,
                    help="scheduled|manual; по умолчанию scheduled, при отсутствии — manual")
    ap.add_argument("--output", default=None)
    ap.add_argument("--fallback", action="store_true",
                    help="если снапшота на дату нет — взять последний снапшот <= даты "
                         "(режим вебхука; отчёт датируется ФАКТИЧЕСКОЙ датой снапшота)")
    ap.add_argument("--reuse-existing", action="store_true",
                    help="быстрый путь вебхука: если файл уже существует — не генерировать заново")
    args = ap.parse_args()
    args.report_date = args.report_date or date.today().isoformat()
    datetime.strptime(args.report_date, "%Y-%m-%d")

    nc = Noco(os.environ.get("NOCODB_URL", ""),
              os.environ.get("NOCODB_TOKEN", ""),
              os.environ.get("NOCODB_BASE", "Analytics-Portal"))
    if not nc.base_url or not nc.token:
        raise SystemExit("ERROR: env NOCODB_URL / NOCODB_TOKEN not set")
    nc.connect()

    # --- индекс истории снапшотов: report_date -> {trigger_source: [строки]} ---
    history = {}
    for r in nc.list_rows(TABLE_TITLES["statement"]):
        f = flat(r)
        rd = (f.get("report_date") or "")[:10]
        if rd:
            history.setdefault(rd, {}).setdefault(
                f.get("trigger_source") or "manual", []).append(f)

    def snapshot_for(rd, order):
        for src in order:
            rows = history.get(rd, {}).get(src, [])
            if rows:
                return rows, src
        return [], None

    order = [args.trigger_source] if args.trigger_source else ["scheduled", "manual"]
    snap, used_source = snapshot_for(args.report_date, order)
    actual_date = args.report_date
    fallback_used = False

    if not snap and args.fallback:
        # последний снапшот <= запрошенной даты (данные кумулятивные)
        older = sorted(rd for rd in history if rd <= args.report_date)
        if older:
            actual_date = older[-1]
            snap, used_source = snapshot_for(actual_date, order)
            fallback_used = actual_date != args.report_date

    if not snap:
        raise SystemExit(f"ERROR: no snapshot for {args.report_date} "
                         f"(tried: {', '.join(order)}"
                         + ("; fallback on" if args.fallback else "") + ")")

    # --- справочник условий ---
    terms = {}
    for r in nc.list_rows(TABLE_TITLES["terms"]):
        f = flat(r)
        if f.get("active", True):
            terms[(f.get("object_name") or "").strip()] = f
    if not terms:
        raise SystemExit("ERROR: [DEV]_reo_contract_terms is empty "
                         "(проверьте, что строки заполнены и чекбокс active отмечен)")

    # --- индекс снапшота ---
    by_obj_period = {}
    bg_by_contract = {}
    for f in snap:
        if f.get("period") == "contract":
            bg_by_contract[f.get("contract_no")] = num(f.get("bg_retention"))
        else:
            by_obj_period[(f.get("object_name"), f.get("period"))] = f

    def db_value(obj, tpl_period, field):
        """значение из БД; годовой '2025' = 2025H1 + 2025H2"""
        if tpl_period == "2025":
            v = 0.0
            for p in ("2025H1", "2025H2"):
                v += num(by_obj_period.get((obj, p), {}).get(field))
            return v
        return num(by_obj_period.get((obj, tpl_period), {}).get(field))

    # --- шаблон ---
    if not os.path.exists(TEMPLATE_PATH):
        raise SystemExit(f"ERROR: template not found: {TEMPLATE_PATH}")
    wb = openpyxl.load_workbook(TEMPLATE_PATH, data_only=False)
    ws = wb.worksheets[0]

    ws["B1"] = datetime.strptime(actual_date, "%Y-%m-%d").date()
    ws["B1"].number_format = "DD.MM.YYYY"

    missing_terms, empty_objects = [], []
    for obj, periods in SECTIONS.items():
        t = terms.get(obj)
        if t is None:
            missing_terms.append(obj)
        for tpl_period, row in periods.items():
            label = LABELS.get(tpl_period, tpl_period)
            ws.cell(row=row, column=4).value = int(label) if label.isdigit() else label
            for field, col in VALUE_COLS.items():
                ws.cell(row=row, column=col).value = round(db_value(obj, tpl_period, field), 2)
        ws.cell(row=I_ROW_OF[obj], column=4).value = "итого"
        # B (стоимость) и I (лимит) — из справочника, в СВОИ строки
        if t is not None:
            ws.cell(row=B_ROWS[obj], column=2).value = round(num(t.get("contract_cost")), 2)
            ws.cell(row=I_ROW_OF[obj], column=9).value = num(t.get("advance_limit"))
        # C (удержания по БГ) — из contract-строки снапшота
        contract_no = (t or {}).get("contract_no")
        if obj in C_TOP and contract_no:
            ws.cell(row=C_TOP[obj], column=3).value = round(
                bg_by_contract.get(contract_no, 0.0), 2)

    # I на договорных ИТОГО-строках (лимит один на договор)
    for r, obj in CONTRACT_I_ROWS.items():
        t = terms.get(obj)
        if t is not None:
            ws.cell(row=r, column=9).value = num(t.get("advance_limit"))
        # объект вообще без данных в снапшоте
        if not any((obj, p) in by_obj_period for p in
                   ("2023", "2024", "2025", "2025H1", "2025H2", "2026")):
            empty_objects.append(obj)

    out_name = args.output or f"{OUTPUT_DIR}/Demo_Report_{actual_date}.xlsx"

    # --- быстрый путь вебхука: файл уже есть — не генерируем заново ---
    if args.reuse_existing and os.path.exists(out_name):
        print(json.dumps({
            "status": "ok",
            "cached": True,
            "report_date": actual_date,
            "requested_date": args.report_date,
            "fallback": fallback_used,
            "trigger_source": used_source,
            "output": out_name,
        }, ensure_ascii=True, indent=1))
        return

    os.makedirs(os.path.dirname(out_name), exist_ok=True)
    wb.save(out_name)

    print(json.dumps({
        "status": "ok",
        "cached": False,
        "report_date": actual_date,
        "requested_date": args.report_date,
        "fallback": fallback_used,
        "trigger_source": used_source,
        "snapshot_rows": len(snap),
        "objects": len(SECTIONS),
        "missing_terms": missing_terms,
        "objects_without_data": empty_objects,
        "bg_contracts": sorted(bg_by_contract),
        "output": out_name,
    }, ensure_ascii=True, indent=1))

if __name__ == "__main__":
    main()
