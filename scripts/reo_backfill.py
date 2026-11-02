# -*- coding: utf-8 -*-
"""
reo_backfill.py — восполнение истории снапшотов из 1С (решение 2026-11-01, вариант A).

Даты (согласованы с пользователем):
  - помесячно (последний день месяца) с 2023-12-31 по 2026-06-30 — 31 дата;
  - еженедельно по вторникам с 2026-07-07 по 2026-10-27 — 17 дат.
Все снапшоты пишутся с trigger_source=BACKFILL — попадают в график динамики
(scheduled + backfill), но не смешиваются с плановыми cron-снапшотами.

Запуск: docker exec python_worker python /data/scripts/reo_backfill.py
Каждый прогон ~10 сек (выгрузка из 1С + агрегация + запись). Итого ~8-10 минут.
"""
import subprocess
import sys
import time

TRANSFORM = "/data/scripts/transform_reo_v1.0.py"

MONTH_ENDS = []
def _me(y, m):
    import calendar
    last_day = calendar.monthrange(y, m)[1]
    MONTH_ENDS.append(f"{y:04d}-{m:02d}-{last_day:02d}")

_me(2023, 12)
for y in (2024, 2025):
    for m in range(1, 13):
        _me(y, m)
for m in range(1, 7):
    _me(2026, m)

TUESDAYS = [
    "2026-07-07", "2026-07-14", "2026-07-21", "2026-07-28",
    "2026-08-04", "2026-08-11", "2026-08-18", "2026-08-25",
    "2026-09-01", "2026-09-08", "2026-09-15", "2026-09-22", "2026-09-29",
    "2026-10-06", "2026-10-13", "2026-10-20", "2026-10-27",
]

DATES = sorted(set(MONTH_ENDS + TUESDAYS))


def main():
    print(f"BACKFILL START: {len(DATES)} дат", flush=True)
    ok, fail = 0, 0
    for i, d in enumerate(DATES, 1):
        t0 = time.time()
        proc = subprocess.run(
            [sys.executable, TRANSFORM, "--fetch",
             "--report-date", d, "--trigger-source", "backfill"],
            capture_output=True, text=True, timeout=300)
        dt = time.time() - t0
        if proc.returncode == 0:
            ok += 1
            print(f"[{i}/{len(DATES)}] {d} OK ({dt:.1f}s)", flush=True)
        else:
            fail += 1
            tail = ((proc.stdout or "") + (proc.stderr or "")).strip()[-300:]
            print(f"[{i}/{len(DATES)}] {d} FAIL: {tail}", flush=True)
    print(f"BACKFILL DONE: ok={ok} fail={fail}", flush=True)


if __name__ == "__main__":
    main()
