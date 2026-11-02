#!/bin/bash
# reo_etl.sh — ETL отчёта Demo_Report (cron и ручной запуск).
# Использование:
#   reo_etl.sh            — плановый запуск (trigger_source=scheduled, для cron)
#   reo_etl.sh manual     — ручной запуск между сессиями (trigger_source=manual)
# Cron (пн 07:00): 0 7 * * 1 /home/demo/analytics-portal/data/scripts/reo_etl.sh >> /home/demo/analytics-portal/data/reo_etl.log 2>&1
# Креды 1С и NocoDB — в environment контейнера python_worker (docker-compose).
set -e
cd /home/demo/analytics-portal

SOURCE="${1:-scheduled}"
case "$SOURCE" in
  scheduled|manual) ;;
  *) echo "USAGE: reo_etl.sh [scheduled|manual]"; exit 1 ;;
esac

echo "=== REO ETL start ($SOURCE): $(date '+%Y-%m-%d %H:%M:%S') ==="

# 1. Выгрузка из 1С + агрегация + запись снапшота в NocoDB
#    (report_date transform вычисляет сам: сегодня-2 раб. дня; печатает в JSON-сводке)
TRANSFORM_OUT=$(docker exec python_worker python /data/scripts/transform_reo_v1.0.py --fetch --trigger-source "$SOURCE")
echo "$TRANSFORM_OUT"
REPORT_DATE=$(echo "$TRANSFORM_OUT" | grep -o '"report_date": *"[0-9-]*"' | grep -o '[0-9-]*' | tail -1)
if [ -z "$REPORT_DATE" ]; then
    echo "ERROR: could not parse report_date from transform output"; exit 1
fi
echo "REPORT_DATE=$REPORT_DATE"

# 2. Генерация Excel из снапшота -> /data/output/reo/Demo_Report_<report_date>.xlsx
docker exec python_worker python /data/scripts/generate_excel_reo_v1.0.py \
    --report-date "$REPORT_DATE" --trigger-source "$SOURCE"

echo "=== REO ETL done: $(date '+%Y-%m-%d %H:%M:%S') ==="
