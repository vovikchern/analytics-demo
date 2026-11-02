@echo off
chcp 65001 >nul
title Analytics Demo
cd /d "%~dp0"

rem --- поиск Python: python в PATH или py-лаунчер ---
set PY=
where python >nul 2>nul && set PY=python
if not defined PY where py >nul 2>nul && set PY=py
if not defined PY (
    echo Python не найден. Установите его с https://www.python.org/downloads/
    echo и отметьте галочку "Add python.exe to PATH" при установке.
    pause
    exit /b 1
)

rem --- openpyxl: ставим автоматически, если нет ---
%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Первая установка зависимости openpyxl...
    %PY% -m pip install --quiet openpyxl
)

echo Запуск демо-портала: http://127.0.0.1:8000  ^(окно можно свернуть, закрытие окна = остановка сервера^)
%PY% demo_server.py
pause
