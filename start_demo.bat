@echo off
chcp 65001 >nul
title Analytics Demo
cd /d "%~dp0"

set PY=

python --version >nul 2>nul
if not errorlevel 1 set PY=python

if "%PY%"=="" py --version >nul 2>nul
if "%PY%"=="" if not errorlevel 1 set PY=py

if "%PY%"=="" (
    echo Python не найден. Установите его с https://www.python.org/downloads/
    echo и отметьте галочку "Add python.exe to PATH" при установке.
    pause
    exit /b 1
)

%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Первый запуск: устанавливаю зависимость openpyxl...
    %PY% -m pip install --quiet openpyxl
)

echo Запуск демо-портала: http://127.0.0.1:8000
echo Закройте это окно, чтобы остановить сервер.
%PY% demo_server.py
pause
