@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set "PYTHONIOENCODING=utf-8"
title Разложить прайсы

set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY python -c "import sys" >nul 2>nul && set "PY=python"
if not defined PY (
    echo.
    echo Не найден Python 3.
    echo Установите его с https://www.python.org/downloads/
    echo При установке поставьте галочку "Add python.exe to PATH", затем запустите этот файл ещё раз.
    echo.
    pause
    exit /b 1
)

%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Первый запуск: устанавливаю модуль openpyxl...
    %PY% -m pip install --user --disable-pip-version-check -q -r "Программа\requirements.txt"
)

%PY% main.py --sort-prices
pause
