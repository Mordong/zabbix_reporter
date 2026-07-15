@echo off
chcp 65001 >nul
title Zabbix Reporter - Build EXE
cd /d "%~dp0"

rem ── Поиск полноценного Python ────────────────────────────────────────────
rem Проверки "python --version" недостаточно: встроенные интерпретаторы
rem сторонних программ (Inkscape, GIMP и т.п.) отвечают на неё, но не
rem содержат pip. Поэтому проверяем именно наличие pip: сначала через
rem лаунчер "py", затем "python" из PATH.
set "PY="
py -3 -m pip --version >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if not defined PY (
    python -m pip --version >nul 2>nul
    if not errorlevel 1 set "PY=python"
)
if not defined PY (
    echo [ОШИБКА] Не найден Python с модулем pip.
    echo Установите Python 3.8+ с https://www.python.org/downloads/
    echo ^(при установке отметьте "Add python.exe to PATH" и "py launcher"^).
    pause
    exit /b 1
)
echo Используется: %PY%
%PY% --version

echo [1/3] Installing dependencies...
%PY% -m pip install --upgrade pip --quiet --no-warn-script-location
%PY% -m pip install pyinstaller openpyxl reportlab matplotlib --quiet --no-warn-script-location
if errorlevel 1 (echo [ERROR] Failed & pause & exit /b 1)
echo Done.
echo [2/3] Building EXE (1-3 min)...
%PY% -m PyInstaller --onefile --windowed --name ZabbixReporter --clean --noconfirm --hidden-import openpyxl --hidden-import openpyxl.styles --hidden-import openpyxl.utils --hidden-import openpyxl.chart --hidden-import reportlab --hidden-import reportlab.pdfbase.ttfonts --hidden-import reportlab.platypus --hidden-import reportlab.graphics.charts.barcharts --hidden-import reportlab.graphics.charts.piecharts --hidden-import matplotlib --hidden-import matplotlib.backends.backend_tkagg zabbix_reporter.py
if errorlevel 1 (echo [ERROR] Build failed & pause & exit /b 1)
echo [3/3] Done!
if exist dist\ZabbixReporter.exe (echo EXE ready: dist\ZabbixReporter.exe & pause & explorer dist) else (echo [WARNING] EXE not found & pause)
