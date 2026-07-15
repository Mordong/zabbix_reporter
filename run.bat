@echo off
rem Файл сохранён в UTF-8; переключаем консоль на UTF-8 (65001),
rem иначе русские сообщения отображаются как нечитаемые символы (CP866).
chcp 65001 >nul
cd /d "%~dp0"

rem ── Поиск полноценного Python ────────────────────────────────────────────
rem Встроенные интерпретаторы сторонних программ (Inkscape, GIMP и т.п.)
rem попадают в PATH раньше системного Python, но не содержат pip.
rem Поэтому проверяем наличие pip: сначала через лаунчер "py", затем "python".
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
%PY% -m pip install -r requirements.txt --quiet
%PY% zabbix_reporter.py
