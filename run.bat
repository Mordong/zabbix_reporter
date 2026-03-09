@echo off
title Zabbix Reporter

:: Проверяем Python
python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo [ОШИБКА] Python не найден. Установите Python 3.10+ с https://python.org
    pause
    exit /b 1
)

:: Устанавливаем зависимости при первом запуске
IF NOT EXIST ".deps_installed" (
    echo Установка зависимостей...
    pip install openpyxl reportlab matplotlib --quiet
    echo. > .deps_installed
)

:: Запуск приложения
python zabbix_reporter.py
pause
