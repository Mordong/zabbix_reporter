@echo off
chcp 1251 >nul
title Zabbix Reporter - Build EXE
python --version >nul 2>&1
if errorlevel 1 (echo [ERROR] Python not found. & pause & exit /b 1)
python --version
echo [1/3] Installing dependencies...
python -m pip install --upgrade pip --quiet --no-warn-script-location
python -m pip install pyinstaller openpyxl reportlab matplotlib --quiet --no-warn-script-location
if errorlevel 1 (echo [ERROR] Failed & pause & exit /b 1)
echo Done.
echo [2/3] Building EXE (1-3 min)...
python -m PyInstaller --onefile --windowed --name ZabbixReporter --clean --noconfirm --hidden-import openpyxl --hidden-import openpyxl.styles --hidden-import openpyxl.utils --hidden-import openpyxl.chart --hidden-import reportlab --hidden-import reportlab.pdfbase.ttfonts --hidden-import reportlab.platypus --hidden-import reportlab.graphics.charts.barcharts --hidden-import reportlab.graphics.charts.piecharts --hidden-import matplotlib --hidden-import matplotlib.backends.backend_tkagg zabbix_reporter.py
if errorlevel 1 (echo [ERROR] Build failed & pause & exit /b 1)
echo [3/3] Done!
if exist dist\ZabbixReporter.exe (echo EXE ready: dist\ZabbixReporter.exe & pause & explorer dist) else (echo [WARNING] EXE not found & pause)
