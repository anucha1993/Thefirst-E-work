@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\pythonw.exe" gui_app.py
) else (
    pythonw gui_app.py
)
