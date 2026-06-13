@echo off
chcp 65001 >nul
title e-WorkPermit Report
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0start.ps1"
