@echo off
chcp 65001 >nul
title CoopBase Server
cd /d "%~dp0"
set "PY=.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"
echo Starting server...
"%PY%" app.py
pause
