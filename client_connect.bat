@echo off
chcp 65001 >nul
title CoopBase Client
cd /d "%~dp0"
set "PY=.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"
echo Starting client connector...
"%PY%" client_connect.py
pause
