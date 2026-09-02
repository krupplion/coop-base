@echo off
setlocal
REM ===== Coop Base - PyInstaller build script (double-click to run) =====
REM NOTE: Keep this file PURE ASCII. Chinese text in a .bat breaks when the
REM       console codepage differs (GBK vs UTF-8 across terminals).
set "VENV=.venv"
set "PYI=%VENV%\Scripts\pyinstaller.exe"
set "SRC=%~dp0"
if "%SRC:~-1%"=="\" set "SRC=%SRC:~0,-1%"

cd /d "%SRC%" || exit /b 1

REM Build args live in server.spec / client.spec (single source of config).
REM Do NOT pass CLI args here, or the two configs will drift.
REM rapidocr_onnxruntime submodules are imported dynamically via importlib,
REM so collect_submodules / collect_data_files in the spec are REQUIRED,
REM otherwise invoice OCR breaks in the packaged exe.

echo [0/3] Cleaning previous build outputs ...
rd /s /q "%SRC%\build\server" 2>nul
rd /s /q "%SRC%\build\client" 2>nul
rd /s /q "%SRC%\dist\server" 2>nul
rd /s /q "%SRC%\dist\client" 2>nul
del /q "%SRC%\out\setup_coop.exe" 2>nul

echo [1/3] Building server.exe (LAN server, prints pairing code) ...
"%PYI%" --noconfirm server.spec
if errorlevel 1 (echo server build FAILED & exit /b 1)

echo [2/3] Building client.exe (client connector, asks pairing code) ...
"%PYI%" --noconfirm client.spec
if errorlevel 1 (echo client build FAILED & exit /b 1)

echo [3/3] Packing NSIS installer (bundled portable NSIS) ...
if not exist "install\nsis\makensis.exe" (
  echo makensis.exe not found, skip installer; outputs are in dist\
  goto :done
)
if not exist "out" mkdir "out"
"install\nsis\makensis.exe" install\coop.nsi
if errorlevel 1 (echo NSIS pack FAILED & exit /b 1)

:done
echo.
echo All done!
echo   Program dirs : dist\server\  and  dist\client\
echo   Installer    : out\setup_coop.exe
endlocal
