@echo off
rem Build WhatsApp_Number_Checker_Web.exe into .\dist
rem Uses a local virtual environment (.venv) with the pinned Playwright 1.63.0.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Creating .venv ...
    python -m venv .venv || exit /b 1
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt || exit /b 1
)
".venv\Scripts\python.exe" -m PyInstaller --clean --noconfirm WhatsApp_Number_Checker_Web.spec
echo.
echo Done: dist\WhatsApp_Number_Checker_Web.exe
echo Put it next to a "Browsers" folder (Chromium 1243) and an "ExcelFiles" folder.
pause
