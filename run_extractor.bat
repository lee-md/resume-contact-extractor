@echo off
REM Double-click to choose a folder, or drag a folder / files onto this icon.
cd /d "%~dp0"
set EXTRACTOR_PAUSE=1
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 "%~dp0extract_contacts.py" %*
) else (
    python "%~dp0extract_contacts.py" %*
)
if errorlevel 9009 (
    echo.
    echo Python was not found. Install it from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" during setup.
    pause
)
