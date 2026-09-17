@echo off
REM ============================================================================
REM  Start.bat -- Acceptance PDF Automation launcher
REM ============================================================================
REM  Double-click this file to run the application. It will:
REM    1. Find the project folder automatically (wherever it's been copied to)
REM    2. Check that the Python virtual environment ("venv") exists
REM    3. Activate that virtual environment
REM    4. Verify Python and the required packages are actually available
REM    5. Warn (but not block) if LibreOffice is missing -- only needed for
REM       Workflow 2 (Prepare Final Signed Documents)
REM    6. Start the Flask server in THIS window, so its log output and any
REM       errors stay visible
REM    7. Open your default browser to the app once the server responds
REM
REM  To stop the application: just close this window (or press Ctrl+C).
REM  The Flask server is a direct child of this window, so closing the
REM  window stops it immediately -- no separate "Stop" step is needed in
REM  normal use. See Stop.bat for the rare cleanup-only fallback case.
REM
REM  This file is safe to copy along with the rest of the project folder to
REM  a different computer or a different folder path -- nothing here is
REM  hardcoded to a specific machine or location.
REM ============================================================================

setlocal enabledelayedexpansion
title Acceptance PDF Automation - Server

REM --- Step 1: Always run from the folder this .bat file lives in -----------
REM %~dp0 expands to the full path of THIS script's folder (with a trailing
REM backslash), regardless of where the user copied the project or which
REM folder they double-clicked from. This is what makes the launcher
REM portable across machines instead of relying on a hardcoded path.
cd /d "%~dp0"

echo ============================================================
echo   Acceptance PDF Automation
echo ============================================================
echo   Project folder: %cd%
echo ============================================================
echo.

REM --- Step 2: Is a previous instance already running on port 5000? ---------
REM Avoids a confusing "port already in use" crash from Flask -- instead we
REM just point the browser at the instance that's already up.
netstat -ano | findstr ":5000" | findstr "LISTENING" >nul 2>nul
if not errorlevel 1 (
    echo [INFO] The application already appears to be running on port 5000.
    echo        Opening your browser to it now...
    start "" "http://127.0.0.1:5000"
    echo.
    echo If this isn't expected, run Stop.bat first and then try Start.bat again.
    echo.
    pause
    exit /b 0
)

REM --- Step 3: Confirm the virtual environment exists ------------------------
if not exist "venv\Scripts\python.exe" (
    echo [ERROR] The Python virtual environment "venv" was not found.
    echo.
    echo This is a one-time setup step that hasn't been done yet on this
    echo computer. To fix it, open Command Prompt in this folder and run:
    echo.
    echo     python -m venv venv
    echo     venv\Scripts\pip install -r requirements.txt
    echo.
    echo Then double-click Start.bat again.
    echo.
    pause
    exit /b 1
)

REM --- Step 4: Activate the virtual environment -------------------------------
echo [1/4] Activating the virtual environment...
call "venv\Scripts\activate.bat"
if errorlevel 1 (
    echo [ERROR] Failed to activate the virtual environment.
    echo The "venv" folder may be corrupted or incomplete.
    echo.
    echo To fix this, delete the "venv" folder and run:
    echo.
    echo     python -m venv venv
    echo     venv\Scripts\pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)
echo       Done.
echo.

REM --- Step 5: Verify Python actually resolves inside the activated venv -----
echo [2/4] Checking Python...
python --version >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python could not be found even after activating the virtual
    echo environment. The "venv" folder is likely corrupted.
    echo.
    echo To fix this, delete the "venv" folder and run:
    echo.
    echo     python -m venv venv
    echo     venv\Scripts\pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)
for /f "delims=" %%V in ('python --version 2^>^&1') do echo       Found %%V.
echo.

REM --- Step 6: Verify the required Python packages are installed -------------
echo [3/4] Checking required packages...
python -c "import flask, flask_login, flask_sqlalchemy, flask_wtf, pandas, openpyxl, docx, fitz" 2>nul
if errorlevel 1 (
    echo [ERROR] One or more required Python packages are missing.
    echo.
    echo To install everything the application needs, run:
    echo.
    echo     venv\Scripts\pip install -r requirements.txt
    echo.
    echo Then double-click Start.bat again.
    echo.
    pause
    exit /b 1
)
echo       All required packages are installed.
echo.

REM --- Step 7: Check for LibreOffice (warning only, not blocking) ------------
REM Only Workflow 2 (Prepare Final Signed Documents) needs LibreOffice, to
REM convert signed Word documents to PDF. Workflow 1 works fine without it,
REM so a missing installation is a warning, not a reason to stop.
echo [4/4] Checking for LibreOffice...
where soffice >nul 2>nul
if errorlevel 1 (
    echo       [WARNING] LibreOffice was not found on this computer's PATH.
    echo                 Workflow 1 ^(Generate Acceptance Test Documents^) will
    echo                 still work normally.
    echo                 Workflow 2 ^(Prepare Final Signed Documents^) needs
    echo                 LibreOffice and will show an error until it's installed.
    echo                 Download: https://www.libreoffice.org/download/
) else (
    echo       Found.
)
echo.

echo ============================================================
echo   Starting the server...
echo ============================================================
echo   Once it's ready, your browser will open automatically to:
echo       http://127.0.0.1:5000
echo.
echo   Keep this window open while you use the application.
echo   To stop the application, simply close this window.
echo ============================================================
echo.

REM --- Step 8: Open the browser automatically once the server responds -------
REM This runs as a short-lived, invisible helper alongside the server: it
REM polls the local URL once a second (up to 30 seconds) and opens the
REM default browser the moment the server responds, then exits on its own.
REM -WindowStyle Hidden keeps it invisible; -Command (rather than a .ps1
REM script file) avoids being blocked by a machine's PowerShell execution
REM policy, which typically only restricts running script files.
start "" powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 3; Start-Process 'http://127.0.0.1:5000'"


REM --- Step 9: Start the Flask server in THIS window (foreground) ------------
REM Running it directly (not in a background/detached process) is what keeps
REM this window open with live logs and errors visible, and is also what
REM makes closing this window stop the server -- it's a normal child process
REM of this console session, so Windows ends it when the console closes.
python main.py

REM --- Step 10: If we get here, the server stopped on its own ----------------
REM (e.g. it crashed, or was stopped some other way, rather than the user
REM closing this window). Keep the window open so any error above stays
REM readable instead of the window vanishing immediately.
echo.
echo ============================================================
echo   The server has stopped.
echo ============================================================
echo   If this was unexpected, scroll up to see if an error was printed.
echo.
pause
