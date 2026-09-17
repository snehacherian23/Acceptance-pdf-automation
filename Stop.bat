@echo off
REM ============================================================================
REM  Stop.bat -- Safety-net cleanup for Acceptance PDF Automation
REM ============================================================================
REM  You normally do NOT need this file.
REM
REM  The Flask server started by Start.bat runs directly inside that same
REM  Command Prompt window -- it is not a separate background process. That
REM  means simply CLOSING the Start.bat window (or pressing Ctrl+C inside it)
REM  already stops the application immediately. That's the normal way to
REM  stop the app.
REM
REM  This script exists only as a fallback for the unusual case where a
REM  server process got left running anyway -- for example, if Start.bat was
REM  run more than once, or the window was closed in a way that didn't fully
REM  end the process. It finds whatever is listening on port 5000 (the
REM  application's port) and stops it.
REM ============================================================================

setlocal enabledelayedexpansion
title Acceptance PDF Automation - Stop

echo ============================================================
echo   Acceptance PDF Automation - Stop
echo ============================================================
echo.

set FOUND=0
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":5000" ^| findstr "LISTENING"') do (
    set FOUND=1
    echo Stopping process with ID %%P on port 5000...
    taskkill /PID %%P /F >nul 2>nul
)

echo.
if "!FOUND!"=="0" (
    echo No running instance was found on port 5000.
    echo The application does not appear to be running.
) else (
    echo The application has been stopped.
)

echo.
pause
