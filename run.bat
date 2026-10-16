@echo off
rem ===================================================================
rem  Class Score - one click launcher
rem
rem  IMPORTANT: keep this file ASCII-only.
rem  cmd.exe parses .bat files with the system ANSI code page, so Chinese
rem  text in a .bat file can be split into broken commands. All Chinese
rem  messages live in launcher.py instead (Python handles UTF-8 fine).
rem ===================================================================
setlocal
cd /d "%~dp0"

rem -------------------------------------------------------------------
rem  1) Never ask for an email address, never send usage statistics.
rem
rem  Without this the very first run prints:
rem      Welcome to Streamlit! ... Email:
rem  and waits for input. Pressing Enter (blank email) counts as INVALID
rem  and Streamlit exits with code -1 - so this window used to vanish
rem  without any message. Three switches are used on purpose:
rem    - .streamlit/config.toml  (server.showEmailPrompt = false)
rem    - the two env vars below  (read by the streamlit CLI)
rem    - --server.showEmailPrompt=false  (added in launcher.py)
rem -------------------------------------------------------------------
set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"
set "STREAMLIT_SERVER_SHOW_EMAIL_PROMPT=false"

rem -------------------------------------------------------------------
rem  2) Package index (download source) for installing dependencies.
rem  Default: Tsinghua mirror - much faster inside mainland China.
rem  Use "official" for the default PyPI, or any other mirror URL.
rem -------------------------------------------------------------------
set "CS_PIP_INDEX=https://pypi.tuna.tsinghua.edu.cn/simple"

rem -------------------------------------------------------------------
rem  3) Find a Python interpreter.
rem -------------------------------------------------------------------
set "PY="
where python >nul 2>nul
if not errorlevel 1 set "PY=python"
if not defined PY (
    where py >nul 2>nul
    if not errorlevel 1 set "PY=py -3"
)
if not defined PY goto nopython

%PY% "%~dp0launcher.py" %*
set "CODE=%ERRORLEVEL%"
if not "%CODE%"=="0" (
    echo.
    echo   [X] The launcher stopped with exit code %CODE%. Read the messages above.
)
goto done

:nopython
echo.
echo   [X] Python not found on this computer.
echo.
echo       1. Download Python 3.10 - 3.13 from https://www.python.org/downloads/
echo       2. During setup, TICK the checkbox: "Add python.exe to PATH"
echo       3. Close this window, then double click run.bat again.
echo.
set "CODE=1"

:done
echo.
echo   ---- Press any key to close this window ----
pause >nul
exit /b %CODE%
