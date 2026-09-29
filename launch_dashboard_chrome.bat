@echo off
rem Starts the Streamlit dashboard server in the background, waits for it to
rem come up, then opens it in Chrome. Double-click this file (or the desktop
rem shortcut pointing to it) to go straight from a cold start to the
rem dashboard in a browser tab -- no terminal typing required.
setlocal
cd /d "%~dp0"

rem Interpreter comes from _python.bat: this machine's .venv-Windows first,
rem then a legacy .venv, then a system Python. (Per-OS venv dirs -- the folder
rem is Dropbox-synced across Windows/macOS/Linux and venvs are not portable.)
call "%~dp0_python.bat" || exit /b 1

set "CHROME=C:\Program Files\Google\Chrome\Application\chrome.exe"
set "PORT=8501"
set "URL=http://localhost:%PORT%"

rem --headless so Streamlit doesn't try to launch its own browser tab; we
rem open Chrome ourselves below once the server is actually ready.
start "Course Dashboard Server" /min "%PYEXE%" %PYARGS% -m streamlit run dashboard.py --server.port %PORT% --server.headless true

rem Poll the port instead of a fixed sleep, so this works whether the server
rem takes one second or ten (first run after a reboot is slower to import).
:wait
powershell -NoProfile -Command "try { (New-Object Net.Sockets.TcpClient('localhost', %PORT%)).Close(); exit 0 } catch { exit 1 }" >nul 2>&1
if errorlevel 1 (
    timeout /t 1 /nobreak >nul
    goto wait
)

if exist "%CHROME%" (
    start "" "%CHROME%" --new-window "%URL%"
) else (
    start "" "%URL%"
)
