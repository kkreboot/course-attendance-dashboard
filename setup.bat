@echo off
rem One-time setup on a Windows machine: creates .venv-Windows and installs
rem requirements.txt into it. Run this once, then use the launcher .bat files.
rem
rem The venv is per-OS (.venv-Windows / .venv-Darwin / .venv-Linux) because
rem this folder is Dropbox-synced across three machines and a venv is not
rem portable -- a single shared .venv only works on whichever machine built
rem it last. See setup.sh for the macOS/Linux side.
setlocal
cd /d "%~dp0"

rem Resolve a *system* interpreter to build the venv with. _python.bat is not
rem used here because it prefers .venv-Windows, which is the thing being built.
set "PYEXE="
set "PYARGS="
if exist "C:\Users\user\AppData\Local\Programs\Python\Python312\python.exe" set "PYEXE=C:\Users\user\AppData\Local\Programs\Python\Python312\python.exe"
if not defined PYEXE where py >nul 2>&1 && (set "PYEXE=py" & set "PYARGS=-3")
if not defined PYEXE where python >nul 2>&1 && set "PYEXE=python"
if not defined PYEXE (
    echo Python not found. Install Python 3.10+ from python.org first. 1>&2
    exit /b 1
)

if not exist ".venv-Windows\Scripts\python.exe" (
    echo Creating venv in .venv-Windows...
    "%PYEXE%" %PYARGS% -m venv .venv-Windows
    if errorlevel 1 exit /b 1
)

echo Installing requirements.txt...
".venv-Windows\Scripts\python.exe" -m pip install --upgrade pip
".venv-Windows\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 exit /b 1

rem Fail here rather than at the first click if the install came up short.
".venv-Windows\Scripts\python.exe" -c "import pandas, streamlit, reportlab, openpyxl, xlrd; print('OK: pandas', pandas.__version__, '| streamlit', streamlit.__version__)"
if errorlevel 1 exit /b 1

echo.
echo Setup complete (venv: .venv-Windows).
echo   - Double-click launch_dashboard_chrome.bat to start the dashboard
echo   - Or .venv-Windows\Scripts\python.exe run.py --help for the CLI
