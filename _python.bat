@echo off
rem Shared Python resolver for the .bat launchers. Sets PYEXE (+ PYARGS) to
rem the best interpreter available on this Windows machine.
rem
rem Order: this machine's own venv (.venv-Windows, made by setup.bat), then a
rem legacy shared .venv, then the old hardcoded install path, then the `py`
rem launcher, then plain `python` on PATH.
rem
rem Per-OS venv dirs exist because the project folder is Dropbox-synced across
rem Windows/macOS/Linux and a venv is not portable between them.
set "PYEXE="
set "PYARGS="

if exist "%~dp0.venv-Windows\Scripts\python.exe" set "PYEXE=%~dp0.venv-Windows\Scripts\python.exe"
if not defined PYEXE if exist "%~dp0.venv\Scripts\python.exe" set "PYEXE=%~dp0.venv\Scripts\python.exe"
if not defined PYEXE if exist "C:\Users\user\AppData\Local\Programs\Python\Python312\python.exe" set "PYEXE=C:\Users\user\AppData\Local\Programs\Python\Python312\python.exe"
if not defined PYEXE where py >nul 2>&1 && (set "PYEXE=py" & set "PYARGS=-3")
if not defined PYEXE where python >nul 2>&1 && set "PYEXE=python"

if not defined PYEXE (
    echo Python not found. Install Python 3.10+ from python.org, then run setup.bat. 1>&2
    exit /b 1
)
exit /b 0
