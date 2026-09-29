@echo off
cd /d "%~dp0"
call "%~dp0_python.bat" || exit /b 1
"%PYEXE%" %PYARGS% -m http.server 8502
