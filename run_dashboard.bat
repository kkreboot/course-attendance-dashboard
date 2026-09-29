@echo off
cd /d "%~dp0"
call "%~dp0_python.bat" || exit /b 1
"%PYEXE%" %PYARGS% -m streamlit run dashboard.py --server.port 8501 --server.headless true
