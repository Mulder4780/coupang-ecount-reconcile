@echo off
setlocal
cd /d "%~dp0"
set "PYTHONIOENCODING=utf-8"
set "PY=C:\Users\hueng\AppData\Local\Programs\Python\Python312\python.exe"
if not exist "reports" mkdir "reports"
"%PY%" collect_sources.py --apply >> "reports\source_organizer.log" 2>&1
"%PY%" source_organizer.py --apply >> "reports\source_organizer.log" 2>&1
endlocal
