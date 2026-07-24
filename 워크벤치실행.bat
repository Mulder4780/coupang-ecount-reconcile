@echo off
rem 쿠팡 통합업무 워크벤치 실행 (더블클릭)
set PYTHONIOENCODING=utf-8
start "" "%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe" "%~dp0coupang_workbench.py"
