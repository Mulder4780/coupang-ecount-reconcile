@echo off
rem 쿠팡 통합업무 앱 서버 + 외부접속 터널 (PC·휴대폰 브라우저 접속용)
set PYTHONIOENCODING=utf-8
start "" /min "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0webapp\app_server.py"
start "" /min "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0webapp\tunnel_run.py"
