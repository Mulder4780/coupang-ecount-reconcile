@echo off
rem Coupang Service Operations System 앱 서버 + 외부접속 터널 (PC·휴대폰 브라우저 접속용)
set PYTHONIOENCODING=utf-8
start "" /b "%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe" "%~dp0webapp\app_server.py"
start "" /b "%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe" "%~dp0webapp\tunnel_run.py"
