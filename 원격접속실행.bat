@echo off
rem 외부(LTE/5G) 접속 터널 실행 — 사전 1회: winget install Cloudflare.cloudflared
set PYTHONIOENCODING=utf-8
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0webapp\tunnel_run.py"
pause
