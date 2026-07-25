@echo off
rem 휴대폰 접속 QR — 현재 살아있는 터널 주소를 QR로 띄운다(주소가 바뀌어도 항상 최신)
set PYTHONIOENCODING=utf-8
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0phone_access.py"
