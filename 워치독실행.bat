@echo off
rem 자가치유 워치독 — 30분마다 작업 스케줄러가 호출 (서버·터널 복구, 버전 아카이브, 리포트 정리)
set PYTHONIOENCODING=utf-8
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0watchdog.py" >> "%~dp0reports\watchdog_log.txt" 2>&1
