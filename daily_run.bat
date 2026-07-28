@echo off
rem 쿠팡 업무 자동대조 에이전트 — Windows 작업 스케줄러 등록용
rem 등록(1회):  schtasks /Create /TN "쿠팡업무_일일자동대조" /TR "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount\daily_run.bat" /SC DAILY /ST 09:50
powershell -NoProfile -Command "$t=(Get-Date).TimeOfDay; if($t -ge [TimeSpan]::FromHours(8) -and $t -lt [TimeSpan]::FromMinutes(570)){exit 10}"
if errorlevel 10 exit /b 0
set PYTHONIOENCODING=utf-8
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0daily_run.py" >> "%~dp0reports\daily_run_log.txt" 2>&1
