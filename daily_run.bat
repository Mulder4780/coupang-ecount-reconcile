@echo off
rem 쿠팡 업무 자동대조 에이전트 — Windows 작업 스케줄러 등록용
rem 등록(1회):  schtasks /Create /TN "쿠팡업무_일일자동대조" /TR "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount\daily_run.bat" /SC DAILY /ST 09:50
rem ★ 아침 입력 보호시간(08:00~09:30)은 2026-08-11 지시로 **퇴역**했다 — 사람은 이제
rem   엑셀이 아니라 앱에만 입력한다. `워치독실행.bat` 은 2026-08-12 에 옮겼는데
rem   **이 파일이 빠져 있었다**(분담판 [50]). 09:50 회차라 지금 당장 막히지는 않지만,
rem   판단이 두 곳에 있으면 한쪽만 고쳐진다 — 그것이 [44] 사고의 모양이었다.
rem   판단은 operation_window 한 곳이다. 되돌리려면 COUPANG_INPUT_WINDOW=08:00-09:30.
set PYTHONIOENCODING=utf-8
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" -c "import sys,os;sys.path.insert(0,r'%~dp0');import operation_window as W;sys.exit(10 if W.is_input_window() else 0)"
if errorlevel 10 exit /b 0
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0daily_run.py" >> "%~dp0reports\daily_run_log.txt" 2>&1
