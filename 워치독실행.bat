@echo off
rem 자가치유 워치독 — 30분마다 작업 스케줄러가 호출 (서버·터널 복구, 리포트 정리)
rem ★ 아침 입력 보호시간(08:00~09:30)은 2026-08-11 지시로 **퇴역**했다 — 사람은 이제
rem   엑셀이 아니라 앱에만 입력한다. 그런데 퇴역이 파이썬(operation_window)에서만
rem   되고 여기 하드코딩이 남아, 2026-08-12 까지 워치독이 **매일 아침 90분 통째로
rem   안 돌았다**(자가치유·회차감시·인계갱신·붙여넣기 치유가 그동안 전부 멈춘다).
rem   스케줄러는 exit 0 을 '성공'으로 적으므로 어느 화면에도 티가 안 났다 — 분담판 [44].
rem   판단은 이제 operation_window 한 곳이다. 사본이 둘이면 한쪽만 고쳐진다.
rem   되돌리려면 환경변수 COUPANG_INPUT_WINDOW=08:00-09:30 하나면 된다.
set PYTHONIOENCODING=utf-8
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" -c "import sys,os;sys.path.insert(0,r'%~dp0');import operation_window as W;sys.exit(10 if W.is_input_window() else 0)"
if errorlevel 10 exit /b 0
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" "%~dp0watchdog.py" >> "%~dp0reports\watchdog_log.txt" 2>&1
