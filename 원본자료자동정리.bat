@echo off
setlocal
cd /d "%~dp0"
set "PYTHONIOENCODING=utf-8"
set "PY=C:\Users\hueng\AppData\Local\Programs\Python\Python312\python.exe"
if not exist "reports" mkdir "reports"
rem ★ 네 단계를 여기서 이어 달리지 않는다 — 단계마다 제한이 없어 한 단계가 회차 제한
rem    (PT3H)을 다 먹고 나머지가 아예 안 돌았다(2026-08-19 실사고). 어느 단계였는지도
rem    남지 않았다. source_tidy_run.py 가 단계마다 제한을 걸고 자국을 남긴다.
rem ★ python.exe 를 그대로 쓴다 — 창은 run_hidden.vbs 가 이미 숨긴다. pythonw 로 바꾸면
rem    실행기 자신이 import 에서 죽을 때 그 트레이스백이 아무 데도 안 남는다([248]).
"%PY%" source_tidy_run.py >> "reports\source_organizer.log" 2>&1
rem ★ 종료코드를 그대로 돌려준다 — 삼키면 죽은 회차가 성공으로 적히고 감시자의 눈이 먼다([248]).
endlocal & exit /b %ERRORLEVEL%
