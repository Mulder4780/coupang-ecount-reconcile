@echo off
rem 실시간 문제감시 — 4시간마다 작업 스케줄러가 호출.
rem ★ 2026-08-13: 창 없이 돈다(`창없이실행.vbs`). 그래서 **출력을 반드시 파일로
rem   남긴다** — 전에는 콘솔로만 나갔고 그 창은 스쳐 지나가 아무도 못 봤다.
rem   창을 숨기면서 로그를 안 붙이면 '조용히 아무 데도 안 남는' 상태가 된다([169]).
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
if not exist "reports" mkdir "reports"
python realtime_monitor.py >> "reports\realtime_monitor_log.txt" 2>&1
