# Mulder Control Hub 주간 DB 회차 설치 : 매주 월요일 07:00 · 창 없이(pythonw)
# 2026-11-17 형님 지시 "멀더 컨트롤 허브로 1주일에 1회씩 정리된 db 보내자 앞으로 쭉"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$py = (Get-Command pythonw.exe -ErrorAction Stop).Source
$act = New-ScheduledTaskAction -Execute $py -Argument "`"$here\mulder_export.py`"" -WorkingDirectory $here
$trg = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 7:00am
$set = New-ScheduledTaskSettingsSet -StartWhenAvailable -Hidden -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
Register-ScheduledTask -TaskName "쿠팡업무_Mulder주간DB" -Action $act -Trigger $trg -Settings $set -Force | Out-Null
Get-ScheduledTask -TaskName "쿠팡업무_Mulder주간DB" | Select-Object TaskName, State
