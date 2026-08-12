param(
    [switch]$Remove
)

$ErrorActionPreference = "Stop"
# Keep this file ASCII-only (see install_ledger_schedule.ps1 for why).
#
# Why this file exists (2026-08-12, verification [228]):
#   This task has been running since 2026-07-24 but had NO installer in the repo.
#   That means (a) rebuilding the machine silently loses the project's main daily
#   round, and (b) schedule_watch.declared() -- which reads install_*.ps1 to learn
#   which rounds MUST exist -- could not see it, so its disappearance would raise
#   no alarm. The settings below were exported from the live task, not invented.
#
# NOTE (do not "fix" silently): the live task has StartWhenAvailable = False, so a
#   day when the PC is asleep at 09:50 is skipped with no trace. That is reproduced
#   here on purpose -- changing it is a behaviour decision for a human. The missed
#   day is now visible anyway: schedule_watch reports "not run" when a scheduled
#   time has passed with no run record.
#
# NOTE: ExecutionTimeLimit is PT3H while the round measured 292.3 min on
#   2026-08-11, so it is force-killed (0xC000013A) every day. Work item [38] owns
#   that fix; this file is the single declared place to change the limit.
#
# Task name: "Coupang work_Daily auto reconcile" in Korean.
$TaskName = -join @(
    [char]0xCFE0, [char]0xD321, [char]0xC5C5, [char]0xBB34, [char]0x005F,
    [char]0xC77C, [char]0xC77C, [char]0xC790, [char]0xB3D9, [char]0xB300, [char]0xC870
)

if ($Remove) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task: $TaskName"
    exit 0
}

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Bat = Join-Path $Root "daily_run.bat"
if (-not (Test-Path -LiteralPath $Bat)) {
    throw "daily_run.bat not found: $Bat"
}

$Action = New-ScheduledTaskAction -Execute $Bat -WorkingDirectory $Root
$Triggers = @(
    (New-ScheduledTaskTrigger -Daily -At "09:50")
)
$Settings = New-ScheduledTaskSettingsSet `
    -MultipleInstances IgnoreNew `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 3)
$Principal = New-ScheduledTaskPrincipal `
    -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel Limited

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Triggers `
    -Settings $Settings `
    -Principal $Principal `
    -Description "Daily 09:50 - the project's main reconcile round (daily_run.py): absorb downloads, reconcile every source against the ledger, judge objective completion, queue results. See ecount/CLAUDE.md." `
    -Force | Out-Null

$Task = Get-ScheduledTask -TaskName $TaskName
$Times = ($Task.Triggers | ForEach-Object { ([datetime]$_.StartBoundary).ToString("HH:mm") }) -join ", "
Write-Host "Registered scheduled task: $TaskName ($Times daily)"
