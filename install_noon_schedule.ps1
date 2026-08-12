param(
    [switch]$Remove
)

$ErrorActionPreference = "Stop"
# Keep this file ASCII-only (see install_ledger_schedule.ps1 for why).
# Task name: Coupang_NoonRound in Korean (쿠팡업무_정오회차).
$TaskName = -join @(
    [char]0xCFE0, [char]0xD321, [char]0xC5C5, [char]0xBB34, [char]0x005F,
    [char]0xC815, [char]0xC624, [char]0xD68C, [char]0xCC28
)

if ($Remove) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task: $TaskName"
    exit 0
}

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = (Get-Command python.exe -ErrorAction Stop).Source
$Pythonw = Join-Path (Split-Path -Parent $Python) "pythonw.exe"
if (-not (Test-Path -LiteralPath $Pythonw)) {
    throw "pythonw.exe not found: $Pythonw"
}

$Action = New-ScheduledTaskAction `
    -Execute $Pythonw `
    -Argument "noon_run.py" `
    -WorkingDirectory $Root

# Daily at 12:00, retried every 10 minutes for 55 minutes.
# The repetition is NOT "run six times" - noon_run.py runs at most once per day
# (marker file) and simply yields to other sessions/rounds, so each retry is a
# fresh chance to find the machine free. The duration stops before 13:00 so no
# attempt ever falls outside the window the user asked for.
$Trigger = New-ScheduledTaskTrigger -Daily -At "12:00"
$Trigger.Repetition = (New-ScheduledTaskTrigger -Once -At "12:00" `
    -RepetitionInterval (New-TimeSpan -Minutes 10) `
    -RepetitionDuration (New-TimeSpan -Minutes 55)).Repetition

$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
$Principal = New-ScheduledTaskPrincipal `
    -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel Limited

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Principal $Principal `
    -Description "Daily 12:00-13:00 - one conflict-aware round (noon_run.py): handoff refresh, synthetic check, PM content survey, camp-name standard diff. Yields to live sessions and running rounds; never writes Excel, never collects, never commits." `
    -Force | Out-Null

$Task = Get-ScheduledTask -TaskName $TaskName
$Start = ([datetime]$Task.Triggers[0].StartBoundary).ToString("HH:mm")
Write-Host "Registered scheduled task: $TaskName (daily $Start, retry every 10 min for 55 min)"
