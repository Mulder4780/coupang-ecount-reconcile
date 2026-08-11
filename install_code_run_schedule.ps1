param(
    [switch]$Remove
)

$ErrorActionPreference = "Stop"
# Keep this file ASCII-only (see install_ledger_schedule.ps1 for why).
# Task name: "Coupang work_Coding round" in Korean.
$TaskName = -join @(
    [char]0xCFE0, [char]0xD321, [char]0xC5C5, [char]0xBB34, [char]0x005F,
    [char]0xCF54, [char]0xB529, [char]0xD68C, [char]0xCC28
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
    -Argument "daily_code_run.py --run" `
    -WorkingDirectory $Root

# Daily 12:00 (user instruction 2026-08-11: once a day between 12:00 and 13:00).
# The slot is deliberately empty of other rounds: 09:50 daily_run and the 11:00 /
# 15:00 archive rounds sit on either side, so this one does not fight them for
# the Z: (SMB) share. daily_code_run.py itself refuses to run outside 12:00-13:00,
# steps aside when another session holds 'code' or a live window is open, and
# never counts stepping aside as a failure.
$Triggers = @(
    (New-ScheduledTaskTrigger -Daily -At "12:00")
)
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 50)
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
    -Description "Daily 12:00-13:00 - unattended coding-session round: run the synthetic check, confirm auto-compact wiring is alive, refresh the session handoff page, and raise only red results. Steps aside (exit 0) when another session holds the code claim, when a live session window is open, or when the daily reconciliation round is still running. Never collects data and never writes the ledger (daily_code_run.py)." `
    -Force | Out-Null

$Task = Get-ScheduledTask -TaskName $TaskName
$Times = ($Task.Triggers | ForEach-Object { ([datetime]$_.StartBoundary).ToString("HH:mm") }) -join ", "
Write-Host "Registered scheduled task: $TaskName ($Times daily)"
