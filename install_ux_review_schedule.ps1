param(
    [switch]$Remove
)

$ErrorActionPreference = "Stop"
# Keep this file ASCII-only (see install_ledger_schedule.ps1 for why).
# Task name: Coupang_UXReview in Korean.
$TaskName = -join @(
    [char]0xCFE0, [char]0xD321, [char]0xC5C5, [char]0xBB34, [char]0x005F,
    [char]0xC5C5, [char]0xBB34, [char]0xC13C, [char]0xD130, [char]0x0055,
    [char]0x0058, [char]0xC810, [char]0xAC80
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
    -Argument "ux_review.py" `
    -WorkingDirectory $Root
# Every 3 days at 12:00 (user instruction 2026-08-05).
$Triggers = @(
    (New-ScheduledTaskTrigger -Daily -DaysInterval 3 -At "12:00")
)
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 15)
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
    -Description "Every 3 days at 12:00 - analyse app usage records and write UX improvement suggestions (ux_review.py, read-only)." `
    -Force | Out-Null

$Task = Get-ScheduledTask -TaskName $TaskName
$Times = ($Task.Triggers | ForEach-Object { ([datetime]$_.StartBoundary).ToString("HH:mm") }) -join ", "
Write-Host "Registered scheduled task: $TaskName ($Times, every 3 days)"
