param(
    [switch]$Remove
)

$ErrorActionPreference = "Stop"
# Keep this file ASCII-only (see install_ledger_schedule.ps1 for why).
# Task name: Coupang_AutomationScout in Korean.
$TaskName = -join @(
    [char]0xCFE0, [char]0xD321, [char]0xC5C5, [char]0xBB34, [char]0x005F,
    [char]0xC790, [char]0xB3D9, [char]0xD654, [char]0xD6C4, [char]0xBCF4,
    [char]0xAC80, [char]0xC0C9
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
    -Argument "auto_scout.py" `
    -WorkingDirectory $Root
$Triggers = @(
    (New-ScheduledTaskTrigger -Daily -At "10:00")
)
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
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
    -Description "Daily 10:00 GitHub/open-source automation candidate scan (auto_scout.py)." `
    -Force | Out-Null

$Task = Get-ScheduledTask -TaskName $TaskName
$Times = ($Task.Triggers | ForEach-Object { ([datetime]$_.StartBoundary).ToString("HH:mm") }) -join ", "
Write-Host "Registered scheduled task: $TaskName ($Times)"
