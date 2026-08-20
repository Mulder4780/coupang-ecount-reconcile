param(
    [switch]$Remove,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
# Keep this file ASCII-only (see install_ledger_schedule.ps1 for why:
# Windows PowerShell 5.1 reads BOM-less UTF-8 as CP949 and mangles Hangul).
# Task name: Coupang_Watchdog in Korean.
$TaskName = -join @(
    [char]0xCFE0, [char]0xD321, [char]0xC5C5, [char]0xBB34, [char]0x005F,
    [char]0xC6CC, [char]0xCE58, [char]0xB3C5
)

if ($Remove) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task: $TaskName"
    exit 0
}

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Root) { $Root = (Get-Location).Path }
$Python = (Get-Command python.exe -ErrorAction Stop).Source
$Pythonw = Join-Path (Split-Path -Parent $Python) "pythonw.exe"
if (-not (Test-Path -LiteralPath $Pythonw)) {
    throw "pythonw.exe not found: $Pythonw"
}

# pythonw.exe, not python.exe: a console window every 30 minutes is pure
# noise for the person using this laptop (see the 2026-08-13 instruction).
$Action = New-ScheduledTaskAction `
    -Execute $Pythonw `
    -Argument ('"' + (Join-Path $Root "watchdog.py") + '"') `
    -WorkingDirectory $Root

# Every 30 minutes from 10:57, plus once at logon.  The :27/:57 offset is
# deliberate - it keeps this round off the :00/:30 pile-up.
$Repeat = New-ScheduledTaskTrigger `
    -Once `
    -At "10:57" `
    -RepetitionInterval (New-TimeSpan -Minutes 30) `
    -RepetitionDuration (New-TimeSpan -Days 3650)
$Repeat.Repetition.StopAtDurationEnd = $false
$Triggers = @($Repeat, (New-ScheduledTaskTrigger -AtLogOn))

# Battery flags are mandatory for every round in this project: the default
# settings set is "do not start on battery + stop when switching to battery",
# and a task stopped that way is recorded as Queued, not Failed - so it leaves
# no trace at all (2026-08-12 incident).
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 72)
$Principal = New-ScheduledTaskPrincipal `
    -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel Limited

if ($DryRun) {
    # Print what would be registered so this installer can be checked without
    # touching the live scheduler.  Re-registering a running round is not free.
    Write-Host "TaskName=$TaskName"
    foreach ($a in @($Action)) {
        Write-Host "ACT exe=[$($a.Execute)] args=[$($a.Arguments)] cwd=[$($a.WorkingDirectory)]"
    }
    foreach ($g in $Triggers) {
        Write-Host "TRG type=$($g.CimClass.CimClassName) start=$($g.StartBoundary) rep=$($g.Repetition.Interval)"
    }
    Write-Host "SET limit=$($Settings.ExecutionTimeLimit) multi=$($Settings.MultipleInstances) swa=$($Settings.StartWhenAvailable) batt_stop=$($Settings.StopIfGoingOnBatteries) batt_nostart=$($Settings.DisallowStartIfOnBatteries)"
    Write-Host "PRIN user=$($Principal.UserId) logon=$($Principal.LogonType) runlevel=$($Principal.RunLevel)"
    exit 0
}

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Triggers `
    -Settings $Settings `
    -Principal $Principal `
    -Description "Every 30 minutes (and at logon) - self-healing round: automation pipeline, stale server/tunnel repair, schedule watch, parked-work resume, handoff snapshot (watchdog.py)." `
    -Force | Out-Null

$Task = Get-ScheduledTask -TaskName $TaskName
Write-Host "Registered scheduled task: $TaskName (every 30 minutes from 10:57, and at logon)"
