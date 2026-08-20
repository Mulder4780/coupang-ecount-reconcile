param(
    [switch]$Remove,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
# Keep this file ASCII-only (see install_ledger_schedule.ps1 for why:
# Windows PowerShell 5.1 reads BOM-less UTF-8 as CP949 and mangles Hangul).
# Task name: Coupang_RealtimeIssueWatch in Korean.
$TaskName = -join @(
    [char]0xCFE0, [char]0xD321, [char]0xC5C5, [char]0xBB34, [char]0x005F,
    [char]0xC2E4, [char]0xC2DC, [char]0xAC04, [char]0xBB38, [char]0xC81C,
    [char]0xAC10, [char]0xC2DC
)

if ($Remove) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task: $TaskName"
    exit 0
}

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Root) { $Root = (Get-Location).Path }
$Vbs = Join-Path $Root "run_hidden.vbs"
$Bat = Join-Path $Root "realtime_monitor.bat"
foreach ($p in @($Vbs, $Bat)) {
    if (-not (Test-Path -LiteralPath $p)) { throw "not found: $p" }
}

# wscript.exe + run_hidden.vbs, not the .bat directly: a .bat opens a console
# window, and run_hidden.vbs hides it while returning the child's exit code
# unchanged (WScript.Quit sh.Run(line, 0, True)).  Swallowing that code would
# blind schedule_watch - hiding the window must not hide the failure.
$Action = New-ScheduledTaskAction `
    -Execute "wscript.exe" `
    -Argument ('"' + $Vbs + '" "' + $Bat + '"') `
    -WorkingDirectory $Root

# Every 4 hours from 01:43:30.  StartWhenAvailable is deliberately NOT set:
# a missed run of a *monitor* should be skipped, not fired late in the middle
# of someone's working hours.
$Repeat = New-ScheduledTaskTrigger `
    -Once `
    -At "01:43:30" `
    -RepetitionInterval (New-TimeSpan -Hours 4) `
    -RepetitionDuration (New-TimeSpan -Days 3650)
$Repeat.Repetition.StopAtDurationEnd = $false
$Triggers = @($Repeat)

# Battery flags are mandatory for every round in this project: the default
# settings set is "do not start on battery + stop when switching to battery",
# and a task stopped that way is recorded as Queued, not Failed - so it leaves
# no trace at all (2026-08-12 incident).
$Settings = New-ScheduledTaskSettingsSet `
    -MultipleInstances IgnoreNew `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 4)
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
    -Description "Every 4 hours - realtime problem watch: reads the latest ledger fingerprint and open issues, writes reports (realtime_monitor.bat, output redirected to reports)." `
    -Force | Out-Null

$Task = Get-ScheduledTask -TaskName $TaskName
Write-Host "Registered scheduled task: $TaskName (every 4 hours from 01:43:30)"
