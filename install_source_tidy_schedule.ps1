param(
    [switch]$Remove
)

$ErrorActionPreference = "Stop"
# Keep this file ASCII-only (see install_ledger_schedule.ps1 for why).
#
# Why this file exists (2026-08-12, verification [228]):
#   This task has been running since 2026-07-29 but had NO installer in the repo,
#   so rebuilding the machine would silently lose it and schedule_watch.declared()
#   could not tell that it MUST exist. Settings below were exported from the live
#   task, not invented.
#
# NOTE: ExecutionTimeLimit is PT3H and this task is force-killed (0xC000013A)
#   daily, same as the 09:50 round. Work item [38] owns that fix; this file is the
#   single declared place to change the limit.
#
# Task name: "Coupang work_Source tidy" in Korean.
$TaskName = -join @(
    [char]0xCFE0, [char]0xD321, [char]0xC5C5, [char]0xBB34, [char]0x005F,
    [char]0xC6D0, [char]0xBCF8, [char]0xC790, [char]0xB8CC,
    [char]0xC790, [char]0xB3D9, [char]0xC815, [char]0xB9AC
)

if ($Remove) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task: $TaskName"
    exit 0
}

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
# The batch file name is Korean too, so build it from code points.
$BatName = (-join @(
    [char]0xC6D0, [char]0xBCF8, [char]0xC790, [char]0xB8CC,
    [char]0xC790, [char]0xB3D9, [char]0xC815, [char]0xB9AC
)) + ".bat"
$Bat = Join-Path $Root $BatName
if (-not (Test-Path -LiteralPath $Bat)) {
    throw "source tidy batch not found: $Bat"
}

# Windowless: see install_daily_schedule.ps1 for why. run_hidden.vbs keeps the
# child's exit code so schedule_watch still knows whether the round really ran.
$Hide = Join-Path $Root "run_hidden.vbs"
if (-not (Test-Path -LiteralPath $Hide)) {
    throw "run_hidden.vbs not found: $Hide"
}
$Action = New-ScheduledTaskAction -Execute "wscript.exe" `
    -Argument ('"{0}" "{1}"' -f $Hide, $Bat) -WorkingDirectory $Root
$Triggers = @(
    (New-ScheduledTaskTrigger -Daily -At "09:35")
)
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
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
    -Description "Daily 09:35 - tidy the single source-of-record folder on Z: (source_tidy) before the 09:50 reconcile round reads it. See ecount/CLAUDE.md." `
    -Force | Out-Null

$Task = Get-ScheduledTask -TaskName $TaskName
$Times = ($Task.Triggers | ForEach-Object { ([datetime]$_.StartBoundary).ToString("HH:mm") }) -join ", "
Write-Host "Registered scheduled task: $TaskName ($Times daily)"
