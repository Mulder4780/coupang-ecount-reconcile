param(
    [switch]$Remove
)

$ErrorActionPreference = "Stop"
$TaskName = "CSOS_AutomationPipeline"

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
    -Argument "automation_pipeline.py --once --trigger scheduler" `
    -WorkingDirectory $Root

# Windows PowerShell 5.1 does not expose a writable Repetition object on the
# Daily trigger returned by New-ScheduledTaskTrigger.  Build the repetition in
# the cmdlet instead; ten years is effectively permanent and rerunning this
# installer safely replaces the task.  IgnoreNew plus the pipeline PID lock
# prevents overlap when one reconciliation takes longer.
$Trigger = New-ScheduledTaskTrigger `
    -Once `
    -At ((Get-Date).AddMinutes(1)) `
    -RepetitionInterval (New-TimeSpan -Minutes 5) `
    -RepetitionDuration (New-TimeSpan -Days 3650)
$Trigger.Repetition.StopAtDurationEnd = $false

$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 120)
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
    -Description "Every five minutes: detect Kakao/Band/ERP changes, reconcile objective evidence into the canonical app DB, and create a verified local Excel archive. Login is never bypassed." `
    -Force | Out-Null

$Task = Get-ScheduledTask -TaskName $TaskName
Write-Host "Registered scheduled task: $TaskName (every 5 minutes)"
