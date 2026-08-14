param(
    [switch]$Remove
)

$ErrorActionPreference = "Stop"
$TaskName = "CSOS_AppServerGuard"
$RunName = "CSOS_AppServerGuard"
$RunKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"

if ($Remove) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $RunKey -Name $RunName -ErrorAction SilentlyContinue
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
    -Argument "webapp\server_guard.py" `
    -WorkingDirectory $Root

# The guard normally stays alive.  The five-minute trigger is a second safety net:
# if the guard process itself disappears, the next trigger starts it again.  IgnoreNew
# and the guard's local singleton socket prevent duplicate supervisors.
$Triggers = @(
    (New-ScheduledTaskTrigger -AtLogOn),
    (New-ScheduledTaskTrigger -Once -At ((Get-Date).AddSeconds(20)) `
        -RepetitionInterval (New-TimeSpan -Minutes 5) `
        -RepetitionDuration (New-TimeSpan -Days 3650))
)
$Triggers[1].Repetition.StopAtDurationEnd = $false

$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1)
$Principal = New-ScheduledTaskPrincipal `
    -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel Limited

try {
    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $Action `
        -Trigger $Triggers `
        -Settings $Settings `
        -Principal $Principal `
        -Description "Always-on lightweight guard for the CSOS app origin and tunnel supervisor. Restarts after three failed health checks and restarts itself if the guard exits." `
        -Force | Out-Null
    Remove-ItemProperty -Path $RunKey -Name $RunName -ErrorAction SilentlyContinue
    Start-ScheduledTask -TaskName $TaskName
    Write-Host "Registered and started scheduled task: $TaskName"
} catch {
    $RegisterError = $_.Exception.Message
    $Guard = Join-Path $Root "webapp\server_guard.py"
    $RunCommand = '"{0}" "{1}"' -f $Pythonw, $Guard
    try {
        # Register-ScheduledTask is denied on some managed PCs even though the
        # current user may create an interactive task through schtasks.exe.
        # Keep the five-minute task and HKCU logon start together: the singleton
        # socket makes duplicate launches harmless, while either path can revive
        # the manager if the other one is unavailable.
        & schtasks.exe /Create /TN $TaskName /TR $RunCommand /SC MINUTE /MO 5 /F /IT /RL LIMITED | Out-Null
        if ($LASTEXITCODE -ne 0) {
            throw "schtasks.exe exit $LASTEXITCODE"
        }
        $FallbackSettings = New-ScheduledTaskSettingsSet `
            -StartWhenAvailable `
            -MultipleInstances IgnoreNew `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit ([TimeSpan]::Zero) `
            -RestartCount 3 `
            -RestartInterval (New-TimeSpan -Minutes 1)
        Set-ScheduledTask -TaskName $TaskName -Settings $FallbackSettings | Out-Null
        New-Item -Path $RunKey -Force | Out-Null
        New-ItemProperty -Path $RunKey -Name $RunName -Value $RunCommand `
            -PropertyType String -Force | Out-Null
        Start-ScheduledTask -TaskName $TaskName
        Write-Warning "Primary registration denied ($RegisterError). Installed five-minute task + logon startup: $TaskName"
    } catch {
        # Last fallback needs no Task Scheduler permission. tunnel_run.py and the
        # 30-minute watchdog independently check the guard heartbeat as well.
        New-Item -Path $RunKey -Force | Out-Null
        New-ItemProperty -Path $RunKey -Name $RunName -Value $RunCommand `
            -PropertyType String -Force | Out-Null
        Start-Process -FilePath $Pythonw -ArgumentList @($Guard) -WorkingDirectory $Root -WindowStyle Hidden
        Write-Warning "Task registration failed. Installed per-user logon startup instead: $RunName"
    }
}
