# install_browser_chain_schedule.ps1 - register the unattended browser-collection round.
#
# Why (2026-08-11, user: "I want this done without my hands"):
#   Keyboard injection already needs no human. What was missing was somebody to START
#   THE NEXT STEP: when the band backfill finishes at ~00:11 nothing else runs until a
#   person shows up. A chat session cannot be that somebody - it ends. Only files and
#   the Task Scheduler survive.
#
# The round itself is deliberately small: one tick starts at most ONE job and exits.
# It never babysits - a step that holds on for hours is exactly the accident recorded
# in incidents [175] and [180].
#
# Once a day at 12:00.  Fifteen-minute retries repeatedly stole the operator's focus
# and typed into Whale/other sites.  A second run is allowed only through the explicit
# `browser_chain.py --manual <target>` command after the user asks for it.
#
# ASCII-only on purpose: PowerShell 5.1 reads BOM-less UTF-8 as CP949 and mangles
# non-ASCII, which once killed a watcher without a single log line.
$ErrorActionPreference = 'Stop'
$root = "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
$name = 'CSOS_BrowserChain'
# pythonw.exe, not python.exe: even one console window interrupts the operator.
# browser_chain.py already guards its
# sys.stdout.reconfigure in try/except, so a None stdout is safe here ([235]).
# Fall back to python.exe if pythonw is missing - a window beats a dead task.
$py = Join-Path (Split-Path -Parent (Get-Command python).Source) 'pythonw.exe'
if (-not (Test-Path -LiteralPath $py)) {
    $py = (Get-Command python).Source
}

$action  = New-ScheduledTaskAction -Execute $py `
             -Argument "band\browser_chain.py" -WorkingDirectory $root
$trigger = New-ScheduledTaskTrigger -Daily -At '12:00'
# Do not use StartWhenAvailable.  If the PC wakes while the operator is working later,
# a missed browser round must not suddenly touch Chrome outside the agreed 12:00 slot.
$set = New-ScheduledTaskSettingsSet `
         -ExecutionTimeLimit (New-TimeSpan -Minutes 20) `
         -MultipleInstances IgnoreNew -AllowStartIfOnBatteries `
         -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $name -Action $action -Trigger $trigger `
    -Settings $set -Description 'Chrome allowlist collection - daily 12:00 or explicit manual command' `
    -Force | Out-Null

Write-Output "registered: $name (daily 12:00; extra runs are manual only)"
Get-ScheduledTask -TaskName $name | Select-Object TaskName, State | Format-List
