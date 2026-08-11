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
# Every 15 minutes. Runs whether or not the user is logged on is NOT used on purpose:
# the payload drives a visible Chrome window, so it must run in the interactive session.
#
# ASCII-only on purpose: PowerShell 5.1 reads BOM-less UTF-8 as CP949 and mangles
# non-ASCII, which once killed a watcher without a single log line.
$ErrorActionPreference = 'Stop'
$root = "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
$name = 'CSOS_BrowserChain'
$py   = (Get-Command python).Source

$action  = New-ScheduledTaskAction -Execute $py `
             -Argument "band\browser_chain.py" -WorkingDirectory $root
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).Date.AddMinutes(7) `
             -RepetitionInterval (New-TimeSpan -Minutes 15)
# ExecutionTimeLimit is the backstop, not the plan: a tick should take under 3 minutes.
# StartWhenAvailable so a missed run (sleep, reboot) is picked up instead of skipped.
$set = New-ScheduledTaskSettingsSet -StartWhenAvailable `
         -ExecutionTimeLimit (New-TimeSpan -Minutes 20) `
         -MultipleInstances IgnoreNew -AllowStartIfOnBatteries `
         -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $name -Action $action -Trigger $trigger `
    -Settings $set -Description 'Band/ERP browser collection - one small step per tick' `
    -Force | Out-Null

Write-Output "registered: $name (every 15 min)"
Get-ScheduledTask -TaskName $name | Select-Object TaskName, State | Format-List
