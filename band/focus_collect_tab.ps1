# focus_collect_tab.ps1 - bring the collector tab back to the front of its Chrome window.
#
# Why (2026-08-11): band paints post bodies with requestAnimationFrame, which never runs
# in a hidden tab, so grab_posts.js pauses itself. "Hidden" includes *the window being
# visible but a different tab being active* - measured tonight: the injection landed on
# www.band.us (proved by the ping) and then collection stalled because the window's
# active tab had moved back to the company page.
#
# collect_step.ps1 always opens the collector in a FRESH tab, so it is the last tab of
# that window: Ctrl+9 jumps there deterministically (Chrome maps Ctrl+9 to "last tab",
# not "tab 9"). Then pin the window small + topmost so nothing can occlude it again.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;using System.Text;public class FC{[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);[DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr h,int c);[DllImport("user32.dll")]public static extern int GetWindowText(IntPtr h,StringBuilder s,int n);}'

$c = Get-Process chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne '' } | Select-Object -First 1
if (-not $c) { Write-Output 'ABORT: no Chrome window'; exit 1 }
[FC]::ShowWindow($c.MainWindowHandle, 9) | Out-Null
[FC]::SetForegroundWindow($c.MainWindowHandle) | Out-Null
Start-Sleep -Milliseconds 900

$h = [FC]::GetForegroundWindow(); $fp = 0
[FC]::GetWindowThreadProcessId($h, [ref]$fp) | Out-Null
$p = Get-Process -Id $fp -ErrorAction SilentlyContinue
if (-not $p -or $p.ProcessName -ne 'chrome') { Write-Output 'ABORT: cannot focus Chrome'; exit 1 }

[System.Windows.Forms.SendKeys]::SendWait('^9')   # last tab = the collector tab
Start-Sleep -Seconds 2

$t = (Get-Process -Id $c.Id).MainWindowTitle
Write-Output "active tab now: $t"
if ($t -notmatch '밴드|band') { Write-Output 'WARN: last tab does not look like band - collector tab may have been closed' }
