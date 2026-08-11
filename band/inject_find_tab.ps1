# inject_find_tab.ps1 -Js <payload> -ExpectHost <substr> [-MaxTabs 8]
#
# Cycle Chrome tabs (Ctrl+Tab) until the console proves it belongs to $ExpectHost,
# then paste the payload there.
#
# Why this exists (2026-08-11):
#   inject_here.ps1 pastes into whatever tab is active and REFUSES when the ping says
#   the console belongs to another origin - correct, but it leaves the caller stuck:
#   the ERP tour parks Chrome on ecount, so a band injection could never start. Tab
#   titles cannot be used to pick the tab - MainWindowTitle lags by seconds and once
#   reported 2 tabs where 4 existed. So we do not guess from titles at all: we let the
#   ping (whose download filename carries location.hostname) be the judge, exactly as
#   inject_here.ps1 already does. Cycling costs one ping per tab, which is cheap and,
#   unlike a title read, cannot be wrong.
#
# Any www.band.us tab will do for the collector: grab_posts.js builds an ABSOLUTE url
# (https://www.band.us/band/<band>/post/<no>), so it does not matter which band the tab
# is showing. It only has to be logged in and, once running, the ACTIVE tab - band
# paints post bodies with rAF, which never runs in a hidden tab.
#
# ASCII-only on purpose: PowerShell 5.1 reads BOM-less UTF-8 as CP949 and mangles
# non-ASCII, which once killed a watcher without a single log line.
param(
  [Parameter(Mandatory=$true)][string]$Js,
  [Parameter(Mandatory=$true)][string]$ExpectHost,
  [int]$MaxTabs = 8
)
$ErrorActionPreference = 'Stop'
$root = "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class FT{[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr h,int c);}'

function Say($m) { Write-Output ((Get-Date -Format 'HH:mm:ss') + ' ' + $m) }

if (-not (Test-Path $Js)) { Say "ABORT: no payload $Js"; exit 1 }

for ($i = 1; $i -le $MaxTabs; $i++) {
    Say "tab attempt $i/$MaxTabs"
    $out = & (Join-Path $root 'band\inject_here.ps1') -Js $Js -ExpectHost $ExpectHost 2>&1
    $txt = ($out | Out-String)
    $out | ForEach-Object { Say ("  | " + $_) }

    if ($txt -match 'INJECTED on') { Say "OK: pasted on attempt $i"; exit 0 }

    # 'FAIL: console belongs to X' is the only error worth advancing a tab for.
    # Anything else (no Chrome, console unreachable, missing file) repeats forever.
    if ($txt -notmatch 'console belongs to') {
        Say 'ABORT: not a wrong-tab failure - stopping so the real cause stays visible'
        exit 2
    }

    # Advance one tab. Re-focus first: the ping's download may have moved focus.
    $c = Get-Process chrome -ErrorAction SilentlyContinue |
         Where-Object { $_.MainWindowTitle -ne '' } | Select-Object -First 1
    if (-not $c) { Say 'ABORT: Chrome window vanished'; exit 1 }
    [FT]::ShowWindow($c.MainWindowHandle, 9) | Out-Null
    [FT]::SetForegroundWindow($c.MainWindowHandle) | Out-Null
    Start-Sleep -Milliseconds 700
    # Close DevTools before switching, otherwise Ctrl+Tab may be eaten by the DevTools
    # window and we would ping the SAME tab again for every remaining attempt.
    [System.Windows.Forms.SendKeys]::SendWait('^+j')
    Start-Sleep -Milliseconds 900
    [System.Windows.Forms.SendKeys]::SendWait('^{TAB}')
    Start-Sleep -Seconds 2
}
Say "FAIL: no tab on '$ExpectHost' after $MaxTabs attempts"
exit 3
