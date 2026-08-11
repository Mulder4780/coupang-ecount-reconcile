# collect_step.ps1 -Url <band url> -Js <collector js>
#
# Deterministic single collection step. Three sins were paid for on 2026-08-11:
#   1. Ctrl+Shift+J is a TOGGLE - with DevTools already open it closed the console and
#      the 26KB collector was pasted into the PAGE. No error, looked like success.
#   2. Focus alone proves nothing: the active tab turned out to be the app dashboard
#      ("Mulder Control Hub"), so a "verified" injection ran on the wrong origin.
#   3. band pauses itself when the tab is hidden (rAF never fires), so a step can sit
#      silent for 25 minutes and look identical to a step that never started.
#
# So this script proves every assumption instead of trusting it:
#   * always opens a FRESH tab (a new tab always has DevTools closed -> the toggle is
#     deterministic, and the new tab is by definition the active one)
#   * the ping encodes location.hostname in its download filename, so we can read back
#     WHICH page the console belongs to before pasting anything large
#   * refuses to paste the collector unless the ping came from band.us
param(
  [Parameter(Mandatory=$true)][string]$Url,
  [Parameter(Mandatory=$true)][string]$Js,
  [string]$ExpectHost = 'band.us'
)
$ErrorActionPreference = 'Stop'
$root = "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
Set-Location $root
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class CS{[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);[DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr h,int c);}'

$dl = Join-Path $env:USERPROFILE 'Downloads'

function Focus-Chrome {
    for ($i = 0; $i -lt 4; $i++) {
        $c = Get-Process chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne '' } | Select-Object -First 1
        if (-not $c) { Start-Sleep -Seconds 2; continue }
        [CS]::ShowWindow($c.MainWindowHandle, 9) | Out-Null
        [CS]::SetForegroundWindow($c.MainWindowHandle) | Out-Null
        Start-Sleep -Milliseconds 900
        $h = [CS]::GetForegroundWindow(); $fp = 0
        [CS]::GetWindowThreadProcessId($h, [ref]$fp) | Out-Null
        $p = Get-Process -Id $fp -ErrorAction SilentlyContinue
        if ($p -and $p.ProcessName -eq 'chrome') { return $c }
        Start-Sleep -Seconds 2
    }
    return $null
}
function SK($k) { [System.Windows.Forms.SendKeys]::SendWait($k) }
function Paste($path) {
    Get-Content -Raw -Encoding UTF8 $path | Set-Clipboard
    Start-Sleep -Milliseconds 400
    SK '^v'; Start-Sleep -Milliseconds 2200
    SK '{ENTER}'
}
# returns the hostname the console belongs to, or $null
function Ping-Host {
    Get-ChildItem $dl -Filter '__ping__*.txt' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    Paste (Join-Path $root 'band\console_ping.js')
    for ($i = 0; $i -lt 14; $i++) {
        Start-Sleep -Milliseconds 700
        $f = Get-ChildItem $dl -Filter '__ping__*.txt' -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($f) {
            $name = $f.Name
            Remove-Item $f.FullName -Force -ErrorAction SilentlyContinue
            if ($name -match '^__ping__([^_]+(?:\.[^_]+)*)__') { return $Matches[1] }
            return 'unknown'
        }
    }
    return $null
}

$win = Focus-Chrome
if (-not $win) { Write-Output 'ABORT: cannot focus Chrome'; exit 1 }
if (-not (Test-Path $Js)) { Write-Output "ABORT: no such file $Js"; exit 1 }

# fresh tab -> active by definition, DevTools closed by definition
Set-Clipboard $Url
SK '^t'; Start-Sleep -Milliseconds 900
SK '^l'; Start-Sleep -Milliseconds 400
SK '^v'; Start-Sleep -Milliseconds 400
SK '{ENTER}'
Start-Sleep -Seconds 12
if (-not (Focus-Chrome)) { Write-Output 'ABORT: lost Chrome after navigation'; exit 1 }

SK '^+j'; Start-Sleep -Seconds 4
$h1 = Ping-Host
if (-not $h1) {
    # console was not reachable - the toggle may have closed an already-open DevTools
    SK '^+j'; Start-Sleep -Seconds 4
    $h1 = Ping-Host
}
if (-not $h1) { Write-Output 'FAIL: console not reachable'; exit 2 }
if ($h1 -notlike "*$ExpectHost*") {
    Write-Output "FAIL: console belongs to '$h1', expected '$ExpectHost' - refusing to paste collector"
    exit 3
}

Paste $Js
Write-Output ("INJECTED on " + $h1 + " : " + $Js + " " + (Get-Date -Format 'HH:mm:ss'))
