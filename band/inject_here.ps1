# inject_here.ps1 -Js <collector js> -ExpectHost <substring>
#
# Inject into the CURRENT active tab (no fresh tab). Needed for ecount ERP, whose
# session lives in the tab the human logged into - we cannot re-open it by URL.
#
# collect_step.ps1 removes two unknowns by opening a fresh tab (DevTools closed by
# definition, active by definition). Here we cannot, so we PROVE both instead:
#   * ping -> marker download proves the console is actually live
#   * the ping filename carries location.hostname, so we learn WHICH page owns that
#     console and refuse to paste the payload anywhere else
#   * if the first ping is silent, Ctrl+Shift+J is toggled once and we prove again
#     (the toggle trap: with DevTools already open it CLOSES the console and a large
#     paste silently lands in the page - incident #34)
#
# ASCII-only on purpose: PowerShell 5.1 reads BOM-less UTF-8 as CP949 and mangles
# any non-ASCII text, which once killed a watcher without a single log line.
# -Browser picks WHICH Chromium process to drive (2026-08-12). It used to be hardcoded
# to 'chrome', which is a guess about where the human logged in - and the guess was
# wrong: band/ecount were open in Naver Whale (band.us is a Naver service, so this is
# the normal case, not the odd one) while the only Chrome window was an unrelated app.
# The injector focused that Chrome, could not open a console there, and reported a
# cause that had nothing to do with the real one. Never name the browser in code.
param(
  [Parameter(Mandatory=$true)][string]$Js,
  [Parameter(Mandatory=$true)][string]$ExpectHost,
  [string]$Browser = 'chrome'
)
$ErrorActionPreference = 'Stop'
$root = "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
Set-Location $root
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class IH{[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);[DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr h,int c);}'

$dl = Join-Path $env:USERPROFILE 'Downloads'

function Focus-Browser {
    for ($i = 0; $i -lt 5; $i++) {
        $c = Get-Process $Browser -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne '' } | Select-Object -First 1
        if (-not $c) { Start-Sleep -Seconds 2; continue }
        [IH]::ShowWindow($c.MainWindowHandle, 9) | Out-Null
        [IH]::SetForegroundWindow($c.MainWindowHandle) | Out-Null
        Start-Sleep -Milliseconds 900
        $h = [IH]::GetForegroundWindow(); $fp = 0
        [IH]::GetWindowThreadProcessId($h, [ref]$fp) | Out-Null
        $p = Get-Process -Id $fp -ErrorAction SilentlyContinue
        if ($p -and $p.ProcessName -eq $Browser) { return $c }
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

$win = Focus-Browser
if (-not $win) { Write-Output "ABORT: cannot focus browser '$Browser'"; exit 1 }
if (-not (Test-Path $Js)) { Write-Output "ABORT: no such file $Js"; exit 1 }
Write-Output ("window title: " + $win.MainWindowTitle)

SK '^+j'; Start-Sleep -Seconds 4
$h1 = Ping-Host
if (-not $h1) { SK '^+j'; Start-Sleep -Seconds 4; $h1 = Ping-Host }
if (-not $h1) { Write-Output 'FAIL: console not reachable (tried both DevTools states)'; exit 2 }
if ($h1 -notlike "*$ExpectHost*") {
    Write-Output "FAIL: console belongs to '$h1', expected '*$ExpectHost*' - refusing to paste payload"
    exit 3
}

Paste $Js
Write-Output ("INJECTED on " + $h1 + " : " + $Js + " " + (Get-Date -Format 'HH:mm:ss'))
# ★ Say success out loud. Without this the caller inherits whatever $LASTEXITCODE
#   happened to hold (empty on a fresh shell), and a caller that tests it reads a
#   successful injection as a failure - measured 2026-08-11, it aborted the
#   liveness probe one second after the payload had actually landed.
exit 0
