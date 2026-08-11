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
  [string]$Browser = 'chrome',
  [string]$OpenUrl = ''
)
$ErrorActionPreference = 'Stop'
$root = "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
Set-Location $root
Add-Type -AssemblyName System.Windows.Forms
# Force() exists because SetForegroundWindow alone is unreliable here (2026-08-12):
# Windows refuses foreground changes requested by a process that does not already own
# the foreground, and this script runs from a background shell and from the Task
# Scheduler. Measured: the same call succeeded at 00:23 and failed three times at
# 01:10, with no difference in the browser - the difference was who had focus. The
# documented way round is to attach to the foreground thread's input queue first, so
# the request comes from a thread that is allowed to make it.
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class IH{[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);[DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr h,int c);[DllImport("user32.dll")]public static extern bool BringWindowToTop(IntPtr h);[DllImport("user32.dll")]public static extern bool AttachThreadInput(uint a,uint b,bool f);[DllImport("kernel32.dll")]public static extern uint GetCurrentThreadId();public static void Force(IntPtr h){uint p=0;uint fg=GetWindowThreadProcessId(GetForegroundWindow(),out p);uint me=GetCurrentThreadId();bool at=(fg!=me)&&AttachThreadInput(me,fg,true);try{ShowWindow(h,9);BringWindowToTop(h);SetForegroundWindow(h);}finally{if(at){AttachThreadInput(me,fg,false);}}}[DllImport("user32.dll")]public static extern void keybd_event(byte vk,byte scan,uint flags,UIntPtr extra);public static void Type(string s){foreach(char c in s){keybd_event(0,(byte)c,4,UIntPtr.Zero);keybd_event(0,(byte)c,6,UIntPtr.Zero);System.Threading.Thread.Sleep(15);}}}'

$dl = Join-Path $env:USERPROFILE 'Downloads'

function Focus-Browser {
    for ($i = 0; $i -lt 5; $i++) {
        $c = Get-Process $Browser -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne '' } | Select-Object -First 1
        if (-not $c) { Start-Sleep -Seconds 2; continue }
        [IH]::Force($c.MainWindowHandle)
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
# Chromium's self-XSS guard swallows the FIRST paste into a fresh console until the
# human enters 'allow pasting'. Nothing errors and nothing downloads, so it looks
# exactly like 'DevTools never opened' - measured 2026-08-12 on Whale.
# ★ It must be entered with IH::Type (SendInput unicode), never SendKeys: SendKeys
#   goes through whatever IME is active and Korean is normal on this machine, so the
#   phrase arrived as jamo. Worse, when the console did not have focus those keystrokes
#   landed in the ADDRESS BAR and navigated the human's tab to a Naver search for
#   'aㅣㅣㅐㅈ ㅔㅁㄴ샤ㅜㅎ' - twice. Blind typing can move the page you meant to read,
#   which is why this runs only AFTER a ping has already failed.
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

# -OpenUrl removes the tab hunt entirely (2026-08-12). Cycling tabs sounds cheap but a
# tab whose console cannot answer stops the walk: measured this day, the ping's blob
# download never landed on a ChatGPT tab, so the band tab two positions further was
# unreachable. A fresh tab has no unknowns - DevTools closed by definition, active by
# definition - and the login rides along because it is the same profile. We are not
# logging in, we are only choosing which page the console belongs to.
#
# The url goes through the CLIPBOARD, never SendKeys. SendKeys types through whatever
# IME is active, and Korean is the normal state on this machine: measured 2026-08-12,
# 'https://www.band.us/band/90610953' arrived as 'ㅗㅅ센://ㅈㅈㅈ.ㅠ뭉.ㅕㄴ/...' and Whale
# ran a Naver search on the jamo. Nothing errored - we simply injected into the wrong
# page. Any literal text sent to a window must be pasted, not typed.
if ($OpenUrl) {
    SK '^t'; Start-Sleep -Milliseconds 900
    Set-Clipboard -Value $OpenUrl
    Start-Sleep -Milliseconds 400
    SK '^v'; Start-Sleep -Milliseconds 600
    SK '{ENTER}'; Start-Sleep -Seconds 8
    Write-Output ("opened: " + $OpenUrl)
}

SK '^+j'; Start-Sleep -Seconds 4
$h1 = Ping-Host
if (-not $h1) {
    [IH]::Type('allow pasting'); Start-Sleep -Milliseconds 400
    SK '{ENTER}'; Start-Sleep -Milliseconds 900
    $h1 = Ping-Host
}
if (-not $h1) { SK '^+j'; Start-Sleep -Seconds 4; $h1 = Ping-Host }
# Ctrl+Shift+J is not guaranteed. Naver Whale ships its own shortcut table, so fall
# back to F12 (opens DevTools on Elements) plus Escape (raises the console drawer).
# Do not assume a shortcut is universal just because the engine is Chromium.
if (-not $h1) {
    SK '{F12}'; Start-Sleep -Seconds 5
    SK '{ESC}'; Start-Sleep -Seconds 2
    $h1 = Ping-Host
}
# hostname 'devtools' means the console we reached belongs to the DevTools frontend
# itself - Ctrl+Shift+J landed while DevTools already had focus and opened
# DevTools-on-DevTools. Peeling one layer puts us back on the page. This is a wrong
# LAYER, not a wrong tab, so cycling tabs here would walk away from the right page.
$peel = 0
while ($h1 -eq 'devtools' -and $peel -lt 3) {
    $peel++
    Write-Output "peeling DevTools layer $peel"
    SK '^+j'; Start-Sleep -Seconds 3
    $h1 = Ping-Host
}
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
