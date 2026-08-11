# inject_verified.ps1 -Js <path>
#
# Why this exists (2026-08-11): Ctrl+Shift+J is a TOGGLE. When DevTools is already
# open it CLOSES it, and the 26KB collector then gets pasted into the PAGE, not the
# console. Nothing runs, no error is shown, and the injection looks successful.
# Step 1 worked at 15:52 and step 2 died at 15:55 for exactly this reason.
#
# So we never guess the DevTools state. We prove the console is live with a tiny
# "ping" script that downloads a marker file, and only then paste the real payload.
# A tiny ping is also safe if it lands on the page by mistake; a 26KB paste is not.
param([Parameter(Mandatory=$true)][string]$Js)
$ErrorActionPreference = 'Stop'
$root = "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
Set-Location $root
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class WV{[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);[DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr h,int c);}'

$PING = Join-Path $env:USERPROFILE 'Downloads\__console_ping.txt'

function Focus-Chrome {
    for ($i = 0; $i -lt 4; $i++) {
        $c = Get-Process chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne '' } | Select-Object -First 1
        if (-not $c) { Start-Sleep -Seconds 2; continue }
        [WV]::ShowWindow($c.MainWindowHandle, 9) | Out-Null
        [WV]::SetForegroundWindow($c.MainWindowHandle) | Out-Null
        Start-Sleep -Milliseconds 900
        $h = [WV]::GetForegroundWindow(); $fp = 0
        [WV]::GetWindowThreadProcessId($h, [ref]$fp) | Out-Null
        $p = Get-Process -Id $fp -ErrorAction SilentlyContinue
        if ($p -and $p.ProcessName -eq 'chrome') { return $true }
        Start-Sleep -Seconds 2
    }
    return $false
}
function SK($k) { [System.Windows.Forms.SendKeys]::SendWait($k) }

function Paste-File($path) {
    Get-Content -Raw -Encoding UTF8 $path | Set-Clipboard
    Start-Sleep -Milliseconds 300
    SK '^v'; Start-Sleep -Milliseconds 1500
    SK '{ENTER}'
}

function Console-Live {
    # ping -> marker download. Proof, not assumption.
    Remove-Item $PING -Force -ErrorAction SilentlyContinue
    Paste-File (Join-Path $root 'band\console_ping.js')
    for ($i = 0; $i -lt 12; $i++) {
        Start-Sleep -Milliseconds 700
        if (Test-Path $PING) { Remove-Item $PING -Force -ErrorAction SilentlyContinue; return $true }
    }
    return $false
}

if (-not (Focus-Chrome)) { Write-Output 'ABORT: cannot focus Chrome'; exit 1 }
if (-not (Test-Path $Js)) { Write-Output "ABORT: no such file $Js"; exit 1 }

$live = Console-Live
if (-not $live) {
    # DevTools was closed (or console lost focus). Toggle once and prove again.
    SK '^+j'; Start-Sleep -Seconds 4
    $live = Console-Live
}
if (-not $live) {
    Write-Output 'FAIL: console not reachable (tried both DevTools states)'
    exit 2
}

Paste-File $Js
Write-Output ("INJECTED(verified) " + $Js + " " + (Get-Date -Format 'HH:mm:ss'))
