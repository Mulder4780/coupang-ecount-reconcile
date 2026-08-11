# band_auto_inject.ps1 - opens logged-in Chrome band tab, injects collectors in sequence.
# Human runs this ONCE. Everything after is automatic (dumps -> dump_watch -> cache).
# ASCII-only on purpose: PS 5.1 without BOM mangles non-ASCII.
$ErrorActionPreference = 'Stop'
Set-Location "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class W{[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);[DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr h,int c);}'

function Get-ChromeWindow {
    Get-Process chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne '' } | Select-Object -First 1
}

function Focus-Chrome {
    for ($i = 0; $i -lt 5; $i++) {
        $c = Get-ChromeWindow
        if (-not $c) { Start-Sleep -Seconds 3; continue }
        [W]::ShowWindow($c.MainWindowHandle, 9) | Out-Null
        [W]::SetForegroundWindow($c.MainWindowHandle) | Out-Null
        Start-Sleep -Milliseconds 900
        $h = [W]::GetForegroundWindow(); $fp = 0
        [W]::GetWindowThreadProcessId($h, [ref]$fp) | Out-Null
        $p = Get-Process -Id $fp -ErrorAction SilentlyContinue
        if ($p -and $p.ProcessName -eq 'chrome') { return $true }
        Start-Sleep -Seconds 3
    }
    return $false
}

function Send-Keys($k) { [System.Windows.Forms.SendKeys]::SendWait($k) }

function Inject-File($jsPath) {
    if (-not (Focus-Chrome)) { Write-Output "FAIL: cannot focus Chrome for $jsPath"; return $false }
    Send-Keys '^+j'   # focus DevTools console (opens it if closed, focuses if open)
    Start-Sleep -Seconds 2
    Get-Content -Raw -Encoding UTF8 $jsPath | Set-Clipboard
    Send-Keys '^v'; Start-Sleep -Milliseconds 1200
    Send-Keys '{ENTER}'
    Write-Output ("INJECTED: " + $jsPath + " at " + (Get-Date -Format 'HH:mm:ss'))
    return $true
}

# --- step 0: open band tab + devtools console (once) ---
if (-not (Focus-Chrome)) { Write-Output 'ABORT: no Chrome window'; exit 1 }
Set-Clipboard 'https://band.us/band/84789192'
Send-Keys '^t'; Start-Sleep -Milliseconds 800
Send-Keys '^l'; Start-Sleep -Milliseconds 400
Send-Keys '^v'; Start-Sleep -Milliseconds 400
Send-Keys '{ENTER}'
Write-Output 'band tab opening, wait 10s'
Start-Sleep -Seconds 10
if (-not (Focus-Chrome)) { Write-Output 'ABORT: lost Chrome'; exit 1 }
Send-Keys '^+j'   # DevTools console
Start-Sleep -Seconds 4

# --- step 1: recollect 84789192 (15 posts, ~2min) ---
if (-not (Inject-File 'band\재수집_붙여넣기_84789192.js')) { exit 1 }
Write-Output 'waiting 210s for step 1'
Start-Sleep -Seconds 210

# --- step 2: new posts 84789192 (40 posts, ~4min) ---
if (-not (Inject-File 'band\수집_붙여넣기_84789192.js')) { exit 1 }
Write-Output 'waiting 330s for step 2'
Start-Sleep -Seconds 330

# --- step 3: new posts 90610953 (35 posts, ~4min) ---
if (-not (Inject-File 'band\수집_붙여넣기_90610953.js')) { exit 1 }
Write-Output 'waiting 330s for step 3'
Start-Sleep -Seconds 330

# --- step 4: comment backfill 90610953 (181 posts, ~18min) ---
if (-not (Inject-File 'band\댓글채우기_붙여넣기_90610953.js')) { exit 1 }
Write-Output 'waiting 1200s for step 4 (comments)'
Start-Sleep -Seconds 1200

Write-Output 'ALL DONE - dumps are picked up automatically (dump_watch running)'
