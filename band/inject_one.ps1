# inject_one.ps1 -Js <path> [-OpenTab <url>] : focus Chrome, (optionally open tab), paste JS into DevTools console.
param([string]$Js, [string]$OpenTab = '')
$ErrorActionPreference = 'Stop'
Set-Location "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class W{[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);[DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr h,int c);}'

function Focus-Chrome {
    for ($i = 0; $i -lt 4; $i++) {
        $c = Get-Process chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne '' } | Select-Object -First 1
        if (-not $c) { Start-Sleep -Seconds 2; continue }
        [W]::ShowWindow($c.MainWindowHandle, 9) | Out-Null
        [W]::SetForegroundWindow($c.MainWindowHandle) | Out-Null
        Start-Sleep -Milliseconds 900
        $h = [W]::GetForegroundWindow(); $fp = 0
        [W]::GetWindowThreadProcessId($h, [ref]$fp) | Out-Null
        $p = Get-Process -Id $fp -ErrorAction SilentlyContinue
        if ($p -and $p.ProcessName -eq 'chrome') { return $true }
        Start-Sleep -Seconds 2
    }
    return $false
}
function SK($k) { [System.Windows.Forms.SendKeys]::SendWait($k) }

if (-not (Focus-Chrome)) { Write-Output 'ABORT: cannot focus Chrome'; exit 1 }

if ($OpenTab -ne '') {
    Set-Clipboard $OpenTab
    SK '^t'; Start-Sleep -Milliseconds 800
    SK '^l'; Start-Sleep -Milliseconds 400
    SK '^v'; Start-Sleep -Milliseconds 400
    SK '{ENTER}'
    Write-Output "tab: $OpenTab (10s wait)"
    Start-Sleep -Seconds 10
    if (-not (Focus-Chrome)) { Write-Output 'ABORT: lost Chrome after nav'; exit 1 }
}

if ($Js -ne '') {
    if (-not (Test-Path $Js)) { Write-Output "ABORT: no such file $Js"; exit 1 }
    SK '^+j'; Start-Sleep -Seconds 4   # Ctrl+Shift+J opens console with input focused
    Get-Content -Raw -Encoding UTF8 $Js | Set-Clipboard
    SK '^v'; Start-Sleep -Milliseconds 2500
    SK '{ENTER}'
    Write-Output ("INJECTED " + $Js + " " + (Get-Date -Format 'HH:mm:ss'))
}
