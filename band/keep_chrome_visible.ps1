# keep_chrome_visible.ps1 [-Restore]
#
# Why (2026-08-11): band is an SPA that paints post bodies via requestAnimationFrame,
# and rAF never fires in a hidden tab. grab_posts.js therefore PAUSES itself whenever
# document.hidden is true - correct behaviour (otherwise it would record empty pages
# as "missing posts"), but it means collection stalls the moment the human covers
# Chrome with Excel or Explorer. Measured: step 2 injected 16:10, still paused 16:35.
#
# Fix without fighting the human for focus: shrink the Chrome window, park it in a
# corner and mark it topmost with SWP_NOACTIVATE. It stays un-occluded (so rAF keeps
# running and collection continues) while keyboard focus stays wherever the human is.
# -Restore puts it back to a normal, non-topmost window.
param([switch]$Restore)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class WP {
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  public static readonly IntPtr TOPMOST = new IntPtr(-1);
  public static readonly IntPtr NOTOPMOST = new IntPtr(-2);
  public const uint NOACTIVATE = 0x0010;
}
'@

$c = Get-Process chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne '' } | Select-Object -First 1
if (-not $c) { Write-Output 'ABORT: no Chrome window'; exit 1 }
$h = $c.MainWindowHandle
[WP]::ShowWindow($h, 9) | Out-Null   # SW_RESTORE - un-minimize; a minimized tab is always hidden

$scr = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea

if ($Restore) {
    [WP]::SetWindowPos($h, [WP]::NOTOPMOST, 60, 60, [int]($scr.Width * 0.75), [int]($scr.Height * 0.85), [WP]::NOACTIVATE) | Out-Null
    Write-Output "restored: $($c.MainWindowTitle)"
} else {
    $w = 620; $ht = 520
    $x = $scr.Right - $w - 20
    $y = $scr.Top + 20
    [WP]::SetWindowPos($h, [WP]::TOPMOST, $x, $y, $w, $ht, [WP]::NOACTIVATE) | Out-Null
    Write-Output "pinned topmost ${w}x${ht} at ${x},${y}: $($c.MainWindowTitle)"
}
