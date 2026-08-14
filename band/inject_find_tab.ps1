# inject_find_tab.ps1 - fail-closed injection into the CURRENT Chrome tab only.
#
# Historical name kept for callers, but tab hunting is intentionally gone.  This file
# never focuses a browser, cycles tabs, opens a URL, or tries Whale/Edge.  The human
# must already be looking at the approved site in Chrome.  Otherwise it exits before
# one keyboard event is sent.
param(
  [Parameter(Mandatory=$true)][string]$Js,
  [Parameter(Mandatory=$true)][string]$SiteKey
)
$ErrorActionPreference = 'Stop'
$root = "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
. (Join-Path $root 'band\browser_guard.ps1')

function Say($m) { Write-Output ((Get-Date -Format 'HH:mm:ss') + ' ' + $m) }

if (-not (Test-Path -LiteralPath $Js)) { Say "ABORT: no payload $Js"; exit 1 }
$context = Get-ForegroundChromeContext -SiteKey $SiteKey
if (-not $context) {
    Say ("ABORT: safe Chrome context required - " + $script:BrowserGuardReason)
    exit 4
}
Say ("verified foreground Chrome host: " + $context.Host)
$out = & (Join-Path $root 'band\inject_here.ps1') `
    -Js $Js -SiteKey $SiteKey -Browser chrome 2>&1
$out | ForEach-Object { Say ("  | " + $_) }
$text = ($out | Out-String)
if ($text -match 'INJECTED on') { Say 'OK: injected into verified Chrome tab'; exit 0 }
Say 'FAIL: verified Chrome tab refused or lost focus'
exit 3
