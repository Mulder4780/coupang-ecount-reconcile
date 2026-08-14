# focus_collect_tab.ps1 - deliberately does NOT focus or switch anything anymore.
# Kept for old callers: it only confirms that an approved Band page is already the
# foreground Chrome tab.  Bringing a hidden collector forward stole the operator's
# keyboard, so absence of the right foreground page now means a safe no-op.
param([string]$SiteKey = 'band-90610953')
$ErrorActionPreference = 'Stop'
$root = "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
. (Join-Path $root 'band\browser_guard.ps1')

if ($SiteKey -notin @('band-90610953', 'band-84789192')) {
    Write-Output 'ABORT: only the two approved Band pages are allowed'; exit 4
}
$context = Get-ForegroundChromeContext -SiteKey $SiteKey
if (-not $context) {
    Write-Output ("ABORT: collector page is not already foreground - " +
                  $script:BrowserGuardReason); exit 4
}
Write-Output ("OK: approved collector page is already foreground: " + $context.Url)
exit 0
