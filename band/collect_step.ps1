# collect_step.ps1 - compatibility wrapper for strict Chrome collection.
# It no longer opens/focuses Chrome or creates a tab.  The approved page must already
# be the foreground Chrome tab; inject_here.ps1 enforces the three-page allowlist.
param(
  [Parameter(Mandatory=$true)][string]$Url,
  [Parameter(Mandatory=$true)][string]$Js,
  [string]$SiteKey = ''
)
$ErrorActionPreference = 'Stop'
$root = "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
. (Join-Path $root 'band\browser_guard.ps1')

if (-not $SiteKey) {
    foreach ($key in $script:BrowserGuardSites.Keys) {
        $site = Get-BGAllowedSite -SiteKey $key
        try { $uri = [Uri]$Url } catch { continue }
        if (Test-BGExactLocation -Uri $uri -SiteKey $key) { $SiteKey = $key; break }
    }
}
if (-not (Get-BGAllowedSite -SiteKey $SiteKey)) {
    Write-Output 'ABORT: URL is not one of the three approved pages'; exit 4
}
& (Join-Path $root 'band\inject_here.ps1') -Js $Js -SiteKey $SiteKey -Browser chrome
exit $LASTEXITCODE
