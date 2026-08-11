# inject_and_verify.ps1 -Js <payload> -ExpectHost <substr> [-Probe <probe js>] [-WaitSec N]
#
# Inject a collector AND PROVE IT IS STILL ALIVE afterwards.
#
# Why this exists (incident #35, 2026-08-11):
#   The injection itself was proven (hostname ping) and still harvested nothing.
#   The top document had navigated to a post page, which destroys the injected
#   script instantly - no error, no console message, cache unchanged. Several
#   sessions then read "comments not scraped yet" when the truth was "scraping
#   started, then died". An injection that is not verified alive is not a
#   collection, and must never be reported as one.
#
# So: inject -> wait -> ask the page for its own state -> report ALIVE or DEAD.
# The probe writes a JSON blob to Downloads; a missing global means the script
# is gone. We print the raw state so the number, not the claim, is the report.
#
# ASCII-only on purpose: PowerShell 5.1 reads BOM-less UTF-8 as CP949 and
# mangles non-ASCII, which once killed a watcher without a single log line.
param(
  [Parameter(Mandatory=$true)][string]$Js,
  [Parameter(Mandatory=$true)][string]$ExpectHost,
  [string]$Probe = '',
  [int]$WaitSec = 45
)
$ErrorActionPreference = 'Stop'
$root = "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
$dl   = Join-Path $env:USERPROFILE 'Downloads'
if (-not $Probe) { $Probe = Join-Path $root 'band\band_dump_state.js' }

function Say($m) { Write-Output ((Get-Date -Format 'HH:mm:ss') + ' ' + $m) }

if (-not (Test-Path $Js))    { Say "ABORT: no payload $Js"; exit 1 }
if (-not (Test-Path $Probe)) { Say "ABORT: no probe $Probe"; exit 1 }

# --- 1) inject (inject_here.ps1 refuses to paste on the wrong origin) ---------
Say "inject: $Js"
$out = & (Join-Path $root 'band\inject_here.ps1') -Js $Js -ExpectHost $ExpectHost 2>&1
$out | ForEach-Object { Say ("  | " + $_) }
# Judge by what the injector SAID, not only by its exit code. A .ps1 that ends on
# Write-Output leaves $LASTEXITCODE untouched, so an empty code is not a failure -
# reading it as one aborted a run whose payload had already landed (2026-08-11).
$txt = ($out | Out-String)
if ($txt -match 'ABORT:|FAIL:') { Say 'ABORT: injector refused'; exit 3 }
if ($txt -notmatch 'INJECTED on') { Say 'ABORT: injector did not confirm a paste'; exit 2 }

# --- 2) let it start ---------------------------------------------------------
Say "wait ${WaitSec}s before liveness probe"
Start-Sleep -Seconds $WaitSec

# --- 3) ask the page for its own state ---------------------------------------
Get-ChildItem $dl -Filter '__grabstate__*.json' -EA SilentlyContinue |
    Remove-Item -Force -EA SilentlyContinue
Say "probe: $Probe"
$out2 = & (Join-Path $root 'band\inject_here.ps1') -Js $Probe -ExpectHost $ExpectHost 2>&1
$out2 | ForEach-Object { Say ("  | " + $_) }
if ($LASTEXITCODE -ne 0) { Say "UNKNOWN: probe could not be pasted (exit $LASTEXITCODE)"; exit 4 }

$state = $null
for ($i = 0; $i -lt 20; $i++) {
    Start-Sleep -Milliseconds 700
    $f = Get-ChildItem $dl -Filter '__grabstate__*.json' -EA SilentlyContinue | Select-Object -First 1
    if ($f) { $state = Get-Content -Raw -Encoding UTF8 $f.FullName; Remove-Item $f.FullName -Force -EA SilentlyContinue; break }
}
if (-not $state) { Say 'UNKNOWN: probe returned nothing - console may have closed'; exit 4 }

Say 'state:'
$state -split "`n" | Select-Object -First 40 | ForEach-Object { Say ("  > " + $_) }

# A missing global is the exact signature of a dead script (see header).
# Each payload has its own name for it: the band collector says NO __GRAB, the
# ERP tour says NOSTATE. Check both - a probe that cannot recognise death will
# happily report ALIVE for a script that never started.
if ($state -match 'NO __GRAB' -or $state -match 'NOSTATE') {
    Say 'DEAD: the global is gone - the script never started or the page navigated. NOT a collection.'
    exit 5
}
Say 'ALIVE: collector is running. Harvest still must be counted from the cache.'
exit 0
