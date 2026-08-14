# inject_here.ps1 - inject only into an already-active, allowlisted Chrome page.
#
# Safety contract:
#   * Chrome only; Whale/Edge/other apps are never focused or touched.
#   * The current address must exactly match one of browser_guard.ps1's three pages.
#   * The guard is checked again immediately before EVERY keyboard input.
#   * No tab hunt, tab switch, URL opening, window restore, or foreground steal.
#   * If DevTools cannot be proven live, stop.  Never type "allow pasting" blindly.
param(
  [Parameter(Mandatory=$true)][string]$Js,
  [Parameter(Mandatory=$true)][string]$SiteKey,
  [string]$Browser = 'chrome',
  [string]$OpenUrl = ''
)
$ErrorActionPreference = 'Stop'
$root = "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
Set-Location $root
. (Join-Path $root 'band\browser_guard.ps1')
Add-Type -AssemblyName System.Windows.Forms

$downloads = Join-Path $env:USERPROFILE 'Downloads'
$site = Get-BGAllowedSite -SiteKey $SiteKey

function Require-SafeChromePage {
    $context = Get-ForegroundChromeContext -SiteKey $SiteKey
    if (-not $context) {
        throw ("ABORT: unsafe browser context - " + $script:BrowserGuardReason)
    }
    return $context
}

function SK($keys) {
    $null = Require-SafeChromePage
    [System.Windows.Forms.SendKeys]::SendWait($keys)
}

function Paste-File($path) {
    $null = Require-SafeChromePage
    Get-Content -Raw -Encoding UTF8 $path | Set-Clipboard
    Start-Sleep -Milliseconds 350
    $null = Require-SafeChromePage
    [System.Windows.Forms.SendKeys]::SendWait('^v')
    Start-Sleep -Milliseconds 2200
    $null = Require-SafeChromePage
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
}

function Ping-PageHost {
    $null = Require-SafeChromePage
    Get-ChildItem $downloads -Filter '__ping__*.txt' -ErrorAction SilentlyContinue |
        Remove-Item -Force -ErrorAction SilentlyContinue
    Paste-File (Join-Path $root 'band\console_ping.js')
    for ($i = 0; $i -lt 14; $i++) {
        Start-Sleep -Milliseconds 700
        $file = Get-ChildItem $downloads -Filter '__ping__*.txt' -ErrorAction SilentlyContinue |
                Select-Object -First 1
        if ($file) {
            $name = $file.Name
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
            if ($name -match '^__ping__([^_]+(?:\.[^_]+)*)__') { return $Matches[1] }
            return 'unknown'
        }
    }
    return $null
}

if ($Browser.ToLowerInvariant() -ne 'chrome') {
    Write-Output 'ABORT: only Chrome is allowed'; exit 4
}
if ($OpenUrl) {
    Write-Output 'ABORT: automatic navigation is disabled; open the approved page in Chrome yourself'; exit 4
}
if (-not $site) {
    Write-Output "ABORT: site is not allowlisted: $SiteKey"; exit 4
}
if (-not (Test-Path -LiteralPath $Js)) {
    Write-Output "ABORT: no such file $Js"; exit 1
}
$context = Get-ForegroundChromeContext -SiteKey $SiteKey
if (-not $context) {
    Write-Output ("ABORT: safe Chrome page required - " + $script:BrowserGuardReason); exit 4
}
Write-Output ("verified Chrome page: " + $context.Url)

# The current page is proven safe before the first shortcut.  If focus or URL changes,
# SK/Paste-File refuses before the next input instead of typing into the new program.
SK '^+j'; Start-Sleep -Seconds 4
$actualHost = Ping-PageHost
if (-not $actualHost) {
    # One deterministic retry handles the DevTools toggle state.  No literal text is
    # ever typed because an unfocused console can send it to the page/address bar.
    SK '^+j'; Start-Sleep -Seconds 4
    $actualHost = Ping-PageHost
}
if (-not $actualHost) {
    Write-Output 'FAIL: console not reachable; no payload was pasted'; exit 2
}
if ($actualHost.ToLowerInvariant() -ne $site.Host) {
    Write-Output "FAIL: console belongs to '$actualHost', expected '$($site.Host)' - refusing payload"
    exit 3
}

$null = Require-SafeChromePage
Paste-File $Js
Write-Output ("INJECTED on " + $actualHost + " at " + $site.Path + " : " + $Js + " " +
              (Get-Date -Format 'HH:mm:ss'))
exit 0
