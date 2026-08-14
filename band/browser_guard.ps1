# browser_guard.ps1 - read-only foreground Chrome + exact-host gate.
#
# This file NEVER focuses a window and NEVER sends input.  Callers must ask it again
# immediately before every SendKeys/SendInput action.  If Chrome is not already the
# foreground app, its address bar cannot be read, or the host is not an exact domain
# boundary match, the only safe result is refusal.

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

if (-not ('BrowserGuardNative' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class BrowserGuardNative {
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
}
'@
}

$script:BrowserGuardReason = ''
$script:BrowserGuardSites = [ordered]@{
    'band-90610953' = [pscustomobject]@{
        Url = 'https://www.band.us/band/90610953/post'
        Host = 'www.band.us'
        Path = '/band/90610953/post'
    }
    'band-84789192' = [pscustomobject]@{
        Url = 'https://www.band.us/band/84789192/post'
        Host = 'www.band.us'
        Path = '/band/84789192/post'
    }
    'erp' = [pscustomobject]@{
        Url = 'https://loginab.ecount.com/ec5/view/erp'
        Host = 'loginab.ecount.com'
        Path = '/ec5/view/erp'
    }
}

function Get-BGAllowedSite {
    param([Parameter(Mandatory=$true)][string]$SiteKey)
    return $script:BrowserGuardSites[$SiteKey]
}

function Test-BGExactLocation {
    param([Uri]$Uri, [Parameter(Mandatory=$true)][string]$SiteKey)
    $site = Get-BGAllowedSite -SiteKey $SiteKey
    if (-not $site -or -not $Uri) { return $false }
    $path = ([string]$Uri.AbsolutePath).TrimEnd('/')
    $expectPath = ([string]$site.Path).TrimEnd('/')
    return $Uri.Scheme -eq 'https' -and
           $Uri.Host.ToLowerInvariant() -eq $site.Host -and
           $path -eq $expectPath
}

function ConvertTo-BGUri {
    param([string]$RawValue)
    $value = ([string]$RawValue).Trim()
    if (-not $value) { return $null }
    if ($value -notmatch '^[a-zA-Z][a-zA-Z0-9+.-]*://') {
        if ($value -notmatch '^[^\s/]+\.[^\s/]+') { return $null }
        $value = 'https://' + $value
    }
    try {
        $uri = [Uri]$value
        if ($uri.Scheme -notin @('http', 'https') -or -not $uri.Host) { return $null }
        return $uri
    } catch {
        return $null
    }
}

function Get-BGAddressUri {
    param([IntPtr]$WindowHandle, [string]$SiteKey)
    try {
        $window = [System.Windows.Automation.AutomationElement]::FromHandle($WindowHandle)
        if (-not $window) { return $null }
        $editCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Edit)
        $edits = $window.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants, $editCondition)
        $windowRect = $window.Current.BoundingRectangle
        $best = $null
        $bestScore = -1
        foreach ($edit in $edits) {
            # Chrome exposes its real omnibox with this native class.  Requiring it
            # prevents a URL-looking <input> inside a web page from impersonating
            # the address bar and passing the allowlist gate.
            if ($edit.Current.FrameworkId -ne 'Chrome' -or
                $edit.Current.ClassName -ne 'OmniboxViewViews') { continue }
            $pattern = $null
            if (-not $edit.TryGetCurrentPattern(
                    [System.Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) { continue }
            $uri = ConvertTo-BGUri $pattern.Current.Value
            if (-not $uri -or -not (Test-BGExactLocation $uri $SiteKey)) { continue }
            $rect = $edit.Current.BoundingRectangle
            # The omnibox is a wide edit control near the top of the Chrome window.
            # A URL-shaped input inside the web page must not be mistaken for it.
            if ($rect.Width -lt 180 -or $rect.Top -gt ($windowRect.Top + 170)) { continue }
            $score = [int]$rect.Width
            if ($score -gt $bestScore) { $best = $uri; $bestScore = $score }
        }
        return $best
    } catch {
        return $null
    }
}

function Get-ForegroundChromeContext {
    param([Parameter(Mandatory=$true)][string]$SiteKey)
    $script:BrowserGuardReason = ''
    $site = Get-BGAllowedSite -SiteKey $SiteKey
    if (-not $site) {
        $script:BrowserGuardReason = "site is not in the three-page allowlist: $SiteKey"
        return $null
    }
    $hwnd = [BrowserGuardNative]::GetForegroundWindow()
    if ($hwnd -eq [IntPtr]::Zero) {
        $script:BrowserGuardReason = 'there is no foreground window'
        return $null
    }
    $pidValue = [uint32]0
    [BrowserGuardNative]::GetWindowThreadProcessId($hwnd, [ref]$pidValue) | Out-Null
    $process = Get-Process -Id $pidValue -ErrorAction SilentlyContinue
    if (-not $process -or $process.ProcessName -ne 'chrome') {
        $name = if ($process) { $process.ProcessName } else { 'unknown' }
        $script:BrowserGuardReason = "foreground app is $name, not chrome"
        return $null
    }
    $uri = Get-BGAddressUri -WindowHandle $hwnd -SiteKey $SiteKey
    if (-not $uri) {
        $script:BrowserGuardReason = "Chrome address is unreadable or is not exactly $($site.Url)"
        return $null
    }
    return [pscustomobject]@{
        Window = $hwnd
        Pid = [int]$pidValue
        Url = $uri.AbsoluteUri
        Host = $uri.Host.ToLowerInvariant()
        SiteKey = $SiteKey
    }
}
