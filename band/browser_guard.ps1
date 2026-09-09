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
        Hosts = @('www.band.us', 'band.us')
        Path = '/band/90610953/post'
    }
    'band-84789192' = [pscustomobject]@{
        Url = 'https://www.band.us/band/84789192/post'
        Host = 'www.band.us'
        Hosts = @('www.band.us', 'band.us')
        Path = '/band/84789192/post'
    }
    # 2026-09-09 실측: 이카운트가 경로를 ec5 -> ec56 으로 올렸다(형님 화면에서
    #   읽힌 값 loginab.ecount.com/ec56/view/erp?w_flag=1...).  그래서 관문이
    #   "주소가 정확히 그것이 아니다" 로 막았고, 그 문구가 사람에게 **주소를
    #   바꾸라**고 두 번 헛 부탁하게 만들었다([172] 틀린 지목).
    # * 옛 경로를 지우지 않는다 - 옛 판을 쓰는 기계에서 그날부터 못 긁는다([172]).
    #   넓히는 것이 아니라 **정확 일치를 둘로** 둔 것이다(정규식으로 헐겁게 하지 않는다).
    'erp' = [pscustomobject]@{
        Url = 'https://loginab.ecount.com/ec56/view/erp'
        Host = 'loginab.ecount.com'
        Path = '/ec56/view/erp'
        Paths = @('/ec56/view/erp', '/ec5/view/erp')
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
    # * Paths 를 가진 자리만 여러 경로를 받는다 - 없는 자리(밴드 둘)는 예전
    #   그대로 Path 하나다([172] 좁히는 것도 넓히는 것도 고장이다).
    # * 2026-09-09 실측: 크롬 주소창이 'www.' 를 숨겨 band.us/band/.../post 로
    #   읽힌다.  www.band.us 와 band.us 는 **같은 사이트**다 - 정확 일치를 둘로
    #   둘 뿐 다른 도메인을 받는 것이 아니다([172]).  Hosts 가 없는 자리는 예전 그대로.
    $expectHost = @()
    if ($site.PSObject.Properties.Match('Hosts').Count -gt 0 -and $site.Hosts) {
        foreach ($h in $site.Hosts) { $expectHost += ([string]$h).ToLowerInvariant() }
    } else {
        $expectHost = @(([string]$site.Host).ToLowerInvariant())
    }
    $expect = @()
    if ($site.PSObject.Properties.Match('Paths').Count -gt 0 -and $site.Paths) {
        foreach ($one in $site.Paths) { $expect += ([string]$one).TrimEnd('/') }
    } else {
        $expect = @(([string]$site.Path).TrimEnd('/'))
    }
    return $Uri.Scheme -eq 'https' -and
           ($expectHost -contains $Uri.Host.ToLowerInvariant()) -and
           ($expect -contains $path)
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
