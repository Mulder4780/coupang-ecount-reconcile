# autoclick_dialog.ps1 - dismiss band's "post not found" alert during collection
#
# Why: opening a missing post makes band show alert(), and that modal FREEZES the whole tab.
# Old post ranges contain many deleted posts, so it can pop hundreds of times.
# grab_posts.js is fixed now, but a batch already running cannot be fixed - so we click from outside.
#
# Safety - it does not click just anything:
#   * only inside the foreground Chrome window on one exact approved Band page
#   * only Buttons whose Name is exactly the OK word (no blind Enter keystrokes)
#   * stops by itself after the given minutes
#
# ASCII only on purpose: Windows PowerShell 5.1 reads .ps1 as ANSI, so Korean literals
# would arrive mangled and every comparison would silently fail. Build them from code points.
#
#   powershell -ExecutionPolicy Bypass -File band\autoclick_dialog.ps1 -Minutes 45

param([int]$Minutes = 30, [string]$SiteKey = 'band-90610953')

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$guardRoot = "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
. (Join-Path $guardRoot 'band\browser_guard.ps1')

$OK = [string][char]0xD655 + [char]0xC778          # 확인
$end = (Get-Date).AddMinutes($Minutes)
$clicked = 0
$scopeDesc = [System.Windows.Automation.TreeScope]::Descendants
$typeProp = [System.Windows.Automation.AutomationElement]::ControlTypeProperty
$btnType = [System.Windows.Automation.ControlType]::Button
$invokePat = [System.Windows.Automation.InvokePattern]::Pattern
$btnCond = New-Object System.Windows.Automation.PropertyCondition($typeProp, $btnType)

if ($SiteKey -notin @('band-90610953', 'band-84789192')) {
    Write-Output 'ABORT: only the two approved Band pages are allowed'; exit 4
}
Write-Output "watching for $Minutes minutes - foreground approved Chrome page only"

while ((Get-Date) -lt $end) {
    try {
        $context = Get-ForegroundChromeContext -SiteKey $SiteKey
        if ($context) {
            $w = [System.Windows.Automation.AutomationElement]::FromHandle($context.Window)
            foreach ($b in $w.FindAll($scopeDesc, $btnCond)) {
                $n = $b.Current.Name
                if ($n -ne $OK -and $n -ne "OK") { continue }
                try {
                    # Re-check after finding the button: focus/url may have changed.
                    if (-not (Get-ForegroundChromeContext -SiteKey $SiteKey)) { continue }
                    $b.GetCurrentPattern($invokePat).Invoke()
                    $clicked++
                    Write-Output ("[{0}] clicked OK (total {1})" -f (Get-Date -Format HH:mm:ss), $clicked)
                } catch { }
            }
        }
    } catch { }
    Start-Sleep -Milliseconds 700
}
Write-Output "done - clicked $clicked times"
