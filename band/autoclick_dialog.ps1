# autoclick_dialog.ps1 - dismiss band's "post not found" alert during collection
#
# Why: opening a missing post makes band show alert(), and that modal FREEZES the whole tab.
# Old post ranges contain many deleted posts, so it can pop hundreds of times.
# grab_posts.js is fixed now, but a batch already running cannot be fixed - so we click from outside.
#
# Safety - it does not click just anything:
#   * only inside a Chrome window whose title contains the target word
#   * only Buttons whose Name is exactly the OK word (no blind Enter keystrokes)
#   * stops by itself after the given minutes
#
# ASCII only on purpose: Windows PowerShell 5.1 reads .ps1 as ANSI, so Korean literals
# would arrive mangled and every comparison would silently fail. Build them from code points.
#
#   powershell -ExecutionPolicy Bypass -File band\autoclick_dialog.ps1 -Minutes 45

param([int]$Minutes = 30)

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

$OK = [string][char]0xD655 + [char]0xC778          # 확인
$BAND = [string][char]0xBC34 + [char]0xB4DC        # 밴드
$end = (Get-Date).AddMinutes($Minutes)
$clicked = 0
$root = [System.Windows.Automation.AutomationElement]::RootElement
$scopeChildren = [System.Windows.Automation.TreeScope]::Children
$scopeDesc = [System.Windows.Automation.TreeScope]::Descendants
$typeProp = [System.Windows.Automation.AutomationElement]::ControlTypeProperty
$btnType = [System.Windows.Automation.ControlType]::Button
$winType = [System.Windows.Automation.ControlType]::Window
$invokePat = [System.Windows.Automation.InvokePattern]::Pattern
$winCond = New-Object System.Windows.Automation.PropertyCondition($typeProp, $winType)
$btnCond = New-Object System.Windows.Automation.PropertyCondition($typeProp, $btnType)

Write-Output "watching for $Minutes minutes - will click the OK button of band dialogs"

while ((Get-Date) -lt $end) {
    try {
        foreach ($w in $root.FindAll($scopeChildren, $winCond)) {
            $t = $w.Current.Name
            if (-not $t) { continue }
            if ($t.IndexOf($BAND) -lt 0) { continue }
            if ($w.Current.ClassName -notmatch "Chrome") { continue }
            foreach ($b in $w.FindAll($scopeDesc, $btnCond)) {
                $n = $b.Current.Name
                if ($n -ne $OK -and $n -ne "OK") { continue }
                try {
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
