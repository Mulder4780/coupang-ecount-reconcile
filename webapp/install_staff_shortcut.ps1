param(
  [Parameter(Mandatory = $true)][string]$Url,
  [Parameter(Mandatory = $true)][string]$Name,
  [Parameter(Mandatory = $true)][string]$Icon
)

$ErrorActionPreference = "Stop"
$fixedPrefix = "https://mulder.tailf14aae.ts.net/staff/"
if (-not $Url.StartsWith($fixedPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
  throw "고정 업무센터 주소가 아닙니다."
}
if (-not (Test-Path -LiteralPath $Icon -PathType Leaf)) {
  throw "앱 아이콘을 찾을 수 없습니다."
}

$browserCandidates = @(
  (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"),
  (Join-Path $env:LOCALAPPDATA "Google\Chrome\Application\chrome.exe"),
  (Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe")
)
$programFilesX86 = ${env:ProgramFiles(x86)}
if ($programFilesX86) {
  $browserCandidates += Join-Path $programFilesX86 "Google\Chrome\Application\chrome.exe"
  $browserCandidates += Join-Path $programFilesX86 "Microsoft\Edge\Application\msedge.exe"
}
$browserCandidates = $browserCandidates |
  Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) }

if (-not $browserCandidates) {
  throw "Chrome 또는 Edge를 찾을 수 없습니다."
}

$safeName = ($Name -replace '[\\/:*?"<>|]', ' ').Trim()
if (-not $safeName) { $safeName = "CSOS 업무센터" }
$desktop = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
$shortcutPath = Join-Path $desktop ($safeName + ".lnk")
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $browserCandidates[0]
$shortcut.Arguments = "--app=`"$Url`""
$shortcut.WorkingDirectory = Split-Path -Parent $browserCandidates[0]
$shortcut.IconLocation = "$Icon,0"
$shortcut.Description = "$Name 고정 업무센터"
$shortcut.Save()

Write-Output "INSTALLED"
