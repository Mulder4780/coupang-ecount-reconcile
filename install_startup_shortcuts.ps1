# 쿠팡 업무 시작프로그램을 콘솔 창 없이 등록한다.
# 재실행해도 같은 두 바로가기를 갱신하므로 멱등이다.
$ErrorActionPreference = "Stop"

$Root = $PSScriptRoot
$Startup = [Environment]::GetFolderPath("Startup")
$PythonW = Join-Path $env:LOCALAPPDATA "Programs\Python\Python312\pythonw.exe"
$WScript = Join-Path $env:WINDIR "System32\wscript.exe"
function Text-FromCodes([int[]]$Codes) {
    return -join ($Codes | ForEach-Object { [char]$_ })
}
# Windows PowerShell 5.1은 BOM 없는 UTF-8 스크립트의 한글 문자열을 깨뜨린다.
# 파일 이름은 코드포인트로 조립해 어느 실행 정책/인코딩에서도 같은 .lnk를 가리킨다.
$ServerLinkName = (Text-FromCodes @(0xCFE0,0xD321,0xC5C5,0xBB34,0xC571,0xC11C,0xBC84)) + ".lnk"
$WorkLinkName = (Text-FromCodes @(0xCFE0,0xD321,0xC5C5,0xBB34)) + "_" +
    (Text-FromCodes @(0xC5C5,0xBB34,0xC2DC,0xC791)) + ".lnk"
$WorkBatchName = (Text-FromCodes @(0xC5C5,0xBB34,0xC2DC,0xC791)) + ".bat"

if (-not (Test-Path -LiteralPath $PythonW)) { throw "pythonw.exe not found: $PythonW" }
if (-not (Test-Path -LiteralPath (Join-Path $Root "run_hidden.vbs"))) {
    throw "run_hidden.vbs not found"
}

$Shell = New-Object -ComObject WScript.Shell

# 서버는 배치 파일을 거치지 않고 GUI Python으로 바로 시작한다.
$ServerLink = $Shell.CreateShortcut((Join-Path $Startup $ServerLinkName))
$ServerLink.TargetPath = $PythonW
$ServerLink.Arguments = '"' + (Join-Path $Root "webapp\server_guard.py") + '"'
$ServerLink.WorkingDirectory = $Root
$ServerLink.WindowStyle = 7
$ServerLink.Save()

# 업무 탭 준비 배치는 최대 90초 대기하므로 VBS가 콘솔을 끝까지 숨긴다.
$WorkLink = $Shell.CreateShortcut((Join-Path $Startup $WorkLinkName))
$WorkLink.TargetPath = $WScript
$WorkLink.Arguments = '"' + (Join-Path $Root "run_hidden.vbs") + '" "' +
    (Join-Path $Root $WorkBatchName) + '"'
$WorkLink.WorkingDirectory = $Root
$WorkLink.WindowStyle = 7
$WorkLink.Save()

Write-Output "Updated 2 startup shortcuts without console windows."
