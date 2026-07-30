param(
    [switch]$Remove
)

$ErrorActionPreference = "Stop"
$TaskName = "쿠팡업무_원장일괄반영"

if ($Remove) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "작업 스케줄러 제거: $TaskName"
    exit 0
}

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = (Get-Command python.exe -ErrorAction Stop).Source
$Pythonw = Join-Path (Split-Path -Parent $Python) "pythonw.exe"
if (-not (Test-Path -LiteralPath $Pythonw)) {
    throw "pythonw.exe를 찾을 수 없습니다: $Pythonw"
}

$Action = New-ScheduledTaskAction `
    -Execute $Pythonw `
    -Argument "ledger_db.py --apply" `
    -WorkingDirectory $Root
$Triggers = @(
    (New-ScheduledTaskTrigger -Daily -At "11:00"),
    (New-ScheduledTaskTrigger -Daily -At "15:00")
)
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 45)
$Principal = New-ScheduledTaskPrincipal `
    -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel Limited

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Triggers `
    -Settings $Settings `
    -Principal $Principal `
    -Description "확정 입력을 SQLite에 모아 매일 11:00·15:00에만 관리대장으로 일괄 반영" `
    -Force | Out-Null

$Task = Get-ScheduledTask -TaskName $TaskName
$Times = ($Task.Triggers | ForEach-Object { ([datetime]$_.StartBoundary).ToString("HH:mm") }) -join " · "
Write-Host "작업 스케줄러 등록: $TaskName ($Times)"
