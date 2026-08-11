# collect_when_idle.ps1 - run the band collection chain only while the human is away.
#
# Why (2026-08-11): two hard constraints collide.
#   * band is an SPA whose post bodies are painted by requestAnimationFrame, so a
#     hidden/covered tab collects nothing (grab_posts.js pauses itself on purpose).
#   * driving the console needs the Chrome window in the foreground, and Windows
#     refuses foreground steals while the human is typing - measured repeatedly today.
# Fighting for focus only interrupts the human's own work. So we wait for real idle
# (GetLastInputInfo) and do the whole chain then. Nothing is lost by waiting: a paused
# collector resumes, and the dumps are absorbed by dump_watch whenever they land.
#
# Progress is written to reports/밴드_무인수집.log so a dead run leaves a trace
# (a step that dies silently is indistinguishable from a step that never ran).
param([int]$IdleSeconds = 180, [int]$MaxHours = 14)
$ErrorActionPreference = 'Continue'
$root = "C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
Set-Location $root
$log = Join-Path $root 'reports\밴드_무인수집.log'
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class Idle {
  [StructLayout(LayoutKind.Sequential)] struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }
  [DllImport("user32.dll")] static extern bool GetLastInputInfo(ref LASTINPUTINFO p);
  [DllImport("kernel32.dll")] static extern uint GetTickCount();
  public static uint Seconds() {
    LASTINPUTINFO i = new LASTINPUTINFO();
    i.cbSize = (uint)Marshal.SizeOf(i);
    if (!GetLastInputInfo(ref i)) return 0;
    return (GetTickCount() - i.dwTime) / 1000;
  }
}
'@

function Say($m) {
    $line = "[{0}] {1}" -f (Get-Date -Format 'MM-dd HH:mm:ss'), $m
    Write-Output $line
    Add-Content -Path $log -Value $line -Encoding UTF8
}

# Only steps whose target posts actually exist. The "new posts" ranges generated today
# pointed above the real latest post number (84789192 max = 3539, plan asked 3540-3579)
# and band answers 200 for numbers that do not exist yet - 40 numbers = ~14 wasted
# minutes and zero harvest. Comments are where today's numbers actually change.
$STEPS = @(
  @{ name = 'comments-90610953'; url = 'https://band.us/band/90610953'; js = 'band\댓글채우기_붙여넣기_90610953.js'; waitMin = 30 },
  @{ name = 'comments-84789192'; url = 'https://band.us/band/84789192'; js = 'band\댓글채우기_붙여넣기_84789192.js'; waitMin = 8 }
)

Say "무인 수집 대기 시작 - 사람 입력이 ${IdleSeconds}초 없으면 시작한다 (최대 ${MaxHours}시간 대기)"
$deadline = (Get-Date).AddHours($MaxHours)

foreach ($s in $STEPS) {
    if (-not (Test-Path $s.js)) { Say "건너뜀 $($s.name) - 파일 없음 $($s.js)"; continue }

    # Wait for the human to be away AND the desktop to be unlocked.
    # 2026-08-11: a 2.8h idle window was wasted because the PC was at the lock screen -
    # LogonUI owns the desktop, so SetForegroundWindow fails and every tab counts as
    # hidden (band renders nothing). Idle alone is not the condition; unlocked+idle is.
    $waited = $false
    while ((Get-Date) -lt $deadline) {
        $locked = [bool](Get-Process logonui -ErrorAction SilentlyContinue)
        if (-not $locked -and [Idle]::Seconds() -ge $IdleSeconds) { break }
        if (-not $waited) {
            $why = if ($locked) { '화면 잠김 - 잠금이 풀리기를 기다린다' } else { '사람이 사용 중 - 자리 비우기를 기다린다' }
            Say "$($s.name): $why"
            $waited = $true
        }
        Start-Sleep -Seconds 20
    }
    if ((Get-Date) -ge $deadline) { Say '대기 시간 초과 - 그만둔다'; break }

    Say "$($s.name): 시작 (유휴 $([Idle]::Seconds())초)"
    $before = (Get-ChildItem "$env:USERPROFILE\Downloads" -Filter 'dump_*.json' -ErrorAction SilentlyContinue | Measure-Object).Count
    $out = & powershell -ExecutionPolicy Bypass -File (Join-Path $root 'band\collect_step.ps1') -Url $s.url -Js $s.js 2>&1
    Say "$($s.name): 주입 결과 - $out"
    if ("$out" -notlike '*INJECTED*') { Say "$($s.name): 주입 실패 - 다음 단계로"; continue }

    # wait for the dump; a paused collector (human came back) just takes longer
    $end = (Get-Date).AddMinutes($s.waitMin)
    $got = $false
    while ((Get-Date) -lt $end) {
        Start-Sleep -Seconds 20
        $now = (Get-ChildItem "$env:USERPROFILE\Downloads" -Filter 'dump_*.json' -ErrorAction SilentlyContinue | Measure-Object).Count
        if ($now -gt $before) { $got = $true; break }
    }
    if ($got) {
        Say "$($s.name): 덤프 도착 - 흡수한다"
        $ab = & python band/dump_watch.py --once 2>&1 | Select-Object -Last 4
        Say "$($s.name): 흡수 - $ab"
    } else {
        Say "$($s.name): $($s.waitMin)분 내 덤프 없음 (탭이 가려져 멈췄을 수 있다 - 다음 회차가 이어받는다)"
    }
}

Say '무인 수집 회차 끝'
