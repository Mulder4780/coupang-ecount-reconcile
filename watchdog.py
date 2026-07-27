# -*- coding: utf-8 -*-
"""
watchdog.py — 자가치유 워치독 (무인 운영의 심장)
===================================================
30분마다 작업 스케줄러가 실행. 사람이 손대지 않아도:
  1. 앱 서버 죽음 감지 → 자동 재시작
  2. 외부접속 터널 죽음/주소소실 감지 → 자동 재시작
  3. 관리대장 구버전 자동 아카이브 (최신 3개만 현역, 나머지 OLD/로)
  4. 30일 지난 리포트 자동 정리
모든 조치는 reports/watchdog_log.txt 에 기록.

실행:  python watchdog.py            # 점검+복구 1회
       python watchdog.py --dry      # 점검만(복구·이동 없음)
"""
import sys, os, re, glob, json, time, subprocess, urllib.request
from datetime import datetime, timedelta

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
PYW = PY.replace("python.exe", "pythonw.exe")
LOG = os.path.join(ROOT, "reports", "watchdog_log.txt")
PORT = 8899


def log(msg):
    os.makedirs(os.path.dirname(LOG), exist_ok=True)
    line = f"[{datetime.now():%m-%d %H:%M}] {msg}"
    print(line)
    try:
        open(LOG, "a", encoding="utf-8").write(line + "\n")
    except Exception:
        pass


# ── 판단 로직 (합성 테스트 대상 — 순수 함수) ─────────────────────
def pick_archive(version_files, keep=3):
    """[(버전, 경로)] → OLD로 옮길 경로 목록 (최신 keep개 제외)"""
    s = sorted(version_files, key=lambda x: x[0], reverse=True)
    return [p for _, p in s[keep:]]


def pick_old_reports(files_with_mtime, days=30, protect=("agent_status.json", "tunnel_url.txt", "watchdog_log.txt")):
    """[(경로, mtime)] → 삭제 대상 (보호 파일 제외, days일 초과)"""
    cut = time.time() - days * 86400
    return [p for p, mt in files_with_mtime
            if mt < cut and os.path.basename(p) not in protect]


# ── 점검·복구 ─────────────────────────────────────────────
def ping():
    try:
        with urllib.request.urlopen(f"http://localhost:{PORT}/api/ping", timeout=5) as r:
            return b"coupang-work" in r.read()
    except Exception:
        return False


def proc_running(image):
    """PowerShell Get-Process 기반(로케일·wmic 무관, 신뢰성 우선)"""
    name = image.replace(".exe", "")
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-Command",
                              f"(Get-Process {name} -ErrorAction SilentlyContinue).Count"],
                             capture_output=True, text=True, timeout=20).stdout.strip()
        return out.isdigit() and int(out) > 0
    except Exception:
        return False


def kill_by_cmdline(needle):
    """CommandLine에 needle 포함된 python 프로세스 종료 (PowerShell CIM — wmic 대체)"""
    ps = ("Get-CimInstance Win32_Process -Filter \"Name='python.exe' or Name='pythonw.exe'\" | "
          f"Where-Object {{ $_.CommandLine -like '*{needle}*' }} | "
          "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }")
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       capture_output=True, timeout=30)
    except Exception:
        pass


def start_hidden(script):
    subprocess.Popen([PYW if os.path.exists(PYW) else PY, os.path.join(ROOT, script)],
                     cwd=ROOT, creationflags=0x00000008)  # DETACHED_PROCESS


def heal_server(dry):
    if ping():
        return "서버 정상"
    if dry:
        return "서버 죽음(dry — 복구 생략)"
    kill_by_cmdline("app_server.py")
    start_hidden(os.path.join("webapp", "app_server.py"))
    time.sleep(4)
    return "서버 재시작 → " + ("성공" if ping() else "실패(다음 주기 재시도)")


def tunnel_alive(url):
    """프로세스 개수가 아니라 **실제 응답**으로 판정.
    (Get-Process는 1개일 때 Count가 비어 나와 살아있는 터널을 죽었다고 오판 →
     불필요한 재시작으로 공개 주소가 바뀌던 실사고가 있었다)"""
    if not url:
        return False
    try:
        with urllib.request.urlopen(url.rstrip("/") + "/api/ping", timeout=12) as r:
            return b"coupang-work" in r.read()
    except Exception:
        return False


def heal_tunnel(dry):
    url_f = os.path.join(ROOT, "reports", "tunnel_url.txt")
    url = ""
    try:
        url = open(url_f, encoding="utf-8").read().strip()
    except Exception:
        pass
    if tunnel_alive(url):
        return f"터널 정상 ({url[8:40]}...)"
    if dry:
        return "터널 죽음(dry)"
    # ★ 좀비 정리를 먼저 한다.
    #   tunnel_run은 포트 8977 싱글톤 락을 쓴다. 예전 tunnel_run이 살아 있으면
    #   새로 띄워도 "이미 실행 중"으로 즉시 끝나 **아무것도 고쳐지지 않는다**.
    #   실제로 cloudflared는 살아 있는데 주소만 만료된 채 이틀 방치됐다(2026-07-27).
    killed = kill_stale_tunnel()
    if not os.path.exists(os.path.join(ROOT, "webapp", "cloudflared.exe")):
        return f"터널 스킵 — cloudflared 없음{killed}"
    start_hidden(os.path.join("webapp", "tunnel_run.py"))
    return f"터널 재시작 지시{killed} (새 주소는 고정주소가 자동으로 따라감)"


def kill_stale_tunnel():
    """죽은 터널을 붙들고 있는 cloudflared·tunnel_run을 정리한다."""
    import subprocess as sp
    ps = ("Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'cloudflared.exe' -or "
          "$_.CommandLine -like '*tunnel_run*' } | ForEach-Object { "
          "Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue; $_.ProcessId }")
    try:
        r = sp.run(["powershell", "-NoProfile", "-Command", ps],
                   capture_output=True, text=True, timeout=60)
        n = len([x for x in (r.stdout or "").split() if x.strip().isdigit()])
        return f" (좀비 {n}개 정리)" if n else ""
    except Exception:
        return ""


def archive_versions(dry):
    try:
        sys.path.insert(0, ROOT)
        from ecount_reconcile import load_config
        folder = os.path.dirname(load_config()["reconcile"]["master_xlsx"])
        vers = []
        for p in glob.glob(os.path.join(folder, "쿠팡_통합업무_일일보고_관리대장_v*.xlsx")):
            m = re.search(r"_v(\d+)\.xlsx$", p)
            if m and "~$" not in p:
                vers.append((int(m.group(1)), p))
        targets = pick_archive(vers, keep=3)
        if not targets:
            return "버전 아카이브 불필요(현역 3개 이하)"
        old = os.path.join(folder, "OLD")
        os.makedirs(old, exist_ok=True)
        moved = 0
        for p in targets:
            if dry:
                moved += 1; continue
            try:
                os.replace(p, os.path.join(old, os.path.basename(p)))
                moved += 1
            except PermissionError:
                pass          # 열려 있으면 다음 주기에
        return f"구버전 {moved}개 → OLD{'(dry)' if dry else ''}"
    except Exception as e:
        return f"아카이브 오류: {e}"


def clean_reports(dry):
    files = [(p, os.path.getmtime(p)) for p in glob.glob(os.path.join(ROOT, "reports", "*"))
             if os.path.isfile(p)]
    targets = pick_old_reports(files, days=30)
    if not dry:
        for p in targets:
            try:
                os.remove(p)
            except Exception:
                pass
    return f"오래된 리포트 {len(targets)}개 정리{'(dry)' if dry else ''}"


def publish_endpoint(dry):
    """터널 주소가 바뀌면 고정 주소(GitHub Pages)에 자동 게시 — 폰 북마크 불변"""
    if dry:
        return "게시(dry)"
    try:
        r = subprocess.run([PY, os.path.join(ROOT, "publish_endpoint.py")], cwd=ROOT,
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=120)
        return (r.stdout or "").strip().splitlines()[-1][:60] if r.stdout.strip() else "게시 무응답"
    except Exception as e:
        return f"게시 오류: {str(e)[:40]}"


def main():
    dry = "--dry" in sys.argv
    results = [heal_server(dry), heal_tunnel(dry), publish_endpoint(dry),
               archive_versions(dry), clean_reports(dry)]
    log(" | ".join(results))


if __name__ == "__main__":
    main()
