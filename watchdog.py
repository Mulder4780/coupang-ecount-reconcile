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


def heal_tunnel(dry):
    url_f = os.path.join(ROOT, "reports", "tunnel_url.txt")
    url = ""
    try:
        url = open(url_f, encoding="utf-8").read().strip()
    except Exception:
        pass
    alive = proc_running("cloudflared")
    if alive and url:
        return f"터널 정상 ({url[8:40]}...)"
    if dry:
        return "터널 죽음(dry)"
    if not os.path.exists(os.path.join(ROOT, "webapp", "cloudflared.exe")) and not alive:
        return "터널 스킵 — cloudflared 없음"
    start_hidden(os.path.join("webapp", "tunnel_run.py"))
    return "터널 재시작 지시(주소는 tunnel_url.txt·앱 대시보드에 갱신됨)"


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


def main():
    dry = "--dry" in sys.argv
    results = [heal_server(dry), heal_tunnel(dry), archive_versions(dry), clean_reports(dry)]
    log(" | ".join(results))


if __name__ == "__main__":
    main()
