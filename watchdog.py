# -*- coding: utf-8 -*-
"""
watchdog.py — 자가치유 워치독 (무인 운영의 심장)
===================================================
30분마다 작업 스케줄러가 실행. 사람이 손대지 않아도:
  1. 앱 서버 죽음 감지 → 자동 재시작
  2. 외부접속 터널 죽음/주소소실 감지 → 자동 재시작
  3. 30일 지난 리포트 자동 정리
관리대장 버전 정리는 `ledger_versions.py` 한 곳에서만 수행한다.
모든 조치는 reports/watchdog_log.txt 에 기록.

실행:  python watchdog.py            # 점검+복구 1회
       python watchdog.py --dry      # 점검만(복구·이동 없음)
"""
import sys, os, re, glob, json, time, subprocess, urllib.request
from datetime import datetime, timedelta
from operation_window import input_window_label, is_input_window
from proc_guard import run_tree

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
def gap_note(last_line, now, expect_min=30, slack_min=15):
    """직전 기록과 지금 사이가 너무 벌어졌으면 그 사실을 남긴다.

    ★ 2026-07-28: 워치독이 16:27~20:43 **4시간 넘게 안 돌았고**, 그 사이 터널 주소가
      죽어 폰 접속이 끊겼다. 그런데 로그에는 '정상'만 줄줄이 남아 있어 아무도 몰랐다
      (작업 스케줄러가 놓친 실행을 따라잡지 않는 설정이었다 — StartWhenAvailable=False).
      쉰 것 자체보다 **쉰 걸 아무도 모르는 것**이 문제다. 공백은 로그에 적어 둔다."""
    if not last_line:
        return ""
    m = re.match(r"\[(\d{2})-(\d{2}) (\d{2}):(\d{2})\]", last_line)
    if not m:
        return ""
    mo, d, hh, mm = (int(x) for x in m.groups())
    try:
        prev = now.replace(month=mo, day=d, hour=hh, minute=mm, second=0, microsecond=0)
    except ValueError:
        return ""
    if prev > now:                      # 해가 바뀌면 작년 기록이다
        prev = prev.replace(year=prev.year - 1)
    gap = (now - prev).total_seconds() / 60.0
    if gap <= expect_min + slack_min:
        return ""
    return "★ 워치독이 %.0f분 쉬었다(예상 %d분) — 그 사이 터널이 죽어도 아무도 못 고친다" % (gap, expect_min)


def last_log_line(path=None):
    try:
        lines = [l for l in open(path or LOG, encoding="utf-8").read().splitlines() if l.strip()]
        return lines[-1] if lines else ""
    except Exception:
        return ""


def pick_archive(version_files, keep=1):
    """[(버전, 경로)] → OLD로 옮길 경로 목록 (현재 정책은 최신 1개 제외)"""
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
        out = run_tree(
            ["powershell", "-NoProfile", "-Command",
             f"(Get-Process {name} -ErrorAction SilentlyContinue).Count"],
            timeout=20, drain_timeout=5,
        ).stdout.strip()
        return out.isdigit() and int(out) > 0
    except Exception:
        return False


def kill_by_cmdline(needle):
    """CommandLine에 needle 포함된 python 프로세스 종료 (PowerShell CIM — wmic 대체)"""
    ps = ("Get-CimInstance Win32_Process -Filter \"Name='python.exe' or Name='pythonw.exe'\" | "
          f"Where-Object {{ $_.CommandLine -like '*{needle}*' }} | "
          "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }")
    try:
        run_tree(["powershell", "-NoProfile", "-Command", ps],
                 timeout=30, drain_timeout=5)
    except Exception:
        pass


def start_hidden(script):
    subprocess.Popen([PYW if os.path.exists(PYW) else PY, os.path.join(ROOT, script)],
                     cwd=ROOT, creationflags=0x00000008)  # DETACHED_PROCESS


def heal_server(dry):
    if ping():
        return heal_stale_server(dry)
    if dry:
        return "서버 죽음(dry — 복구 생략)"
    kill_by_cmdline("app_server.py")
    start_hidden(os.path.join("webapp", "app_server.py"))
    time.sleep(4)
    return "서버 재시작 → " + ("성공" if ping() else "실패(다음 주기 재시도)")


def heal_stale_server(dry):
    """살아 있지만 **옛 코드로 도는** 서버를 스스로 새 코드로 올린다 (2026-08-08).

    ★ 죽은 서버는 원래 잡고 있었다. 그런데 **살아 있는데 옛 코드인 경우**는 아무도
      안 봤다 — 서버는 200 을 주고 화면도 숫자를 보여 주므로 정상으로 보인다.
      2026-08-08 하루에만 세 번 그 상태가 됐고(어제 20:48 서버가 하루치 변경을 통째로
      못 실은 채 돌았다), 그중 두 번은 **옆 세션이 파일을 고쳐서** 생겼다.
      사람이 고칠 때마다 기억해서 눌러야 하는 일은 결국 안 눌린다.

    ★ 안전장치 — 함부로 끄지 않는다:
      · 파일이 **1분 이상** 안정된 뒤에만 올린다. 편집 중간에 끄면 반쯤 저장된 코드로
        올라간다(그게 더 나쁘다).
      · 30분 주기 안에서 **한 번만** 시도한다. 실패가 반복되면 그냥 두고 인계에 남긴다
        — 계속 끄고 켜면 쓰는 사람이 아무것도 못 한다.
      · 재시작 뒤 응답을 확인한다. 못 살아나면 그 사실을 그대로 적는다.
    """
    try:
        sys.path.insert(0, os.path.join(ROOT, "webapp"))
        import restart_server
        s = restart_server.stale()
        if not s:
            return "서버 정상"
        _pid, _when, newer = s
        # 편집이 끝났는지 — 가장 최근에 바뀐 파일이 60초 넘게 조용해야 한다.
        newest = max((os.path.getmtime(os.path.join(ROOT, f))
                      for f in newer if os.path.isfile(os.path.join(ROOT, f))),
                     default=0)
        if time.time() - newest < 60:
            return "서버 옛코드 — 방금 편집 중이라 미룸(%s)" % ", ".join(newer[:2])
        if dry:
            return "서버 옛코드(dry — 재시작 생략): %s" % ", ".join(newer[:3])
        rc = restart_server.main([])
        ok = ping()
        return ("서버 옛코드 → 재시작 %s (%s)"
                % ("성공" if (rc == 0 and ok) else "실패(다음 주기 재시도)",
                   ", ".join(newer[:3])))
    except Exception as exc:
        return "서버 코드나이 확인 실패: %s" % str(exc)[:40]


def heal_band_evidence(dry):
    """'없음 확인' 근거가 **추월됐으면** 스스로 무효로 만든다 (2026-08-11 지시).

    ★ `[217]` 은 읽는 쪽 둘(수집 계획·인계 문서)이 추월된 근거를 **거르게** 했다.
      거르는 것만으로는 근거가 틀린 채로 남아, 바로잡는 길이 사람이 밴드 피드를 열어
      `real_latest.py --latest` 를 적어 주는 것뿐이었다. 그 한 줄이 마지막 사람 몫이었다.
    ★ **붙여넣기 파일보다 먼저** 온다. 파일에 어떤 번호를 담을지를 정하는 것이 이
      근거이므로, 순서가 뒤집히면 그 회차는 **틀린 근거로 만든 목록**을 사람 손에 쥐여
      준다 — 없는 번호 한 개가 21초다.
    ★ **밴드에 접속하지 않는다.** 캐시와 근거 파일만 읽고 근거 한 장만 고친다 —
      수집이 아니므로 코딩 세션 금지 규칙과 어긋나지 않는다(`heal_stale_pastefiles` 와 같다).
    """
    band_dir = os.path.join(ROOT, "band")
    try:
        if band_dir not in sys.path:
            sys.path.insert(0, band_dir)
        import real_latest as RL
        fixed = RL.heal(apply=not dry)
    except Exception as exc:                        # 근거 하나 때문에 회차를 세우지 않는다
        return "밴드 근거 확인 실패: %s" % str(exc)[:60]
    if not fixed:
        return "밴드 근거 정상"
    return "밴드 근거 추월 정정%s: %s" % (
        "(dry)" if dry else "",
        ", ".join("%s(%s→모름)" % (f["밴드"], f["이전"]) for f in fixed[:3]))


def heal_stale_pastefiles(dry):
    """붙여넣기 파일이 **옛 수집 JS**를 담고 있으면 다시 만든다 (2026-08-08).

    ★ 서버가 옛 코드로 도는 것과 **똑같은 종류의 조용한 사고**다. 수집 규칙을 고쳐도
      사람 손에 가는 것은 디스크에 있는 그 파일이다. 2026-08-08 에 댓글 수집을
      붙였는데 band/*_붙여넣기_*.js 네 개는 전부 그 이전에 만들어진 것이었다 —
      그대로 붙여넣었으면 **댓글이 한 건도 안 들어오는데 수집은 성공으로 끝났다.**
      개수도 날짜도 멀쩡해서 아무도 몰랐을 것이다.
    ★ 판단은 **파일 mtime 대 만드는 쪽 mtime** 이다(내용 비교가 아니다) —
      회차 번호는 매번 달라지므로 내용은 원래 다르다.
    ★ **'만드는 쪽'은 수집기만이 아니다** (2026-08-11). 예전에는 `grab_posts.js` 하나만
      봤는데, 파일에 **어떤 번호를 담을지**를 정하는 것은 `recheck_plan`·`make_oneclick`
      이다. 그래서 없는 번호 40개를 담던 규칙을 고쳐도 **디스크의 그 파일은 그대로**
      남았다 — 사람은 여전히 옛 목록을 붙여넣고 14분을 버린다. [162] 와 똑같은 모양이
      한 겹 위에서 반복된 것이다. 이제 셋 중 **가장 최근**을 기준으로 본다.
    ★ 이것은 **수집이 아니라 파일 만들기**다. 캐시를 읽기만 하고 밴드에 접속하지
      않는다 — 코딩 세션이 해도 되는 일이다(CLAUDE.md 의 수집 금지와 어긋나지 않는다).
    """
    import glob as _g
    band_dir = os.path.join(ROOT, "band")
    js = os.path.join(band_dir, "grab_posts.js")
    try:
        js_mt = os.path.getmtime(js)
    except OSError:
        return "붙여넣기 확인 생략(grab_posts.js 없음)"
    # ★ **번호를 고르는 쪽 mtime 은 make_oneclick 이 만든 파일에만 댄다** (2026-08-11).
    #   첫판은 모든 붙여넣기 파일에 댔는데, `댓글채우기_*` 는 comment_backfill 이,
    #   `재수집_*` 은 recollect 가 만든다. 남의 파일까지 '낡음'으로 몰면 여기서 지워지고
    #   **make_oneclick 은 그 파일을 다시 만들지 않는다** — 실측으로 두 개가 그렇게
    #   사라졌다(다음 09:50 회차 전까지 사람 손에 아무것도 안 남는다).
    #   낡음의 기준은 **그 파일을 만드는 쪽**이어야 한다.
    mk_mt = js_mt
    for maker in ("make_oneclick.py", "recheck_plan.py"):
        try:
            mk_mt = max(mk_mt, os.path.getmtime(os.path.join(band_dir, maker)))
        except OSError:
            pass                       # 없는 파일은 기준이 못 된다(있는 것만으로 판단)
    old = [p for p in _g.glob(os.path.join(band_dir, "*붙여넣기_*.js"))
           if os.path.getmtime(p) < (mk_mt if os.path.basename(p).startswith("수집_")
                                     else js_mt)]
    if not old:
        return "붙여넣기 파일 최신"
    names = ", ".join(os.path.basename(p) for p in old[:3])
    if dry:
        return "붙여넣기 옛 JS(dry — 생성 생략): %s" % names
    made, failed = 0, 0
    for p in old:
        base = os.path.basename(p)
        band = base.rsplit("_", 1)[-1][:-3]
        try:
            if not base.startswith("수집_"):
                # 남의 회차가 만든 파일이다 — 재수집은 08:00, 댓글채우기·댓글은 09:50
                # 회차가 **대상 번호를 정한다.** 여기서 대상까지 새로 고르면 그 회차의
                # 판단을 가로챈다(make_oneclick 을 불러 봤자 제 이름의 파일만 쓴다).
                # 그래서 **지운다.** 다음 회차가 새 JS 로 다시 만든다 —
                # 옛 JS 를 남겨 두는 것보다 없는 편이 낫다.
                os.unlink(p)
                made += 1
                continue
            run_tree([sys.executable, os.path.join(band_dir, "make_oneclick.py"),
                      "--band", band], cwd=ROOT, timeout=180, drain_timeout=10)
            # ★ **끝난 코드로 성공을 판단하지 않는다.** 훑을 것이 없는 밴드에서는
            #   생성기가 아무 파일도 안 쓰고 정상 종료한다(0). 그걸 성공으로 세면
            #   낡은 파일이 그대로 남은 채 "새로 만들었다"고 적힌다 — 거짓 보고다.
            #   실제로 파일이 새로워졌는지 mtime 으로 본다.
            if os.path.exists(p) and os.path.getmtime(p) >= js_mt:
                made += 1
            else:
                # 만들 것이 없다 = 이 파일은 있을 이유가 없다. 남겨 두면 사람이
                # 옛 JS 를 붙여넣는다. 지우면 필요할 때 다시 만들어진다.
                try:
                    os.unlink(p)
                    made += 1
                except OSError:
                    failed += 1
        except Exception:
            failed += 1
    return "붙여넣기 옛 JS %d개 → 새로 %d개%s (%s)" % (
        len(old), made, (" · 실패 %d" % failed) if failed else "", names)


def heal_fixed_funnel(dry):
    """Check the same public path a phone uses and refresh stale Funnel TLS."""
    try:
        from tailscale_serve import (
            FIXED_HOST, ensure_public_funnel, hostname, public_funnel_alive,
        )
        actual_host = hostname()
        if not actual_host:
            return "고정 Funnel 스킵 — Tailscale 로그인 없음"
        if actual_host != FIXED_HOST:
            return "고정 Funnel 주소 불일치 — 자동 주소변경 금지"
        if public_funnel_alive(FIXED_HOST):
            return "고정 Funnel 휴대폰 경로 정상"
        if dry:
            return "고정 Funnel 휴대폰 경로 죽음(dry)"
        ok, repaired = ensure_public_funnel(repair=True)
        if ok:
            return "고정 Funnel 공개경로 재등록 → 성공" if repaired else "고정 Funnel 정상"
        return "고정 Funnel 공개경로 재등록 → 실패(다음 주기 재시도)"
    except Exception as exc:
        return f"고정 Funnel 검사 오류: {str(exc)[:40]}"


def tunnel_alive(url):
    """프로세스 개수가 아니라 **실제 응답**으로 판정.
    (Get-Process는 1개일 때 Count가 비어 나와 살아있는 터널을 죽었다고 오판 →
     불필요한 재시작으로 공개 주소가 바뀌던 실사고가 있었다)"""
    if not url:
        return False
    try:
        from net_probe import probe
        return probe(url.rstrip("/") + "/api/ping", 12)[0]
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
    ps = ("Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'cloudflared.exe' -or "
          "$_.CommandLine -like '*tunnel_run*' } | ForEach-Object { "
          "Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue; $_.ProcessId }")
    try:
        r = run_tree(["powershell", "-NoProfile", "-Command", ps],
                     timeout=60, drain_timeout=10)
        n = len([x for x in (r.stdout or "").split() if x.strip().isdigit()])
        return f" (좀비 {n}개 정리)" if n else ""
    except Exception:
        return ""


def snapshot_handoff(dry):
    """세션 인계 스냅샷 — **세션이 갑자기 끊겨도 여기까지는 남는다.**

    종료 체크리스트는 끝낼 시간이 있을 때만 지켜진다. 컨텍스트가 차거나 크레딧이 끊기면
    그럴 기회가 없고, 그때 점유·큐·임시파일이 방치된 채 남는다. 워치독이 30분마다
    상태를 적어 두면 다음 세션은 최대 30분 전 상태에서 이어받는다.
    ★ 읽기 전용이다 — 무엇이 걸렸는지 적을 뿐, 점유를 풀거나 큐를 반영하지 않는다.
      상대 AI 가 일하는 중일 수 있어 함부로 가로채면 안 된다."""
    if dry:
        return "세션인계(dry)"
    try:
        r = run_tree([PY, os.path.join(ROOT, "session_handoff.py"), "--snapshot"],
                     cwd=ROOT, timeout=120, drain_timeout=10)
        line = [x for x in (r.stdout or "").splitlines() if "세션인계" in x]
        return line[-1] if line else "세션인계 갱신"
    except Exception as e:
        return f"세션인계 실패: {str(e)[:40]}"


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
        r = run_tree([PY, os.path.join(ROOT, "publish_endpoint.py")], cwd=ROOT,
                     timeout=120, drain_timeout=10)
        return (r.stdout or "").strip().splitlines()[-1][:60] if r.stdout.strip() else "게시 무응답"
    except Exception as e:
        return f"게시 오류: {str(e)[:40]}"


def sync_cloud_queue(dry):
    """PC 재가동 뒤 휴대폰 예약을 관리대장에 반영한다."""
    if dry:
        return "클라우드 예약 반영(dry)"
    try:
        r = run_tree(
            [PY, os.path.join(ROOT, "cloud_queue_sync.py")],
            cwd=ROOT,
            timeout=180,
            drain_timeout=10,
        )
        line = (r.stdout or r.stderr or "").strip().splitlines()
        return ("클라우드 예약 " + line[-1][:80]) if line else "클라우드 예약 확인"
    except Exception as e:
        return f"클라우드 예약 오류: {str(e)[:40]}"


def sync_uploads(dry):
    """단일 투입함을 30분마다 정본 분류하고, 새 원본이면 전체 대조를 한 번 깨운다."""
    if dry:
        return "업로드 투입함(dry)"
    try:
        r = run_tree(
            [PY, os.path.join(ROOT, "upload_intake.py"), "--apply"],
            cwd=ROOT, timeout=300, drain_timeout=15,
        )
        line = (r.stdout or r.stderr or "").strip().splitlines()
        summary = line[-1] if line else "업로드 분류 무응답"
        m = re.search(r"업로드 원본 분류:\s*(\d+)건", summary)
        moved = int(m.group(1)) if m else 0
        if r.returncode == 0 and moved:
            # daily_run의 프로세스 잠금이 이미 실행 중인 중복 기동을 안전하게 막는다.
            start_hidden("daily_run.py")
            return f"{summary} → 전체 대조 시작"
        return summary[:100]
    except Exception as e:
        return f"업로드 투입함 오류: {str(e)[:50]}"


def resume_deferred_apply(dry):
    """엑셀 열림으로 연기된 반영의 30분 안전망(2026-08-03 지시).

    닫힘 감시자(ledger_db --resume-watch)가 죽었어도 마커가 남아 있으면
    여기서 재개하거나 감시자를 다시 띄운다."""
    try:
        import ledger_db
        if not ledger_db._defer_state().get("slots"):
            return "연기 회차 없음"
        if dry:
            return "연기 회차 있음(dry — 재개 생략)"
        r = ledger_db.resume_check()
        return "연기 반영 점검 → " + str(r.get("상태") or r)
    except Exception as exc:
        return f"연기 반영 점검 실패({type(exc).__name__})"


def heal_autopilot(dry):
    """공용 자원 장애로 미뤄 둔 안전 작업을 30분마다 조금씩 이어 간다."""
    try:
        import autopilot
        result = autopilot.heal(limit=2, budget_seconds=600, dry=dry)
        actions = result.get("actions") or []
        if not actions:
            return "자율복구 대기 없음" if not result.get("active") else (
                "자율복구 %d건 대기(재시도 시각 전·인증 대기)" % result.get("active", 0))
        done = sum(1 for x in actions if x.get("result") == "done")
        return "자율복구 %d건 실행 · 완료 %d · 남음 %d" % (
            len(actions), done, result.get("active", 0))
    except Exception as exc:
        return "자율복구 점검 실패(%s)" % type(exc).__name__


def sync_worklog(dry):
    """일지 원본이 바뀌면 **기다리지 않고** 바로 대조해 큐에 넣는다 (2026-08-09 지시).

    사용자 지시: "돌발 AS 일지랑 정기점검 일지 엑셀과 앱에 자동 반영 ... 담당자 손댈 필요 없이".

    ★ 조각은 이미 다 있었다 — `upload_intake` 가 정본 자리로 옮기고, 09:50 회차가
      `work_log_sync --queue` 를 돌리고, 11:00·15:00 이 엑셀에 쓴다.
      빠진 것은 **속도**다: 오후에 올린 일지는 다음 날 09:50 까지 아무 데도 안 반영된다.
      그 사이 화면은 옛 숫자를 멀쩡히 보여 주므로 **올린 사람만 반영된 줄 안다.**
    ★ 파일이 안 바뀌었으면 아무것도 안 한다. 매 30분 대조를 돌리면 Z: 를 계속 훑어
      앱까지 같이 느려진다([168] · 사고 #29). 판정 근거는 **정본 폴더 최신 mtime** 하나다.
    """
    if dry:
        return "일지 감시(dry)"
    stamp = os.path.join(ROOT, "reports", ".worklog_seen.json")
    try:
        from source_dirs import work_log_dirs
        newest, newest_p = 0.0, ""
        for d in work_log_dirs():
            if not os.path.isdir(d):
                continue
            for f in glob.glob(os.path.join(d, "**", "*.xls*"), recursive=True):
                if os.path.basename(f).startswith("~$"):
                    continue
                m = os.path.getmtime(f)
                if m > newest:
                    newest, newest_p = m, f
        if not newest:
            return "일지 원본 없음"
        seen = 0.0
        if os.path.exists(stamp):
            try:
                seen = float(json.load(open(stamp, encoding="utf-8")).get("mtime") or 0)
            except Exception:
                seen = 0.0
        if newest <= seen:
            return "일지 그대로"
        r = run_tree([PY, os.path.join(ROOT, "work_log_sync.py"), "--queue"],
                     cwd=ROOT, timeout=900, drain_timeout=30, output_limit=120_000)
        # ★ 실패했으면 자국을 남기지 않는다 — 남기면 '봤다'가 되어 **다시는 안 본다.**
        #   그 일지는 영영 반영되지 않고 오류도 안 난다.
        if r.returncode != 0 or r.timed_out:
            why = "시간초과" if r.timed_out else "rc=%d" % r.returncode
            return "일지 대조 실패(%s) — 다음 회차에 다시 시도" % why
        with open(stamp, "w", encoding="utf-8", newline="") as f:
            json.dump({"mtime": newest, "파일": os.path.basename(newest_p),
                       "본때": datetime.now().isoformat(timespec="seconds")},
                      f, ensure_ascii=False)
        tail = [l for l in (r.stdout or "").strip().splitlines() if l.strip()]
        return "새 일지 반영 대기열 적재 · %s" % (tail[-1][:70] if tail else "요약 없음")
    except Exception as e:
        return "일지 감시 오류: %s" % str(e)[:50]


def run_incremental_pipeline(dry):
    """Five-minute task의 안전망 — 새 원본이면 앱 DB와 보관본까지 끝낸다.

    별도 작업 스케줄러가 주 경로다. 워치독에도 맨 앞에 두는 이유는 스케줄러가
    비활성화되거나 PC가 예약 시각에 꺼져 있었어도 다음 30분 회차에서 복구하기
    위해서다. 파이프라인 자체 PID 잠금과 지문 멱등성이 중복 실행을 막는다.
    """
    if dry:
        return "증분 자동화(dry)"
    result = run_tree(
        [PY, os.path.join(ROOT, "automation_pipeline.py"), "--once", "--trigger", "watchdog"],
        cwd=ROOT,
        timeout=1500,
        drain_timeout=30,
        output_limit=100_000,
    )
    if result.timed_out:
        return "증분 자동화 시간초과 — 자식나무 정리·다음 회차 재시도"
    if result.returncode != 0:
        tail = [line for line in (result.stdout or "").splitlines() if line.strip()][-1:]
        return "증분 자동화 실패(rc=%d)%s" % (
            result.returncode,
            (" · " + tail[0][:80]) if tail else "",
        )
    return "증분 자동화 완료"


def resume_parked(dry):
    """세워 둔 일이 풀렸으면 알리고(아무도 없으면 AI 에게 넘기고), 보류된 푸시를 민다.

    ★ 2026-08-11 지시 "이 세션도 완전 자동화시켜". 그날 실측 둘 — 옆 세션이 `code` 를
      놓았는데 **아무 화면도 "이제 된다"고 말하지 않아** 사람이 "하던 작업 진행" 을 두 번
      쳤고, 자동 마무리가 보류한 푸시는 그 세션이 사라진 뒤에도 **미는 사람이 없었다.**
      판단은 `worksplit_auto` 한 곳이 한다 — 여기서 다시 세면 두 벌이 된다.
    """
    try:
        import worksplit_auto
        return worksplit_auto.run(dry=dry)["한줄"]
    except Exception as exc:
        return "세션자동화 실패: %s: %s" % (type(exc).__name__, str(exc)[:80])


def main():
    # 류지영 매니저 입력 중에는 로그 파일조차 갱신하지 않고 즉시 종료한다.
    if is_input_window():
        print(f"입력 보호시간({input_window_label()}) — 워치독 무동작 종료")
        return
    dry = "--dry" in sys.argv
    # 원장 버전 정리는 daily_run의 ledger_versions.py 한 곳에서만 수행한다.
    # 워치독이 낮은 버전 포크를 OLD로 옮겨 증거를 숨기는 일을 막는다.
    gap = gap_note(last_log_line(), datetime.now())     # 기록은 healing 전에 읽는다
    results = [run_incremental_pipeline(dry), sync_uploads(dry), sync_worklog(dry),
               sync_cloud_queue(dry), heal_server(dry), heal_fixed_funnel(dry),
               # ★ 근거 정정이 **붙여넣기 파일 만들기보다 먼저**다 — 목록에 담을 번호를
               #   정하는 것이 그 근거다(2026-08-11, `[223]`).
               heal_band_evidence(dry),
               heal_stale_pastefiles(dry),
               heal_autopilot(dry),
               resume_parked(dry),
               heal_tunnel(dry), publish_endpoint(dry), clean_reports(dry),
               snapshot_handoff(dry), resume_deferred_apply(dry)]
    if gap:
        results.insert(0, gap)
    log(" | ".join(results))


if __name__ == "__main__":
    main()
