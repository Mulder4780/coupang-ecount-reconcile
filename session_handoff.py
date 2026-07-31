# -*- coding: utf-8 -*-
"""session_handoff.py — 세션이 갑자기 끊겨도 다음 세션이 그대로 이어받게.

사용자 지시(2026-07-29): "세션 컨텍스트가 다 차도 다음 세션에서 문제없이 구동되게."

왜 필요한가
  종료 체크리스트는 **끝낼 시간이 있을 때만** 지켜진다. 컨텍스트가 갑자기 차거나
  크레딧이 끊기면 그걸 할 기회가 없다. 그때 남는 것들:
    · 점유(ai_claim)가 잡힌 채 방치 — 다음 세션이 원장을 못 고친다
    · 입력 큐에 적재만 되고 반영 안 된 셀
    · vN+1 을 만들다 만 `.tmp.xlsx`
    · 커밋 안 된 변경 / 푸시 안 된 커밋
  이걸 **사람이 기억해서** 넘기게 하면 안 된다. 파일이 기억해야 한다.

두 가지 모드
  --snapshot : 지금 상태를 `reports/세션인계.md|json` 에 남긴다.
               워치독이 30분마다 부른다 → 세션이 언제 죽든 최대 30분 전 상태가 남는다.
  --check    : 새 세션이 시작할 때 읽는다. **막힌 것부터** 보여주고 그 다음 할 일을 준다.
               (CLAUDE.md 시작 체크리스트 0번)

★ 이 도구는 아무것도 고치지 않는다 — 무엇이 걸려 있는지 알려 주고 명령을 제시할 뿐이다.
  자동으로 점유를 풀거나 큐를 반영하면, 상대 AI가 일하는 중인데 가로채게 된다.
"""
import argparse
import glob
import json
import os
import subprocess
import sys
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(BASE, "reports")
CHECKPOINT_PATH = os.path.join(REPORT_DIR, "진행체크포인트.json")
STALE_MIN = 45          # ai_claim 자동 해제 기준과 같게 — 넘으면 죽은 세션의 잔재로 본다

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def git(*args):
    try:
        r = subprocess.run(["git"] + list(args), cwd=BASE, capture_output=True,
                           text=True, encoding="utf-8", errors="replace", timeout=60)
        return (r.stdout or "").strip()
    except Exception:
        return ""


def queue_left():
    p = os.path.join(BASE, "updates", "pending_updates.json")
    try:
        return len(json.load(open(p, encoding="utf-8")))
    except Exception:
        return 0


def temp_files():
    """vN+1 을 만들다 만 흔적. 남아 있으면 그 작업이 중간에 끊긴 것이다."""
    try:
        from ecount_reconcile import load_config, resolve_master
        folder = os.path.dirname(resolve_master(load_config()["reconcile"]["master_xlsx"]))
        return [os.path.basename(p) for p in glob.glob(os.path.join(folder, "*.tmp.xlsx"))]
    except Exception:
        return []


def pid_alive(pid):
    """그 프로세스가 아직 살아 있나. **시간보다 확실한 판정이다** —
    세션이 죽으면 45분을 기다릴 것 없이 그 자리에서 잔재로 볼 수 있다.
    판정이 안 되면 None 을 돌려 시간 기준으로 넘긴다(모르면 함부로 죽었다고 하지 않는다)."""
    if not pid:
        return None
    try:
        import ctypes
        h = ctypes.windll.kernel32.OpenProcess(0x1000, False, int(pid))  # QUERY_LIMITED_INFO
        if not h:
            return False
        ctypes.windll.kernel32.CloseHandle(h)
        return True
    except Exception:
        return None


def claims():
    try:
        import ai_claim
        data = ai_claim.load() or {}
    except Exception:
        return []
    import time as _t
    out = []
    for lock, info in data.items():
        if not isinstance(info, dict) or not info.get("who"):
            continue
        # ai_claim 은 `at` 을 **에포크 초(float)** 로 적는다. ISO 문자열이 아니다.
        mins = None
        try:
            at = float(info.get("at") or 0)
            if at > 0:
                mins = int((_t.time() - at) // 60)
        except (TypeError, ValueError):
            pass
        alive = pid_alive(info.get("pid"))
        stale = (alive is False) or (alive is not True and mins is not None and mins >= STALE_MIN)
        out.append({"lock": lock, "who": info.get("who"), "why": info.get("why", ""),
                    "mins": mins, "pid": info.get("pid"), "alive": alive, "stale": stale})
    return out


def ledger():
    try:
        from workbook_patch import latest_master
        path, ver = latest_master()
        st = os.stat(path)
        return {"버전": ver, "파일": os.path.basename(path),
                "수정": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M")}
    except Exception as e:
        return {"오류": str(e)[:60]}


def next_tasks():
    """AGENTS.md 의 '다음 세션이 이어서 할 일' 을 그대로 가져온다 — 두 곳에 적지 않는다."""
    try:
        text = open(os.path.join(BASE, "AGENTS.md"), encoding="utf-8").read()
    except Exception:
        return []
    lines, on = [], False
    for ln in text.splitlines():
        if ln.startswith("## 다음 세션이 이어서 할 일"):
            on = True
            continue
        if on and ln.startswith("## "):
            break
        if on and ln.strip():
            lines.append(ln.rstrip())
    return lines[:24]


def read_checkpoint():
    """대화 컨텍스트와 무관하게 현재 작업의 정확한 재개 지점을 읽는다."""
    try:
        value = json.load(open(CHECKPOINT_PATH, encoding="utf-8"))
        return value if isinstance(value, dict) and value.get("상태") == "진행중" else {}
    except Exception:
        return {}


def write_checkpoint(objective="", done=None, pending=None, notes=None):
    """현재 작업을 원자적으로 저장한다.

    같은 세션에서 여러 번 호출하면 전달된 항목만 갱신한다. `--pending`을 새로
    주면 남은 작업 목록을 교체하고, `--done`·`--note`는 기존 내용 뒤에 보탠다.
    """
    old = read_checkpoint()
    value = {
        "상태": "진행중",
        "갱신시각": datetime.now().isoformat(timespec="seconds"),
        "목표": objective or old.get("목표", ""),
        "완료": list(old.get("완료", [])),
        "남은작업": list(old.get("남은작업", [])),
        "메모": list(old.get("메모", [])),
        "기준커밋": git("rev-parse", "HEAD"),
        "작업트리": [l for l in git("status", "--short").splitlines() if l.strip()],
    }
    if pending is not None:
        value["남은작업"] = [x for x in pending if x]
    for key, items in (("완료", done or []), ("메모", notes or [])):
        for item in items:
            if item and item not in value[key]:
                value[key].append(item)
    os.makedirs(REPORT_DIR, exist_ok=True)
    temp = CHECKPOINT_PATH + ".tmp"
    with open(temp, "w", encoding="utf-8") as fh:
        json.dump(value, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    os.replace(temp, CHECKPOINT_PATH)
    return value


def rule_copies():
    """루트 CLAUDE.md·AGENTS.md 는 ecount/CLAUDE.md(정본)의 사본이어야 한다.

    루트 파일은 git 밖이라 조용히 낡는다 — 실제로 2026-07-31 루트 CLAUDE.md 가
    '엑셀 두 번' 규칙 이전 판으로 남아 옛 절차(ledger_writer --apply)를 지시하고 있었다.
    다르면 옛 규칙을 읽는 세션이 생기므로 '먼저 처리할 것'으로 올린다."""
    try:
        master = open(os.path.join(BASE, "CLAUDE.md"), encoding="utf-8").read()
    except OSError:
        return []
    out = []
    for name in ("CLAUDE.md", "AGENTS.md"):
        p = os.path.join(os.path.dirname(BASE), name)
        try:
            if open(p, encoding="utf-8").read() != master:
                out.append(name)
        except OSError:
            out.append(name)
    return out


def collect():
    unstaged = [l for l in git("status", "--short").splitlines() if l.strip()]
    unpushed = [l for l in git("log", "origin/master..HEAD", "--oneline").splitlines() if l.strip()]
    return {
        "시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "원장": ledger(),
        "큐잔량": queue_left(),
        "임시파일": temp_files(),
        "점유": claims(),
        "미커밋": unstaged,
        "미푸시": unpushed,
        "최근커밋": [l for l in git("log", "-5", "--oneline").splitlines() if l.strip()],
        "다음할일": next_tasks(),
        "진행체크포인트": read_checkpoint(),
        "지시문사본": rule_copies(),
    }


def blockers(st, for_sol=False):
    """다음 세션이 **먼저 처리해야** 하는 것 — 안 하면 조용히 어긋난다."""
    out = []
    if st["큐잔량"]:
        out.append(("입력 큐에 %d건이 반영되지 않았다" % st["큐잔량"],
                    "python ledger_db.py --intake  # Excel은 다음 11:00·15:00 회차"))
    if st["임시파일"]:
        out.append(("원장 임시파일이 남았다(만들다 끊김): %s" % ", ".join(st["임시파일"][:3]),
                    "내용 확인 후 삭제 — 정식 vN+1 로 승격되지 않은 파일이다"))
    for c in st["점유"]:
        m = c["mins"] if c["mins"] is not None else "?"
        if c["stale"]:
            why = "프로세스 %s 가 없다" % c.get("pid") if c.get("alive") is False else "%s분 경과" % m
            out.append(("★ '%s' 점유가 **죽은 세션의 잔재**로 보인다 — %s (%s · %s)"
                        % (c["lock"], why, c["who"], c["why"]),
                        "python ai_claim.py --who %s --free %s" % (c["who"], c["lock"])))
        else:
            out.append(("%s 가 '%s' 를 잡고 있다(%s분, 살아 있음) — 배타 작업은 피할 것"
                        % (c["who"], c["lock"], m),
                        "조회·분석으로 돌리거나 상대가 놓을 때까지 기다린다"))
    if st["미푸시"]:
        out.append(("푸시되지 않은 커밋 %d개" % len(st["미푸시"]),
                    "git pull --rebase && git push  (비밀 스캔 후)"))
    if st.get("지시문사본"):
        out.append(("루트 %s 가 정본(ecount/CLAUDE.md)과 다르다 — 옛 규칙을 읽는 세션이 생긴다"
                    % "·".join(st["지시문사본"]),
                    "내용 비교 → 정본에 반영 후 복사: python -c \"import shutil;"
                    "[shutil.copy('ecount/CLAUDE.md',d) for d in ('CLAUDE.md','AGENTS.md')]\""))
    if for_sol and st.get("terra_sol_review", {}).get("pending"):
        out.append(("Terra 작업분의 Sol 사전 검토가 필요합니다 (%s)" %
                    st["terra_sol_review"].get("reason", "검토 미완료"),
                    "python handoff_review.py --review-sol"))
    return out


def to_md(st, for_sol=False):
    L = ["# 세션 인계 — 지금 어디까지 됐나", "",
         "- 기준: %s · 관리대장 **v%s**(%s)" % (st["시각"], st["원장"].get("버전", "?"),
                                               st["원장"].get("수정", "")),
         "- 이 문서는 워치독이 30분마다 갱신한다. **세션이 갑자기 끊겨도 여기까지는 남는다.**", ""]
    cp = st.get("진행체크포인트") or {}
    if cp:
        L += ["## ★ 진행 중 작업 — 여기서 바로 재개", "",
              "- 목표: **%s**" % (cp.get("목표") or "미기입"),
              "- 체크포인트: %s · 기준 커밋 `%s`" %
              (cp.get("갱신시각", "?"), str(cp.get("기준커밋", "?"))[:12]), ""]
        if cp.get("완료"):
            L += ["### 완료", *["- %s" % x for x in cp["완료"]], ""]
        if cp.get("남은작업"):
            L += ["### 다음 순서", *["%d. %s" % (i + 1, x)
                                    for i, x in enumerate(cp["남은작업"])], ""]
        if cp.get("메모"):
            L += ["### 재개 메모", *["- %s" % x for x in cp["메모"]], ""]
    bl = blockers(st, for_sol=for_sol)
    L += ["## 먼저 처리할 것 (%d)" % len(bl), ""]
    if not bl:
        L.append("걸린 것 없음 — 바로 새 작업을 시작해도 된다.")
    for why, how in bl:
        L += ["- **%s**" % why, "  - `%s`" % how]
    L.append("")
    if st["미커밋"]:
        L += ["## 커밋되지 않은 변경 (%d)" % len(st["미커밋"]), "",
              "```", *st["미커밋"][:20], "```", "",
              "> 상대 AI 가 작업 중일 수 있다. **내 파일만** 커밋할 것.", ""]
    L += ["## 최근 커밋", "", "```", *st["최근커밋"], "```", ""]
    if st["다음할일"]:
        L += ["## 다음 할 일 (AGENTS.md 발췌)", "", *st["다음할일"], ""]
    L += ["## 시작 절차", "",
          "1. `python session_handoff.py --check` (이 문서)",
          "2. `ecount/AGENTS.md` 읽기 — 절대규칙·현재 상태",
          "3. `python tests/synthetic_check.py` → ALL GREEN 확인 후 실작업",
          "4. `python data_status.py --print` — 자료가 지금 얼마나 있나", ""]
    if for_sol:
        L += ["", "## Sol 전용 시작 관문", "",
              "`python handoff_review.py --review-sol`이 PASS 하기 전에는 쓰기 작업을 시작하지 않습니다."]
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--for-sol", action="store_true", help="Terra -> Sol 검토 관문도 함께 확인")
    ap.add_argument("--snapshot", action="store_true", help="상태를 파일로 남긴다(워치독용)")
    ap.add_argument("--check", action="store_true", help="새 세션 시작 시 읽는다")
    ap.add_argument("--checkpoint", action="store_true", help="현재 작업 재개 지점을 저장한다")
    ap.add_argument("--objective", default="", help="현재 작업 목표")
    ap.add_argument("--done", action="append", default=[], help="완료 항목(여러 번 사용 가능)")
    ap.add_argument("--pending", action="append", help="남은 작업(지정 시 기존 목록 교체)")
    ap.add_argument("--note", action="append", default=[], help="재개 메모(여러 번 사용 가능)")
    ap.add_argument("--clear-checkpoint", action="store_true", help="현재 작업이 끝났을 때 체크포인트를 닫는다")
    a = ap.parse_args()
    if a.clear_checkpoint:
        try:
            value = json.load(open(CHECKPOINT_PATH, encoding="utf-8"))
        except Exception:
            value = {}
        value.update({"상태": "완료", "완료시각": datetime.now().isoformat(timespec="seconds")})
        os.makedirs(REPORT_DIR, exist_ok=True)
        with open(CHECKPOINT_PATH, "w", encoding="utf-8") as fh:
            json.dump(value, fh, ensure_ascii=False, indent=2)
            fh.write("\n")
        print("진행 체크포인트 완료 처리")
        return 0
    if a.checkpoint:
        value = write_checkpoint(a.objective, a.done, a.pending, a.note)
        print("진행 체크포인트 갱신 — 남은 작업 %d건" % len(value.get("남은작업", [])))
        return 0
    st = collect()
    if a.for_sol:
        try:
            import handoff_review
            st["terra_sol_review"] = handoff_review.review_state()
        except Exception as exc:
            st["terra_sol_review"] = {"pending": True, "reason": "검토 상태 확인 실패: %s" % exc}
    md = to_md(st, for_sol=a.for_sol)
    if a.snapshot:
        os.makedirs(REPORT_DIR, exist_ok=True)
        open(os.path.join(REPORT_DIR, "세션인계.md"), "w", encoding="utf-8").write(md)
        json.dump(st, open(os.path.join(REPORT_DIR, "세션인계.json"), "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1, default=str)
        bl = blockers(st, for_sol=a.for_sol)
        print("세션인계 갱신 — 걸린 것 %d건 · 관리대장 v%s" % (len(bl), st["원장"].get("버전", "?")))
        return 0
    print(md)
    return 0


if __name__ == "__main__":
    sys.exit(main())
