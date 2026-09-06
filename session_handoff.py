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

  --adopt    : **다른 계정·다른 창이 이어받을 때** 한 번 부른다(2026-08-06 지시:
               "이 세션은 완료되면 다른 계정으로 로그인해서 사용할거야, 그때 아무 문제
               없이 처리될 수 있는 알고리즘 구성해"). 기계가 **확실히 판단할 수 있는
               것만** 스스로 풀고, 사람 몫만 남겨 보여 준다. 자세한 것은 adopt() 참조.

★ --snapshot·--check 는 아무것도 고치지 않는다 — 무엇이 걸려 있는지 알려 주고
  명령을 제시할 뿐이다. 자동으로 점유를 풀거나 큐를 반영하면, 상대 AI가 일하는
  중인데 가로채게 된다. 고치는 것은 --adopt 하나뿐이고, 거기서도 **주인 세션이
  죽었다는 증거(pid)** 가 있을 때만 손댄다.
"""
import argparse
import glob
import json
import os
import re
import subprocess
import sys
from datetime import datetime, timedelta

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(BASE, "reports")
CHECKPOINT_PATH = os.path.join(REPORT_DIR, "진행체크포인트.json")
STALE_MIN = 45          # ai_claim 자동 해제 기준과 같게 — 넘으면 죽은 세션의 잔재로 본다

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _main_root():
    """공용 상태의 주인(본체 체크아웃). 워크트리에서 돌 때만 BASE 와 달라진다.

    worktree_state 를 못 읽어도 죽지 않는다 — 그때는 예전처럼 제 폴더를 쓴다."""
    try:
        from worktree_state import main_root
        return main_root()
    except Exception:
        return BASE


def _worktree_state():
    """워크트리가 본체 상태와 이어져 있나. 본체에서 돌면 None(볼 것이 없다)."""
    try:
        import worktree_state
        if not worktree_state.is_worktree():
            return None
        return worktree_state.status()
    except Exception:
        return None


def git(*args):
    try:
        r = subprocess.run(["git"] + list(args), cwd=BASE, capture_output=True,
                           text=True, encoding="utf-8", errors="replace", timeout=60, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
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


def stranded_editor():
    """**사람이 옛 버전을 열어 놓고 편집 중인가.** (2026-08-07 실측)

    엑셀은 파일을 열면 옆에 `~$이름.xlsx` 잠금파일을 만든다. 그 잠금이 최신본이 아닌
    **옛 vN** 에 걸려 있으면 조용한 손실이 시작된 것이다 — 그 사람이 v538 에 손으로
    적어 넣는 동안 회차가 v539·v540·v541 을 만들고, 다음 회차는 v541 에서 이어간다.
    **v538 에 적은 것은 어디에도 넘어가지 않는다.** 파일이 열려 있으니 오류도 안 난다.

    이건 "하루 두 번만 반영" 규칙이 줄여 주는 문제지, 없애 주는 문제가 아니다.
    사람에게 알리는 것 말고 기계가 할 수 있는 일은 없다 — 남의 엑셀을 닫지 않는다.
    """
    return _editor_locks()[0]


def _editor_locks():
    """(옛버전 잠금, 최신본 잠금) — 한 번 훑어 둘 다 준다."""
    try:
        from ecount_reconcile import load_config, resolve_master
        cur = resolve_master(load_config()["reconcile"]["master_xlsx"])
        folder, latest = os.path.dirname(cur), os.path.basename(cur)
        old, new = [], []
        for p in glob.glob(os.path.join(folder, "~$*.xlsx")):
            target = os.path.basename(p)[2:]          # `~$` 를 뗀 것이 열려 있는 파일
            (new if target == latest else old).append(target)
        return sorted(old), sorted(new)
    except Exception:
        return [], []


def latest_viewer():
    """**최신본을 열어 둔 사람이 있는가** (2026-08-11 지시 — 엑셀 손입력 종료).

    예전에는 '최신본을 보고 있는 건 정상'이라 조용히 넘겼다. 앱 전용 입력 뒤에는
    뜻이 하나 늘었다 — 열람은 여전히 정당하지만, **거기 적은 값은 정본에 안 들어간다.**
    그래서 경보가 아니라 안내로 띄운다(경보로 올리면 열람마다 울려 아무도 안 본다)."""
    return _editor_locks()[1]


HAND_EDIT_LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "reports", "엑셀_손입력_감지.json")


def _naive_iso(ts):
    """어느 시계로 적힌 시각이든 **이 PC 시각**(타임존 없음)으로 옮긴다.

    ★ 판정은 `error_book.to_local` **한 곳**을 빌린다([162]) — 여기서 다시 적으면
      같은 물음에 두 답이 생기고, 갈린 뒤에는 어느 쪽이 맞는지 아무도 모른다.
    ★ 못 빌리면 **원본 그대로** 돌려준다 — 부르는 쪽이 `ValueError`·`TypeError` 를
      받아 그 기록만 건너뛴다.  인계 한 장을 통째로 죽이는 것보다 낫다([169]).
    """
    try:
        import error_book
        return error_book.to_local(ts)
    except Exception:                            # noqa: BLE001
        return str(ts or "")


def hand_edit_signal():
    """손입력 감지 기록의 싼 요약([168] — 여기서 해시 계산 금지, 읽기만).

    쓰는 쪽은 둘이다: realtime_monitor(내용 변경)·ledger_db.human_editing(열림).
    여기는 마지막 항목과 24시간 안 건수만 읽어 인계 문서에 올린다.

    ★ **그 두 손이 시각을 서로 다른 모양으로 적는다**(2026-08-27 실사고).
      realtime_monitor 는 `korea_now()`(타임존 있음), ledger_db 는
      `datetime.now()`(없음)다.  예전에는 타임존 없는 것만 받아서, **손입력이
      처음 감지된 그날** 이 함수가 `TypeError` 로 죽고 **인계 문서가 통째로
      안 나왔다** — 알리려던 기능이 화면을 없앤 셈이다([169]).
      그리고 그 자리는 `daily_run` 의 **0단계**(관문)까지 죽였다.
    ⚠ `except ValueError` 로는 못 받는다 — 그것은 **비교**에서 나는 오류다.
      파싱은 성공하고 `>=` 에서 터진다.
    """
    try:
        rows = json.load(open(HAND_EDIT_LOG, encoding="utf-8"))
        if not isinstance(rows, list) or not rows:
            return None
        # ★ **쓰는 쪽은 갈랐는데 읽는 쪽이 안 갈랐다** (2026-08-28 실사고).
        #   [475] 가 realtime_monitor 를 고쳐 '값은 그대로고 수식만 바뀐 것'을
        #   `기계도구(경보아님)` 으로 따로 적게 했다 — 그런데 여기는 **24시간 안
        #   모든 줄**을 세어 그것까지 경보로 올렸다.  그러면 `expand_rows` 가 행을
        #   늘릴 때마다 인계 맨 위에 거짓 경보가 서고, 그 조치는 사람을
        #   *"넣을 값이 없는데 앱에 다시 넣어라"* 로 보낸다([172]).
        #   그리고 거짓 경보는 **진짜 손입력을 덮는다**([170]).
        # ★ 내리는 근거는 **지어낸 것이 아니라 적힌 것** 둘뿐이다:
        #   ① 기계가 스스로 그렇게 적은 갈래 ② 사람이 재 보고 붙인 `정정`([475]).
        #   판정 문장을 뜯어 짐작하지 않는다 — 넓히면 진짜 손입력이 조용히 사라진다.
        try:
            import realtime_monitor
            tool_kind = realtime_monitor.HAND_EDIT_TOOL_KIND
        except Exception:                          # noqa: BLE001
            tool_kind = ""      # 못 읽으면 **아무것도 안 뺀다** — 예전 그대로 경보다([169])
        cut = datetime.now() - timedelta(hours=24)
        fresh, dropped, last = 0, 0, None
        for r in rows:
            try:
                if datetime.fromisoformat(_naive_iso(r.get("시각", ""))) < cut:
                    continue
            except (ValueError, TypeError):
                continue
            if (tool_kind and r.get("종류") == tool_kind) or r.get("정정"):
                dropped += 1
                continue
            fresh += 1
            last = r          # ⚠ 마지막 **경보 대상**이다 — 기계 줄이 뒤에 와도 안 밀린다
        return {"최근24h": fresh, "마지막": last, "뺀건수": dropped} if fresh else None
    except (OSError, ValueError):
        return None


def pid_alive(pid, born_before=None, pid_started_at=None):
    """그 프로세스가 아직 살아 있나. **시간보다 확실한 판정이다** —
    세션이 죽으면 45분을 기다릴 것 없이 그 자리에서 잔재로 볼 수 있다.
    판정이 안 되면 None 을 돌려 시간 기준으로 넘긴다(모르면 함부로 죽었다고 하지 않는다).
    `born_before`(에포크 초)를 주면 그 시각 뒤에 태어난 프로세스를 남으로 본다 —
    pid 재사용 오판 방지(2026-08-11 실사고 · 검증 [210])."""
    if not pid:
        return None
    # ★ 판정은 pid_alive.py 한 곳에서 한다 (2026-08-06 실사고 · 검증 [121]).
    #   여기 있던 옛 판정은 윈도우에서 **끝난 프로세스도 살아 있다**고 했다 —
    #   OpenProcess 는 종료된 프로세스에도 핸들을 준다. 죽은 세션의 점유가
    #   영원히 안 풀리는 쪽으로 틀렸다.
    try:
        import pid_alive
        return pid_alive.owner_alive(
            pid, pid_started_at=pid_started_at, born_before=born_before)
    except Exception:
        return None


def _owner_is_python(pid):
    """회차·자국·잠금의 주인은 언제나 **python 프로세스**다 — 이름이 읽히는데
    python 이 아니면 번호만 물려받은 남이다(pid 재사용 · 2026-08-11 실사고 두 번째).
    생성시각(born_before)이 못 가르는 경우를 마저 가른다: 재사용한 프로세스가
    기록 시각보다 **먼저** 떠 있던 상주 서비스면 시각 판정은 통과해 버린다.
    못 읽으면 None — 이름만으로 산 주인을 죽었다고 하지 않는다. 검증 [211]."""
    try:
        import pid_alive as _pa
        name = _pa.image_name(pid)
    except Exception:
        return None
    if not name:
        return None
    return "python" in name


def _is_mine(info):
    """내 세션 것인가 — 안내 문구를 고르는 데 쓴다(남의 것에 --free 를 권하면 안 된다)."""
    try:
        import ai_claim
        return bool(ai_claim._is_mine(info, info.get("who") or "claude"))
    except Exception:
        return False


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
        born = None       # 점유를 적은 시각 — 주인은 그 전에 떠 있었어야 한다([210])
        try:
            at = float(info.get("at") or 0)
            if at > 0:
                mins = int((_t.time() - at) // 60)
                born = at
        except (TypeError, ValueError):
            pass
        # ★ 주인은 `pid` 가 아니라 **`agent_pid`** 다 (2026-08-06 실사고).
        #   `pid` 는 ai_claim 을 실행한 CLI 프로세스라 명령이 끝나는 즉시 죽는다 —
        #   그것으로 판정하면 **살아 있는 옆 세션의 점유까지 '죽은 잔재'로** 표시하고,
        #   "이 명령으로 푸세요" 라고 안내한다. 실제로 그 안내대로 하면 ai_claim 이
        #   거부하므로(남의 것) 사람은 영문도 모른 채 막힌다. 판정은 한 벌이어야 한다.
        use_agent = bool(info.get("agent_pid"))
        owner_pid = info.get("agent_pid") or info.get("pid")
        owner_started_at = (info.get("agent_pid_started_at") if use_agent
                            else info.get("pid_started_at"))
        alive = pid_alive(owner_pid, born_before=born,
                          pid_started_at=owner_started_at)
        stale = (alive is False) or (alive is not True and mins is not None and mins >= STALE_MIN)
        out.append({"lock": lock, "who": info.get("who"), "why": info.get("why", ""),
                    "mins": mins, "pid": owner_pid, "alive": alive, "stale": stale,
                    "mine": _is_mine(info)})
    return out


def ledger():
    try:
        from workbook_patch import latest_master
        path, ver = latest_master()
        st = os.stat(path)
        return {"버전": ver, "파일": os.path.basename(path),
                "수정": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M")}
    except (Exception, SystemExit) as e:
        # ★ `SystemExit` 는 `Exception` **이 아니다**.  `latest_master()` 가
        #   그것을 던지므로(workbook_patch:64 · [467] 이 일부러 그렇게 뒀다)
        #   예전 `except Exception` 은 **있는 것처럼 보이는데 실제로는 없었다**
        #   ([212] 가 시각 비교에서 겪은 것과 같은 모양).
        # ★ 그래서 Z: 가 끊긴 날 **인계 한 장이 통째로 없어졌다**(2026-09-06
        #   실측: `--check` 가 한 줄짜리 오류만 냈다).  인계는 세션의 **첫
        #   명령**이라 그때가 가장 필요한 순간이다 —
        #   **덜 세는 것은 견딜 수 있고 화면이 없어지는 것은 못 견딘다**([212]).
        # ★ 조용히 넘기지 않는다([169]) — `버전` 자리에 그렇게 적어
        #   **읽는 쪽을 한 글자도 안 고치고** 말한다([172]).
        return {"오류": str(e)[:60], "버전": "? (못 읽음)"}


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
    # ★ 루트 사본은 **본체 체크아웃 위**에만 있다. 워크트리에서 돌 때
    #   `dirname(BASE)` 를 보면 `.claude/worktrees/CLAUDE.md` 를 찾아 "정본과 다르다"는
    #   거짓 경보가 매번 '먼저 처리할 것' 맨 위에 떴다(실측: 해시는 같았다).
    root = os.path.dirname(_main_root())
    for name in ("CLAUDE.md", "AGENTS.md"):
        p = os.path.join(root, name)
        try:
            if open(p, encoding="utf-8").read() != master:
                out.append(name)
        except OSError:
            out.append(name)
    return out


# 며칠 밀리면 '밀린 것'으로 볼까. 기준은 **대표 보고가 무엇을 필요로 하나** 다.
# 10:30 보고는 늘 '어제'를 말하므로, 밴드에 어제 글이 없으면 그 보고는 이미 틀렸다
# → 밴드는 1일까지만 봐준다(밀린일 2면 어제가 비었다는 뜻). 카톡은 사람이 내보내야 해
# 하루 늦을 수 있고, ERP 는 주 단위 화면이라 더 길게 둔다.
FRESH_LIMIT = {"밴드": 1, "카카오톡": 2, "ERP 내보내기": 7}


INDEX_CACHE = os.path.join(BASE, "db", "source_index_cache.json")


def _index_newest_day(marker):
    """원본 색인 캐시에서 경로에 `marker` 가 든 파일의 **가장 최근 수정일**.

    ★ Z: 를 직접 훑지 않는다. `--check` 는 매 세션의 첫 명령인데, 네트워크 드라이브
      전수 스캔은 수 분이 걸린다 — 느린 점검은 결국 아무도 안 돌리게 된다.
      캐시 키가 `경로|크기|mtime` 이라 파일을 열 필요도 없다.
    캐시가 낡았으면 실제보다 **오래돼 보인다** — 안전한 쪽으로 틀린다(수집하라고 말한다).
    """
    best = 0
    try:
        keys = json.load(open(INDEX_CACHE, encoding="utf-8"))
    except Exception:
        return ""
    for k in keys:
        if marker not in k:
            continue
        try:
            best = max(best, int(k.rsplit("|", 1)[1]))
        except (IndexError, ValueError):
            pass
    return datetime.fromtimestamp(best).strftime("%Y-%m-%d") if best else ""

def _index_age_days():
    """원본 색인 캐시가 **며칠째 안 돌았나**.  못 읽으면 None(모름).

    ★ 이것이 없으면 `_index_newest_day` 의 침묵이 **수집 밀림으로 읽힌다**
      (분담판 `[116]`).  색인은 09:35 원본정리 회차가 갱신하는데 그 회차가 죽으면
      캐시가 그 자리에 멈춘다 — 그러면 파일이 멀쩡히 들어와 있어도 화면은
      *"카카오톡 수집이 6일 밀렸다"* 고 말한다.  **틀린 지목은 못 잡는 것보다
      나쁘다**(`[172]`) — 사람은 이미 올린 파일을 또 내보내러 간다.
    """
    try:
        return (datetime.now() - datetime.fromtimestamp(
            os.path.getmtime(INDEX_CACHE))).total_seconds() / 86400.0
    except OSError:
        return None


def band_latest_days():
    """밴드**마다** 담고 있는 가장 최근 글의 작성일. 파일 시각이 아니라 글 날짜다.
    (파일을 다시 저장만 해도 mtime 은 오늘이 된다 — 그러면 밀린 걸 못 잡는다)

    ★ 밴드별로 따로 본다. 합쳐서 최댓값을 쓰면 **뒤처진 밴드가 가려진다** —
      실제로 매출처업무 밴드는 8/5 인데 쿠팡AS 밴드는 8/4 에 멈춰 있었고,
      돌발AS·정기점검이 오는 곳은 뒤처진 쪽이었다(2026-08-06).
    """
    out = {}
    for p in glob.glob(os.path.join(BASE, "band", "cache", "*.json")):
        name = os.path.basename(p)[:-5]
        if not name.isdigit():
            continue
        try:
            doc = json.load(open(p, encoding="utf-8")) or {}
        except Exception:
            continue
        best = 0
        for v in (doc.get("posts") or {}).values():
            try:
                best = max(best, int((v or {}).get("created_at") or 0))
            except (TypeError, ValueError):
                pass
        if best:
            out[doc.get("band_name") or name] = \
                datetime.fromtimestamp(best / 1000).strftime("%Y-%m-%d")
    return out


def band_dateless():
    """밴드마다 **본문은 있는데 작성일이 없는 글**이 몇 건인가 (2026-08-07).

    `band_latest_days()` 는 가장 최근 '날짜 있는' 글만 본다. 그래서 날짜 없는
    글은 신선도 판정에 전혀 안 잡힌다 — 밴드를 다 긁어 왔는데도 그중 621건이
    조용히 대조 밖에 있었다. 구멍도 아니고(번호가 있다) 오래된 것도 아니라
    (오늘 받았다) **어느 목록에도 안 뜨는** 종류다.

    ★ 처음 적었던 원인은 **틀렸다** (2026-08-07 2차, 실측으로 뒤집혔다).
      예전 설명: "밴드가 본문을 먼저 칠하고 작성시각을 뒤에 채우는데 그 사이에
      가져가면 날짜가 빈다 — 다시 열기만 하면 된다."
      그 설명이 맞다면 본문은 글마다 **제각각**이어야 한다. 실제로 세어 보니
      98건이 본문 **2종**, 523건이 **7종**이었다. 즉 본문까지 남의 것이었다.
      진짜 원인은 밴드가 `/post/<번호>` 를 iframe 으로 열면 **피드로 되돌리는** 것이고,
      그래서 껍데기에 남은 피드 맨 위 글이 통째로 잡혔다. 재수집으로는 못 고친다 —
      같은 경로로 다시 열면 같은 가짜가 또 들어온다(실측: 60건 재수집 → ok 0).
      그 621건은 `band/clean_contaminated.py` 가 `contaminated` 로 표시했고,
      여기서는 **따로 센다.** 수집 경로를 새로 만들기 전까지는 재수집 목록에 넣지 않는다.
    """
    out = {}
    for p in glob.glob(os.path.join(BASE, "band", "cache", "*.json")):
        name = os.path.basename(p)[:-5]
        if not name.isdigit():
            continue
        try:
            doc = json.load(open(p, encoding="utf-8")) or {}
        except Exception:
            continue
        n = sum(1 for v in (doc.get("posts") or {}).values()
                if isinstance(v, dict) and not v.get("deleted")
                and not v.get("contaminated") and not v.get("created_at"))
        if n:
            out[doc.get("band_name") or name] = n
    return out


def band_contaminated():
    """가짜로 판정돼 표시된 글이 몇 건인가 (2026-08-07).

    모은 것도 아니고 다시 훑지도 않는 상태다 — **보이게** 두어야 잊히지 않는다.
    푸는 방법은 재수집이 아니라 **수집 경로 재설계**다(iframe 상세페이지가 막혔다).
    """
    out = {}
    for p in glob.glob(os.path.join(BASE, "band", "cache", "*.json")):
        name = os.path.basename(p)[:-5]
        if not name.isdigit():
            continue
        try:
            doc = json.load(open(p, encoding="utf-8")) or {}
        except Exception:
            continue
        n = sum(1 for v in (doc.get("posts") or {}).values()
                if isinstance(v, dict) and v.get("contaminated"))
        if n:
            out[doc.get("band_name") or name] = n
    return out


def band_quiet():
    """밴드마다 '받을 것이 남았는가'의 근거 (2026-08-07 지시).

    `convert_dump._record_probe` 가 남긴 `reports/밴드_확인시각.json` 을 읽는다.
    한 밴드에 대해 "수집 최대 번호 바로 다음이 없음으로 확인됨 + 그 시각" 이 들어 있다.
    """
    try:
        return json.load(open(os.path.join(BASE, "reports", "밴드_확인시각.json"),
                              encoding="utf-8")) or {}
    except Exception:
        return {}


def _absent_judge(band, rec, today):
    """이 근거를 지금도 믿어도 되나 — 판정은 `recheck_plan.absent_line` **한 곳**이다.

    ★ 2026-08-11: 여기에도 `[217]` 과 **같은 구멍**이 있었다. 이 함수가 생기기 전에는
      근거의 **나이만** 봤다 — 신선하기만 하면 그대로 '(조용함)'을 적었다. 그런데
      근거는 **추월될 수 있다**: 실측 90610953 은 근거가 `없음확인 5438` 인데 캐시에
      **5447 이 진짜 글로** 들어와 있었다. 근거가 신선한 채로 추월되면 인계 문서가
      **없는 조용함을 확언**한다 — 낡은 근거보다 나쁜 것이 틀린 근거다.
      수집 계획(`recheck_plan`)은 이미 그것을 거르는데 인계 문서만 안 걸렀다.
      같은 파일을 보면서 판정이 갈리면 사람이 무엇을 믿을지 모르게 된다.

    근거(`rec`)는 **여기서 읽어 넘긴다**(`band_quiet()`). 판정만 빌리는 것이지 파일을
    두 번 읽는 것이 아니다 — 두 곳이 각자 읽으면 언젠가 서로 다른 한 장을 보게 된다.

    돌려주는 값: `(cut, 이유)`. `cut` 이 있으면 그 번호부터 위는 없다고 확인된 것이다.
    **못 물어보면 `(None, "")` 로 조용히 밀림 쪽에 둔다** — 밀림은 한 번 더 보라는
    말이라 잃는 것이 없지만, 잘못된 조용함은 아무도 다시 안 본다.
    """
    try:
        band_dir = os.path.join(BASE, "band")
        if band_dir not in sys.path:
            sys.path.insert(0, band_dir)
        import recheck_plan as RP
        return RP.judge_absent(rec, RP.load(band) or {}, today)
    except Exception:
        return None, ""

def band_numbers():
    """밴드 이름 -> 밴드번호. 캐시 파일의 `band_name` 에서 만든다.

    ★ 2026-08-20 실사고 — 이 함수가 없어서 조용한 밴드가 **매일** '밀렸다'로 떴다.
      `band_latest_days()` 는 키가 **이름**인데(`(주)유니버셜리프트 매출처업무`)
      근거 파일 `reports/밴드_확인시각.json` 과 `recheck_plan.load()` 는 **번호**를
      쓴다(`84789192`). 그래서 `quiet.get(이름)` 이 늘 빈 사전이었고 판정은
      `근거 없음` 으로 떨어졌다 — **한 건도 안 걸리면서 오류도 안 난다**(`[165]`).
      실측: 매출처 밴드는 8/14 이후 새 글이 없는 **조용한** 상태인데 6일 밀림으로
      떴다. 그 경보를 믿으면 사람이 **없는 번호를 긁으러 간다** — `[217]` 이 막으려던
      바로 그 사고이고, 그 고침의 인계 쪽 절반이 이 키 때문에 한 번도 안 돌았다.

    ★ 못 읽으면 **빈 사전**이다. 부르는 쪽은 그때 예전처럼 이름을 그대로 넘기므로
      판정이 '밀림'에 머문다 — 잘못된 조용함보다 낫다(`[169]`).
    """
    out = {}
    for p in glob.glob(os.path.join(BASE, "band", "cache", "*.json")):
        no = os.path.basename(p)[:-5]
        if not no.isdigit():
            continue
        try:
            doc = json.load(open(p, encoding="utf-8")) or {}
        except Exception:
            continue
        nm = str(doc.get("band_name") or "").strip()
        if nm:
            out[nm] = no
    return out


def data_freshness(today=None):
    """수집이 **얼마나 밀렸나**. 오늘(2026-08-06) 사고의 진짜 원인이 여기였다.

    밴드 최신 글이 8/4 에서 멈춰 있었는데 아무도 몰랐고, 그래서 8/5 돌발AS·정기점검이
    원장에 들어오지 않아 대표 보고가 1건·0건으로 나갔다. 사람이 "밴드 긁었더라?" 를
    기억하게 두면 다음에도 똑같이 샌다 — 기계가 날짜를 센다.

    ★ 밴드·이카운트는 **사람 로그인**이 있어야 긁을 수 있다(절대규칙). 그래서 여기서
      할 수 있는 최선은 "밀렸다"를 알리고 로그인부터 하라고 말해 주는 것이다.
    """
    day = str(today or datetime.now().strftime("%Y-%m-%d"))[:10]
    band_how = ("크롬 'Claude' 탭 그룹에서 밴드 로그인 → band/grab_posts.js 주입 →"
                " __grabStart(밴드번호, 번호목록) (한 배치 250건) → __grabSave()")
    rows = [("밴드: %s" % name, latest, band_how)
            for name, latest in sorted(band_latest_days().items())]
    rows += [
        ("카카오톡", _index_newest_day("3. 카카오톡 내보내기"),
         "카톡방 내보내기 → 다운로드 폴더에 두면 download_intake 가 흡수한다"),
        ("ERP 내보내기", _index_newest_day("1. ERP 내보내기"),
         "크롬 'Claude' 탭 그룹에서 이카운트 로그인 → 화면별 Excel 내보내기"),
    ]
    quiet = band_quiet()
    numbers = band_numbers()   # 이름 -> 번호 (근거 파일 키가 번호다)
    # ★ **밴드 자동 수집을 멈췄으면 밴드만 밀림에서 내린다** (2026-09-01 지시 · [326]).
    #   형님 지시: "밴드 자동 수집은 앞으로 하지마 이제 밴드에 자료 안올라올거야".
    #   그러면 밴드에 새 글이 없는 것이 **정상**인데 이 판정은 그것을 모른다 —
    #   실측(2026-09-01)으로 **9/5 부터 두 밴드가 매일 거짓 밀림**이고, 거짓 경보가
    #   쌓이면 진짜 경보가 묻힌다([170]).
    #   ★ 판정은 `band.collect_switch` 한 곳에서 **빌린다**([162]) — 여기서 다시
    #     재면 자동 경로 넷이 서로 다른 답을 한다.
    #   ★ **밴드만이다**([172] — 좁히는 것도 고장이다). 카톡·ERP 는 계속 들어오므로
    #     그쪽 밀림은 예전 그대로 경보한다.
    #   ★ **못 읽으면 예전 그대로 밀림이다**([169]). 여기서 기우는 방향은
    #     `collect_switch` 와 **반대다** — 저쪽이 잘못 막으면 헛 수집 한 번이지만,
    #     여기서 잘못 조용해지면 **못 받은 것을 아무도 못 본다**.
    try:
        from band import collect_switch as _CS
        band_off, band_off_why = _CS.stopped()
    except Exception:
        band_off, band_off_why = False, ""
    out = []
    for name, latest, how in rows:
        limit = FRESH_LIMIT.get(name.split(":")[0].strip(), 3)
        late = None
        if latest:
            try:
                late = (datetime.strptime(day, "%Y-%m-%d")
                        - datetime.strptime(latest, "%Y-%m-%d")).days
            except ValueError:
                late = None
        row = {"이름": name, "최신": latest or "없음", "밀린일": late,
               "한도": limit, "되살리는법": how,
               "밀림": late is not None and late > limit}
        # ★ **색인의 침묵을 수집 밀림이라 부르지 않는다**(분담판 [116] · [169]).
        #   카톡·ERP 최신일은 원본 색인에서 온다. 그 색인은 09:35 회차가 갱신하므로
        #   회차가 죽으면 캐시가 그 자리에 멈추고, 파일이 멀쩡히 들어와 있어도
        #   '밀렸다'가 뜬다. 밀린 날수가 **색인이 안 돈 날수에 먹혔으면** 그것은
        #   밀림이 아니라 **확인 못 함**이다 — 조치가 다르다(내보내기 vs 색인 돌리기).
        if row["밀림"] and not name.startswith("밴드:"):
            age = _index_age_days()
            row["색인나이"] = age
            if age is None:
                row["색인탓"] = None          # 못 읽었다 — 모른다고 적는다
            elif late is not None and late <= age + 1:
                row["색인탓"] = True
            else:
                # ★ **읽었는데 색인 탓이 아니다.** 칸을 안 넣으면 읽는 쪽 `.get()` 이
                #   '못 읽음'과 구별을 못 한다([247] — 키가 아예 없는 것과 값이 None 인
                #   것은 다른 말이다). 실측 2026-08-31: 카톡이 3일 밀리고 색인은 4시간
                #   전에 돌았는데 화면이 *"색인 나이를 못 읽어 못 갈랐다"* 고 말했다 —
                #   그러면 사람은 **색인을 돌리러 간다**. 필요한 것은 카톡 내보내기다([172]).
                row["색인탓"] = False
        # ★ '밀렸다'와 '밴드가 조용하다'는 다른 일이다 (2026-08-07 지시).
        #   날짜 있는 최신 글만 보면 새 글이 없는 날도 밀림으로 나온다. 그 경보를 믿고
        #   없는 번호를 긁으면 오늘처럼 쓰레기가 캐시로 들어간다. 그래서 '수집 최대 번호
        #   바로 다음이 없음으로 확인'된 근거가 **최근 것일 때만** 밀림을 내린다.
        #   근거가 오래됐으면 그 사이에 새 글이 올라왔을 수 있으므로 그대로 밀림이다.
        #   ★ 나이만 보면 안 된다 — 근거는 **추월될 수 있다**(2026-08-11, `[217]`).
        #   그래서 판정은 수집 계획과 같은 자리(`recheck_plan.absent_line`)에 맡긴다.
        if row["밀림"] and name.startswith("밴드:") and band_off:
            # ★ **조용히 빼지 않는다**([169]) — 왜 안 받는지 칸에 적어 두고 신선도
            #   표가 `(수집 중단)` 으로 보여 준다. 그냥 빼면 나중에 밴드를 다시 켰을 때
            #   '왜 이 밴드만 안 들어오지' 를 물을 근거가 아무 데도 없다.
            #   ★ 여기서 `_absent_judge` 를 안 부른다 — 그것은 밴드 캐시를 훑는
            #     비싼 판정인데([168]) 중단 중에는 답이 무의미하다.
            row["밀림"] = False
            row["수집중단"] = band_off_why or "밴드 자동 수집 중단"
        elif row["밀림"] and name.startswith("밴드:"):
            band = name.split(":", 1)[1].strip()
            no = numbers.get(band) or band   # 못 찾으면 예전대로(밀림에 머문다)
            q = quiet.get(no) or quiet.get(band) or {}
            cut, why = _absent_judge(no, q, day)
            if cut:
                row["밀림"] = False
                row["조용함"] = "%s번까지 수집 완료 · %s 에 새 글 없음 확인" % (
                    q.get("수집최대"), q.get("확인시각"))
            elif why:
                # 왜 아직 밀림인가를 그대로 적는다. '밀림'만 있고 이유가 없으면
                # 사람이 또 없는 번호를 긁으러 간다(그것이 [217] 의 시작이었다).
                row["근거"] = why
        out.append(row)
    return out


def unpushed_commits():
    """아직 원격에 없는 커밋. 기준은 **내 브랜치의 upstream** 이다.

    ★ 2026-08-06: 예전엔 `origin/master..HEAD` 로 셌다. master 위에서 일할 때는
      맞지만, 워크트리처럼 **기능 브랜치**에서 일하면 이미 푸시를 끝낸 커밋도
      계속 '미푸시' 로 잡힌다 — master 에 아직 안 들어갔을 뿐인데. 실측:
      브랜치를 푸시하고도 '먼저 처리할 것' 에 "푸시되지 않은 커밋 1개" 가 남았고,
      제시된 명령(`git pull --rebase && git push`)은 그 상황에 맞지도 않았다.
      새 계정이 그 말을 믿고 따라 하면 엉뚱한 브랜치를 만진다.
      upstream 이 없을 때만 예전 기준으로 돌아간다(그때는 정말 안 올라간 것이다).
    """
    base = git("rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}") or "origin/master"
    return [l for l in git("log", "%s..HEAD" % base, "--oneline").splitlines() if l.strip()]


def unmerged_commits():
    """master 에 아직 안 들어간 커밋. **막는 것은 아니다** — 브랜치 작업의 정상 상태다.
    다만 다음 사람이 "이 작업이 어디까지 갔나" 를 알아야 하므로 문서에 남긴다."""
    branch = git("rev-parse", "--abbrev-ref", "HEAD")
    if branch in ("master", "HEAD", ""):
        return []
    return [l for l in git("log", "origin/master..HEAD", "--oneline").splitlines() if l.strip()]


def collect():
    unstaged = [l for l in git("status", "--short").splitlines() if l.strip()]
    unpushed = unpushed_commits()
    return {
        "시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "브랜치": git("rev-parse", "--abbrev-ref", "HEAD"),
        "미머지": unmerged_commits(),
        "원장": ledger(),
        "큐잔량": queue_left(),
        "임시파일": temp_files(),
        "옛버전편집": stranded_editor(),
        "최신본열람": latest_viewer(),
        "손입력감지": hand_edit_signal(),
        "점유": claims(),
        "미커밋": unstaged,
        "미푸시": unpushed,
        "최근커밋": [l for l in git("log", "-5", "--oneline").splitlines() if l.strip()],
        "다음할일": next_tasks(),
        "진행체크포인트": read_checkpoint(),
        "지시문사본": rule_copies(),
        "수집신선도": data_freshness(),
        "밴드날짜없음": band_dateless(),
        "밴드오염": band_contaminated(),
        "워크트리": _worktree_state(),
        "일일대조": daily_run_health(),
        # 예산 초과로 건너뛴 단계([342]) — blockers 는 이 값만 읽는다(위 주석).
        "건너뜀": (daily_step_now() or {}).get("건너뜀") or [],
        "밴드재수집": band_recollect(),
        "업무흐름": work_flow_change(),
        "사실대조": truth_gap(),
        "대표대화못읽음": ceo_unreadable(),
        # ★ `org_gap()` 은 [297] 때 만들어졌는데 **아무도 안 불렀다**(2026-08-19
        #   실측). 지시문에는 "인계 '먼저 처리할 것'에도 올라간다"고 적혀 있었지만
        #   배선이 없어 오류도 없이 한 줄도 안 떴다 — 이 프로젝트가 반복해 당한
        #   모양이다([169]). 코드가 있는 것과 그것이 도는 것은 다른 말이다.
        # ★ 자율복구가 **오래 못 푸는 일**(2026-08-23 형님 지시 "문제되는거 처리해").
        #   판정은 `autopilot.stuck()` 한 곳이고 여기서는 담기만 한다([162]).
        #   ⚠ `blockers()` 안에서 직접 부르면 **합성 스냅샷으로 부르는 검증이 통째로
        #     막힌다**(t380 실측 — [291] 이 t111 에서 이미 겪은 자리다).
        "자율복구굳음": _autopilot_stuck(),
        "카톡보류": _kakao_held(),
        "클라우드사본": _cloud_snapshot_gap(),
        # ★ 관문 여유도 여기서 담는다 — blockers 가 파일을 직접 읽으면
        #   합성 스냅샷 검증이 막힌다(t380 실측 · [291]·[404] 와 같은 자리).
        "관문시간": gate_budget(),
        # ★ 워치독 회차가 주기를 넘겼나 — 담기만 한다([291]·[404] 와 같은 자리).
        "워치독회차": watchdog_round(),
        "조직도": org_gap(),
        "파일정리": _cleanup_notice(),
        "캠프원본": camp_source_gap(),
        # [188] — 자국은 [359] 가 남기고 여기서 읽는다(다시 세지 않는다).
        "밴드등록모호": band_register_ambiguous(),
        "오류사전": _error_book_lines(),
        "시스템진단": _system_audit_lines(),
        "세션자동화": session_auto(),
        "스케줄러": schedule_health(),
        "이어받기": takeover_health(),
        "조율": coordination_health(),
        "크롬수집": userscript_health(),
        "앱서버": app_server_health(),
    }


#: 진단은 앱 서버가 **15분마다** 다시 만든다 — 두 회차를 놓치면 낡은 것이다.
_AUDIT_STALE_MIN = 30


def _system_audit_lines():
    """앱·Claude Code·Codex 공용 진단의 **캐시**만 인계에 싣는다.

    ★ **그 캐시가 몇 분 된 것인지 같이 싣는다** (2026-08-26 실사고 · `[320]` 과
      같은 모양, 자리가 다르다). 실측 19:59:16 — 인계가
      *"`[P0]` 워치독 30분 회차가 멈춤 — 마지막 로그가 **282분 전**"* 을 실었는데
      그 순간 워치독 로그는 **1.4분 전**이었고(19:57:58) **3초 뒤** 다시 만들어진
      진단에는 그 경보가 **아예 없다.** 곧 낡은 글에 새 시각을 찍은 것이다.
      그대로 두면 사람이 **이미 풀린 고장을 고치러 간다**(`[172]`).
    ★ **조용히 빼지 않는다**(`[169]`) — 낡았어도 싣되 **몇 분 된 것인지 말한다**.
      빼 버리면 진짜 P0 가 통째로 사라진다.
    ★ **못 읽으면 지어내지 않는다**(`[169]`) — 나이를 모르면 `None` 을 그대로 싣고
      부르는 쪽이 '언제 것인지 못 읽었다'고 적는다.
    """
    try:
        import system_audit
        report = system_audit.read_cached()
        age = report.get("report_age_minutes")
        rows = []
        for row in (report.get("findings") or []):
            if row.get("priority") not in ("P0", "P1"):
                continue
            row = dict(row)
            row["진단나이분"] = age
            rows.append(row)
        return rows
    except Exception:
        return []


def userscript_health():
    """크롬 전용 수집이 **정말 돌고 있나** (2026-08-13, `[247]`).

    ★ `schedule_health` 가 스케줄러 회차에 대해 하는 일을 **브라우저**에 대해 한다.
      실측 2026-08-13: 유저스크립트가 나흘 동안 한 번도 안 돌았는데(Tampermonkey
      미설치) 그 사실을 말해 주는 화면이 어디에도 없었다.
    ★ 여기서 판정을 새로 만들지 않는다(`[162]`) — `userscript_watch` 것을 빌린다.
      읽는 것은 작은 JSON 두 개뿐이라 비싸지 않다(`[168]`).
    ★ 못 읽으면 `None` 이다. **빈 목록(=정상)과 다르다** — 아래에서 갈라 쓴다(`[169]`).
    """
    try:
        from band import userscript_watch
        return userscript_watch.lines()
    except Exception:
        return None


def schedule_health():
    """회차가 **정말 돌았나** — 스케줄러의 마지막 결과 (2026-08-12, `[228]`).

    ★ 아래 '일일자동대조' 판정은 **완주 기록**을 본다 — 즉 "안 끝났다"까지만 안다.
      **왜** 안 끝났는지는 스케줄러만 안다(제한시간에 걸려 끊겼는지, 앞 회차에 막혀
      거부됐는지, 아예 등록이 안 돼 한 번도 안 돌았는지). 둘은 겹치지 않는다.
    ★ 여기서 스케줄러를 다시 묻지 않는다(`[168]`) — 인계 문서는 자주 만들어지고
      조회는 비싸다. 워치독(30분)이 써 둔 판정을 읽기만 한다.
    """
    try:
        import schedule_watch
        return schedule_watch.banner()
    except Exception:
        return None


def coordination_health():
    """겹친 일이 **정말 누군가에 의해 되었나** (2026-08-17 지시, `[293]`).

    양보는 "저쪽이 그 일을 한다"는 주장이다. 주인마저 못 끝내면 그 일은 아무도 안 한
    것인데, 양보는 실패가 아니라서 오늘은 아무 경보도 없다.
    ★ 여기서 새 판단을 만들지 않는다 — `coordinate` 가 제 표에 적어 둔 것만 읽는다.
    """
    try:
        import coordinate
        return coordinate.notices()
    except Exception:                     # noqa: BLE001
        return None


def takeover_health():
    """다른 계정이 **지금 이어받을 수 있나** (2026-08-17 지시, `[291]`).

    크레딧이 떨어진 창은 훅이 없어 스스로 인계를 못 남긴다 — pid 는 살아 있고
    대화기록만 멈추므로 어떤 계기도 "멈췄다"고 말하지 않는다. 그 침묵을 읽는다.
    ★ 여기서 새 판단을 만들지 않는다 — `takeover` 가 이미 정한 것을 가져온다.
    """
    try:
        import takeover
        return takeover.notices()
    except Exception:
        return None                # 못 읽음 ≠ 없음. 부르는 쪽이 빈 목록으로 안 센다.


def session_auto():
    """세워 둔 일의 **막힘이 풀렸나** · 보류된 푸시가 남았나 (2026-08-11 지시).

    ★ 이것이 없어서 사람이 "하던 작업 진행" 을 두 번 쳤다. 분담판에 `[34]` 가 대기로
      앉아 있고 그것을 막던 옆 세션의 `code` 점유는 이미 풀렸는데, **그 사실을 말하는
      화면이 한 곳도 없었다.** 막힌 것을 적어 두는 것만으로는 부족하다 — 풀린 것도
      말해야 한다. 판정은 `worksplit_auto.banner()` 한 곳이 한다(여기서 점유판을 다시
      읽으면 같은 판단이 두 벌이 된다).
    """
    try:
        import worksplit_auto
        return worksplit_auto.banner()
    except Exception:
        return {}


def work_flow_change():
    """돌발AS·정기점검 **단계 정의**가 바뀌었나 (2026-08-10 지시).

    앱 화면은 이 정의를 읽어 단계 선택지를 만든다 — 관리대장 드롭다운이나 앱 흐름도가
    바뀌면 화면이 **조용히 따라가 버린다.** 따라가는 것 자체는 옳지만, 따라간 사실이
    어디에도 안 뜨면 "어제와 선택지가 다른데 왜인지 모르는" 상태가 된다.
    판정은 `work_flow.banner()` 한 곳이 한다.
    """
    try:
        import work_flow
        return work_flow.banner()
    except Exception:
        return None


def restart_defers(since_epoch):
    """이 서버가 뜬 뒤로 **몇 번 재시작을 미뤘나** (2026-08-20 지시).

    사용자 지시: **"항상 류지영 업무가 우선이야 알고리즘에 반영해"**.

    ★ **미루는 것은 옳다.** 담당자가 쓰는 중에 내리면 그 화면이 10초쯤
      끊긴다([265]). 그래서 워치독은 `--force` 로 올라가지 않는다.
    ★ 그런데 그 결과 **업무 시간 내내 옛 코드가 산다**(분담판 [175] ·
      실측 하루 21회 미룸). 그러면 사람은 인계가 시킨 `restart_server.py` 를
      눌러도 또 미뤄지고, **왜 안 되는지 어느 화면에도 안 뜬다**([169]).
      조치가 고장을 안 고치면 그것은 없는 조치다([172]).
    ★ 그러므로 **자동으로 뺏지 않고 숫자로 말한다** — 언제 내릴지는 사람이
      정한다. 그것이 이 프로젝트에서 "류지영 우선"이 뜻하는 바다.
    ★ 못 읽으면 0 이 아니라 **None** 이다([169]) — 안 미룬 것과 못 센 것은
      다른 사실이고, 뭉치면 "한 번도 안 미뤘다"는 거짓이 나간다.
    """
    try:
        sys.path.insert(0, os.path.join(BASE, "webapp"))
        import restart_server
        hist = json.load(open(restart_server.DEFER_LOG, encoding="utf-8"))
    except Exception:
        return None
    if not isinstance(hist, list) or since_epoch is None:
        return None
    n = 0
    for rec in hist:
        try:
            when = datetime.fromisoformat(str(rec.get("때") or "")).timestamp()
        except Exception:
            continue
        if when >= float(since_epoch):
            n += 1
    return n


def app_server_health():
    """앱 서버가 **옛 코드로 돌고 있나** (2026-08-08 실사고).

    그날 반나절을 이것에 썼다. 어제 저녁에 뜬 서버가 하루치 변경을 하나도 반영하지
    못한 채 돌고 있었는데, 서버는 200 을 주고 화면도 숫자를 보여 줬다 —
    **고친 사람만 모르고 있었다.** 코드를 고치고 "왜 안 바뀌지"를 반복하는 것이
    이 사고의 모양이다. 그래서 세션 인계가 매번 이것을 본다.
    """
    try:
        sys.path.insert(0, os.path.join(BASE, "webapp"))
        import restart_server
        s = restart_server.stale()
        if not s:
            return {"떠있음": bool(restart_server.running()), "옛코드": False}
        info = {"떠있음": True, "옛코드": True, "pid": s[0], "뜬시각": s[1],
                "더새로운파일": s[2]}
        # 몇 번 미뤘나 — 담당자가 쓰는 중이라 못 내린 것인지 사람이 알아야 한다.
        info["보류"] = restart_defers(restart_server._started_epoch(s[1]))
        return info
    except Exception as e:
        # ★ **못 본 것을 이상 없음이라 하지 않는다**([169]). 예전에는 그냥
        #   `{}` 였고, 그래서 없는 이름(`ROOT`) 하나가 이 검사를 **12일 동안**
        #   통째로 껐는데 어느 화면에도 안 떴다(2026-08-08 14:43 커밋 ~
        #   2026-08-20 실측). 오류도 안 나고 목록도 그럴듯했다 — [156] 이
        #   "이제 기계가 먼저 본다"고 적어 둔 그 기계가 눈이 멀어 있었다.
        return {"확인못함": str(e)[:150]}


def band_recollect():
    """08:00 재수집 회차가 **최근 30일에서 무엇이 달라졌나**를 찾았는가.

    사용자 지시(2026-08-08): "바뀐 게 있으면 인계 문서 맨 위에 올린다."
    판정은 `band/recollect.banner()` 한 곳이 한다 — 여기서 따로 세면 두 곳이 어긋난다.
    """
    try:
        sys.path.insert(0, os.path.join(BASE, "band"))
        import recollect as RC
        return RC.banner()
    except Exception:
        return None


DAILY_STALE_H = 20          # 하루 한 번 도는 것이니 20시간이면 한 회차를 통째로 건넜다
DAILY_SLOW_H = 3            # 정상 회차는 길어도 ~25분이다. 3시간이면 무언가에 막힌 것이다


def _daily_run_inflight():
    """지금 돌고 있는 daily_run 회차의 나이(시간). 없으면 None.

    ★ 잠금 파일이 있다고 돌고 있는 것이 아니다 — 죽은 회차의 잠금이 남아 있을 수 있다.
      판정은 `pid_alive` 한 곳에서 한다(검증 [121]). 모르면 '살아 있다'로 본다.
    """
    try:
        from pid_alive import owner_alive
    except Exception:
        return None
    try:
        d = json.load(open(os.path.join(REPORT_DIR, ".daily_run.lock"), encoding="utf-8"))
        started = datetime.fromisoformat(str(d.get("started_at")))
        # ★ 잠금 시각보다 뒤에 태어난 프로세스는 주인이 아니다 — pid 재사용이다.
        #   실사고(2026-08-11): 죽은 회차의 pid 를 quick_share_server 가 물려받아
        #   "5시간째 돌고 있다 — 기다려라"(정반대 지시)가 떴다. 검증 [210].
        if owner_alive(d.get("pid"), pid_started_at=d.get("pid_started_at"),
                       born_before=started.timestamp()) is not True:
            return None
        # 생성시각이 우연히 잠금보다 앞서는 재사용(상주 서비스)은 이름으로 가른다([211]).
        if _owner_is_python(d.get("pid")) is False:
            return None
        return round((datetime.now(started.tzinfo) - started).total_seconds() / 3600.0, 1)
    except Exception:
        return None


def daily_run_health():
    """일일자동대조가 **완주**한 지 얼마나 됐나 (2026-08-07 실사고).

    스케줄러는 09:50 작업을 매일 '성공(0)'으로 보고했다. 그런데 daily_run 은 앞 회차가
    아직 돌고 있으면 한 줄 찍고 **정상 종료**한다 — 잠금을 못 잡은 것을 실패로 보지 않는다.
    그래서 8/6 21:01 이후 20시간 동안 한 번도 완주하지 않았는데 어디에도 빨간불이 없었다.
    (앞 회차가 3시간씩 걸리니 다음 회차는 늘 잠겨 있다. 서로를 가려 준다.)
    완주 표식은 `finish()` 가 쓰는 agent_status.json 하나뿐이므로 그 나이를 본다.
    """
    # ★ '완주한 지 오래됨' 과 '지금 돌고 있음' 은 **다른 사실**이다 (2026-08-08 실측).
    #   그날 09:50 회차가 **12시간째** 돌고 있었는데, 이 함수는 완주 시각만 보니
    #   "20시간째 완주하지 않았다"고만 말할 수 있었다 — 안 돈 것과 구별이 안 된다.
    #   조치가 정반대다: 안 돌았으면 **띄워야** 하고, 돌고 있으면 **기다려야** 한다.
    #   그래서 잠금 파일에서 앞 회차의 나이를 같이 읽는다(살아 있는 pid 일 때만).
    running = _daily_run_inflight()
    p = os.path.join(REPORT_DIR, "agent_status.json")
    try:
        age_h = (datetime.now().timestamp() - os.path.getmtime(p)) / 3600.0
    except OSError:
        return {"완주없음": True, "경과시간": None, "중단": False, "실패단계": [],
                "진행중": running, "밀림": True}
    # ★ 나이만 보면 놓친다 — 마지막 회차가 **중단(aborted)** 으로 끝났을 수 있다.
    #   실측 2026-08-06 21:01 회차가 그랬다: aborted=True 인데 파일은 최신이라 조용했다.
    aborted, failed = False, []
    try:
        d = json.load(open(p, encoding="utf-8"))
        aborted = bool(d.get("aborted"))
        failed = [s.get("name") for s in (d.get("steps") or []) if not s.get("ok")]
    except (OSError, ValueError, TypeError, AttributeError):
        pass
    stale = age_h >= DAILY_STALE_H
    # ★ **늦은 것과 아직 예정이 안 온 것은 다른 사실이다** (2026-08-21 실사고).
    #   이 회차는 하루 한 번 09:50 이라, 어제 12:20 에 완주했으면 오늘 08:20~09:50 이
    #   그대로 20시간을 넘는다 — 회차가 늦은 것이 아니라 아직 순서가 안 온 것이다.
    #   그런데 조치가 `python daily_run.py` 라, 그대로 하면 **63분 뒤 예정된 회차를
    #   잠금으로 막는** 150분짜리 Z: 회차를 지금 띄운다(그 회차는 조용히 건너뛰고
    #   스케줄러는 '성공'이라 적는다) — 못 잡는 것보다 나쁜 조치다(`[172]`).
    #   실측 81개 간격 중 20시간 초과는 8개인데 **그 절반이 이 아침 구간**이었다.
    # ★ 판정을 새로 만들지 않는다(`[162]`·`[168]`) — 회차가 써 둔 스케줄러 사실을
    #   `schedule_watch.due_state` 에서 빌린다. 못 갈랐으면(None) **예전 그대로 밀림**
    #   이다 — '확인 못 함'을 '괜찮음'으로 치면 2026-08-07 사고가 되살아난다(`[169]`).
    # ★ **중단(aborted)·완주없음은 한 글자도 안 건드린다.** 그것은 시각과 무관한 사실이다.
    # ⚠ 여기 `not aborted` 는 **비용 절약이지 안전핀이 아니다.** 안전핀은 아래
    #   `stale or aborted` 다 — 계기 자기시험으로 확인했다(`[272]`): 이 조건을 지워도
    #   중단은 그대로 밀림이라 검사가 아무 말도 안 한다. 그러니 "검사가 이 줄을
    #   지킨다"고 여기지 말 것.
    not_due = None
    if stale and not aborted:
        try:
            import schedule_watch
            not_due = schedule_watch.due_state(
                "daily_run.py", done_at=datetime.fromtimestamp(os.path.getmtime(p)))
        except Exception:
            not_due = None
        if not_due and not_due.get("아직"):
            stale = False
    return {"완주없음": False, "경과시간": round(age_h, 1), "중단": aborted,
            "실패단계": [f for f in failed if f], "진행중": running,
            "아직예정": (not_due or {}).get("왜") if (not_due or {}).get("아직") else None,
            "밀림": stale or aborted}


def _progress_owner_alive(d):
    """진행 자국을 쓴 회차 프로세스가 **아직 그 프로세스인가**.

    ★ pid 생존만 보면 안 된다 (2026-08-11 실사고 두 번째 · 검증 [211]).
      회차 pid 37128 이 11:02~11:15 사이에 죽고 그 번호를 quick_share_server.exe
      (11:15:09 시작)가 물려받아, 인계 문서가 다섯 시간 동안 "돌고 있다 — 기다려라"
      (정반대 지시)를 냈다. 잠금([210] acquire_run_lock)과 같은 기준을 자국에도 건다:
      ① 자국의 '회차시작'·'시각'보다 **뒤에 태어난** 프로세스는 주인이 아니다
      ② 이름이 읽히는데 python 이 아니면 번호만 같은 남이다
    True=주인 살아 있음 · False=죽음(재사용 포함) · None=판정 불가(모르면 함부로
    죽었다고 하지 않는다)."""
    born = None
    for k in ("회차시작", "시각"):
        try:
            t = datetime.fromisoformat(str(d.get(k))).timestamp()
            born = t if born is None else min(born, t)
        except (TypeError, ValueError):
            continue
    res = pid_alive(d.get("pid"), born_before=born,
                    pid_started_at=d.get("pid_started_at"))
    if res is True and _owner_is_python(d.get("pid")) is False:
        return False
    return res


def daily_step_now():
    """지금 회차가 **어느 단계**에 있나 — `.daily_run.progress.json` 을 읽는다.

    ★ 2026-08-09 지시("32시간째 미완주 왜그런거야"). 그때까지 경보는 '몇 시간째'만
      말할 수 있었다. 종합리포트는 **맨 끝에 한 번** 써지므로 완주하지 못한 회차는
      **기록을 한 줄도 안 남긴다** — 그래서 원인을 물어도 댈 말이 없었다.
      이제 daily_run 이 단계마다 자국을 남기고, 여기서 그것을 읽어 **이름을 댄다.**
    ★ 자국이 있다고 돌고 있는 것이 아니다 — 죽은 회차의 자국이 남는다(그게 요점이다).
      그래서 `살아있음` 을 같이 준다([211]): '(회차 끝)' 자국은 완주한 회차라 pid 가
      죽는 것이 정상이므로 판정하지 않는다(None).
    """
    p = os.path.join(REPORT_DIR, ".daily_run.progress.json")
    try:
        with open(p, encoding="utf-8") as fh:
            d = json.load(fh)
    except (OSError, ValueError):
        return None
    try:
        since = (datetime.now().astimezone()
                 - datetime.fromisoformat(d["시각"])).total_seconds() / 60.0
    except (KeyError, ValueError, TypeError):
        since = None
    alive_now = None if d.get("단계") == "(회차 끝)" else _progress_owner_alive(d)
    return {"단계": d.get("단계"), "상태": d.get("상태"), "머문분": (round(since, 1) if since is not None else None),
            "경과분": d.get("경과분"), "예산분": d.get("예산분"),
            "끝낸수": len(d.get("끝난단계") or []), "살아있음": alive_now,
            # ★ '몇 시간째'만 말하고 **어느 단계가 그 시간을 썼는지** 못 대면 원인을
            #   영영 못 찾는다 — 292분 회차가 매일 강제 종료되는데 범인을 짐작으로만
            #   말하던 자리다(2026-08-12, `[228]`).
            "느린단계": d.get("느린단계") or [],
            # ★ 예산 초과로 **건너뛴 단계**(분담판 [82]). 회차는 '완주'로 끝나므로
            #   이것을 안 읽으면 뒤쪽 단계가 통째로 안 돈 날이 조용히 지나간다 —
            #   [180] 은 리포트에 적게 했는데 **인계는 리포트를 안 읽었다**([169]).
            "건너뜀": d.get("건너뜀") or []}


def _reboot_note(stayed_min):
    """죽은 회차 줄에 붙일 한 마디 — **그 사이 이 PC 가 재부팅됐나**. 없으면 "".

    ★ **'회차가 고장 났다'와 '기계가 꺼졌다'는 다른 사실이다** (2026-08-28 실사고).
      회차(pid 63700)가 11:34 에 떠서 **60단계를 끝내고** 13:54:42 자국을 남긴 뒤
      사라졌다. 같은 순간 워치독 13:57 회차도 자국 없이 끊겼고 `server_guard` 는
      14:08 에 '자동 시작'이었으며 Tailscale 로그인까지 없어졌다. 근거는 하나였다 —
      **부팅 13:59:22**(그 3.5분 전 이벤트 1074 — 사람이 시작 메뉴에서 전원
      끄기를 눌렀다). 곧 코드 고장이 아니다. 그런데 이 줄은
      *"★회차 프로세스가 죽었다"* 라고만 말해, 사람이 **멀쩡한 코드를 뒤진다**([172]).

    ⚠ **재는 자리는 `system_audit._boot_time` 한 곳이다**([162]) — 여기서 다시
      물으면 같은 물음에 두 답이 생긴다. 2026-08-28 에 두 창이 같은 시각에
      각자 만들어 실제로 그렇게 됐고, 옆 창 것을 정본으로 두고 물렸다([104]).

    ★ **갈래는 안 바꾼다**([385]·[468] 이 워치독·스케줄러에서 지킨 그대로).
      죽은 것은 여전히 사실이고 다시 돌려야 한다 — 덧붙이는 것은 **사실 하나**까지다.

    ★ **`살아있음 is False` 갈래에만** 붙인다([172]) — 도는 회차·완주한 회차는
      재부팅과 상관이 없다. 좁히는 것도 고장이므로 그 둘은 한 글자도 안 건드린다.

    ⚠ **`_step_hint()` 에 인자를 늘리지 않는다** — 옛 검사가 그 서명을 **글자로
      얼려** 두었다([219]·[338]). 그래서 자국 시각은 `머문분` 으로 되살린다.

    ★ **못 재면 아무 말도 안 한다**([169]) — 모르는 것을 아는 것처럼 적으면
      그것이 곧 다음 사람의 틀린 출발점이 된다.
    """
    if stayed_min is None:
        return ""
    try:
        from datetime import datetime as _dt, timedelta as _td  # noqa: PLC0415
        sys.path.insert(0, BASE)
        from system_audit import _boot_time  # noqa: PLC0415  (늦게 — 순환 방지)
        boot, _why = _boot_time()
    except Exception:
        return ""
    if not boot:
        return ""
    # 자국 시각 = 지금 - 머문분. 부팅이 그보다 **나중**이어야 근거가 선다.
    # ⚠ 여유 60초 — 자국은 초 단위로 잘려 적히고 부팅 시각도 1초쯤 흔들린다.
    #   아슬아슬한 자리에서 '재부팅했다'를 **지어내지 않는 쪽**으로 기운다([169]).
    try:
        last = _dt.now() - _td(minutes=float(stayed_min))
        if boot <= last + _td(seconds=60):
            return ""
        stamp = boot.strftime("%H:%M")
    except Exception:
        return ""
    return (" · 그 사이 이 PC 가 **재부팅됐다**(%s) — 회차는 거기 끊긴 것이다"
            "(코드가 깨진 것이 아니다)" % stamp)


def _step_hint():
    """경보 뒤에 붙일 한 줄 — **어느 단계에 머물러 있나**. 없으면 빈 문자열."""
    s = daily_step_now()
    if not s or not s.get("단계"):
        return ""
    if s.get("단계") == "(회차 끝)":
        # ★ **끝난 회차를 「지금 단계」라고 말하지 않는다** (2026-08-20 실사고).
        #   `daily_step_now` 는 이 자국에 `살아있음=None` 을 준다 — 완주한 회차는
        #   pid 가 죽는 것이 정상이라 **판정하지 않는다**([211]). 그런데 아래 갈래가
        #   `is False` 만 걸러서 **None 이 「도는 중」 쪽으로 떨어졌다.**
        #   실측 인계 문구: `마지막 회차가 **중단**으로 끝났다 · 지금 단계:
        #   **(회차 끝)**(실패) 464분째` — **한 문장이 「끝났다」와 「464분째」를**
        #   **같이 말한다.** 그러면 사람은 끝난 회차를 기다리거나 없는 것을 찾아
        #   나선다([172]·[325]). 바로 아래 주석이 막으려던 그 모양인데 갈래 하나가
        #   새 나갔다 — **자국이 남는 것과 그 회차가 도는 것은 다른 말이다.**
        #   ★ `(회차 끝)` 은 단계 이름이 아니라 **표식**이다 — 되풀이하지 않는다.
        txt = " · 마지막 자국 (%s)" % (s.get("상태") or "")
        if s.get("머문분") is not None:
            txt += " %.0f분 전" % s["머문분"]
        txt += " — **그 회차는 이미 끝났다**(지금 도는 중이 아니다)"
        if s.get("끝낸수"):
            txt += " · 끝낸 단계 %d개" % s["끝낸수"]
        if s.get("경과분") and s.get("예산분"):
            txt += " · 회차 %.0f/%s분" % (s["경과분"], s["예산분"])
        return txt + _slow_hint(s)
    if s.get("살아있음") is False:
        # ★ 죽은 회차의 자국이다 — '몇 분째'라고 적으면 돌고 있는 것처럼 읽혀
        #   사람이 다섯 시간을 기다린다(2026-08-11 실사고 · [211]). 반대로 말해야 한다.
        txt = " · 마지막 자국: **%s**(%s)" % (s["단계"], s.get("상태") or "")
        if s.get("머문분") is not None:
            txt += " %.0f분 전" % s["머문분"]
        return (txt + " — ★회차 프로세스가 죽었다(끝나기를 기다리지 말 것)"
                + _reboot_note(s.get("머문분")))
    txt = " · 지금 단계: **%s**(%s)" % (s["단계"], s.get("상태") or "")
    if s.get("머문분") is not None:
        txt += " %.0f분째" % s["머문분"]
    if s.get("끝낸수"):
        txt += " · 끝낸 단계 %d개" % s["끝낸수"]
    if s.get("경과분") and s.get("예산분"):
        txt += " · 회차 %.0f/%s분" % (s["경과분"], s["예산분"])
    return txt + _slow_hint(s)


def _slow_hint(s):
    """이 회차에서 **가장 오래 걸린 단계** 둘. 회차 예산은 단계 *사이*에서만 보므로
    한 단계가 길면 그냥 지나친다 — 그 단계를 이름으로 대야 조일 곳이 정해진다."""
    slow = [r for r in (s.get("느린단계") or []) if int(r.get("초") or 0) >= 300]
    if not slow:
        return ""
    return " · 오래 걸린 단계: " + ", ".join(
        "%s %.0f분" % (r.get("단계"), int(r.get("초") or 0) / 60) for r in slow[:2])


def _error_book_lines():
    """사람이 앱에서 막힌 자리. 회차가 써 둔 리포트를 읽기만 한다([168])."""
    try:
        import error_book
        return error_book.handoff_lines()
    except Exception:
        return []


def truth_gap():
    """`truth_watch` 회차가 써 둔 판정을 **읽기만** 한다.

    여기서 다시 세면 인계 문서를 만들 때마다 밴드 8천 건을 다시 파싱한다([168]).
    그리고 판정이 두 곳이 되어 화면·회차·인계가 서로 다른 답을 하게 된다([162])."""
    p = os.path.join(REPORT_DIR, "화면_사실대조.json")
    try:
        d = json.load(open(p, encoding="utf-8"))
    except Exception:
        # 파일이 없다 = 회차가 아직 안 돌았다. **'이상 없음'이 아니다** — 그러나
        # 첫날부터 경보를 내면 아무도 안 보므로 조용히 빈 것을 돌려준다.
        return {}
    return {"때": d.get("때", ""),
            "경보": [q for q in (d.get("물음") or []) if q.get("등급") == "경보"],
            "못물음": d.get("못물음") or []}


def ceo_unreadable():
    """`ceo_events` 회차가 '한 건도 못 읽었다'고 적어 둔 원본 — **읽기만** 한다([168]).

    2026-08-14 실사고: 대표 통화 녹취가 카톡 폴더에 멀쩡히 들어와 있었는데 파서가
    0건을 돌려줬고 **오류가 한 줄도 안 났다.** 그날 지시가 반영된 유일한 이유는
    사람이 그 파일을 눈으로 읽었기 때문이다 — 기계 경로로는 0건이었다([169])."""
    p = os.path.join(REPORT_DIR, "대표대화_추출.json")
    try:
        d = json.load(open(p, encoding="utf-8"))
    except Exception:
        return []
    # ★ 키가 **없으면** 그 회차가 아직 안 돈 옛 리포트다 — '다 읽었다'가 아니다.
    #   그렇다고 첫날부터 경보를 내면 아무도 안 보므로 조용히 빈 것을 돌려준다([247]).
    return list(d.get("못읽은원본") or [])


def _cleanup_notice():
    """쓸데없는 파일이 넘쳤나 — `cleanup_files.notice()` 를 **빌리기만** 한다([162]).

    ★ 왜 (2026-09-02): 자동 청소는 이미 매일 돌고 있었다(파이썬캐시·회차산출물·
      찌꺼기 · 실측 하루 100~200MB). 그런데 **보관본**은 되돌리기 증거라 일부러
      사람 몫으로 뒀고([463]), 그 사실을 **아무 화면도 말하지 않아** 8.5GB 까지
      쌓였다. `notice()` 는 만들어져 있었는데 **부르는 곳이 0곳**이었다 —
      코드가 있는 것과 그것이 도는 것은 다른 말이다([328]).
    ★ 여기서 무엇을 지울지 정하지 않는다 — 말하는 것까지다. 지우는 것은 사람이 정한다.
    ★ 넘칠 때만 말한다([170]) · 못 재면 아무 말도 안 한다([169])."""
    try:
        import cleanup_files as _cf
        return _cf.notice()
    except Exception:
        return None

def org_gap():
    """`org_watch` 회차가 써 둔 판정을 **읽기만** 한다 (2026-08-13 지시, [297]).

    여기서 `build()` 를 부르면 인계 한 장마다 `app_server` 를 import 하고 조직도를
    다시 조립한다([168]). 그리고 판정이 두 곳이 되어 회차와 인계가 서로 다른 답을
    한다([162]) — 회차가 적어 둔 것을 그대로 싣는다.
    ★ 파일이 없으면 = 회차가 아직 안 돌았다. '이상 없음'이 아니지만, 첫날부터
      경보를 내면 아무도 안 보므로 조용히 빈 것을 돌려준다([247] 과 같은 규칙)."""
    try:
        import org_watch
        return list(org_watch.notices(org_watch._load()))
    except Exception:
        return []


def band_register_ambiguous():
    """`band_canonical` 회차가 써 둔 **모호 자국**을 읽기만 한다 ([188] · [359] 나머지 반쪽).

    [359] 는 모호를 실패에서 빼고 자국으로 남기게 고쳤다 — 그것은 옳다(한 건이
    5분 회차를 **하루 288번** 빨갛게 만들고 있었다, [170]).  그런데 **읽는 쪽을
    아무도 안 만들어서** 자국은 남고 어느 화면에도 안 떴다.  실패가 아닌 것을
    조용히 묻으면 그것이 [169] 다 — 사람이 정할 것이 영영 사람 앞에 안 선다.
    ★ 여기서 `sync_records()` 를 부르지 않는다 — 그것은 **앱 DB 에 쓰는 길**이라
      인계 한 장이 밴드 전체를 다시 등록하게 된다([162]·[168]).
    ★ 갈래 셋을 가른다([247]): 파일 없음(=모호 없음.  회차가 없어지면 지운다 [228]) ·
      못 읽음(깨졌다 — '확인 못 함') · 목록(사람이 정할 것)."""
    try:
        import band_canonical
        path = band_canonical.AMBIGUOUS_TRACE
    except Exception:
        return {}
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as fh:
            d = json.load(fh)
        return {"건수": int(d.get("건수") or 0),
                "모호": [str(x) for x in (d.get("모호") or [])],
                "적은때": str(d.get("적은때") or "")}
    except Exception as exc:
        return {"확인못함": "%s: %s" % (type(exc).__name__, exc)}


def camp_source_gap():
    """`watch_camp_source` 회차가 써 둔 판정을 **읽기만** 한다 (2026-08-19, [328]).

    여기서 `sched_stale()` 을 직접 부르면 인계 한 장마다 Z:(SMB) 를 훑는다
    (실측 2.78초 · [168]). 그리고 판정이 두 곳이 되어 회차와 인계가 서로 다른
    답을 한다([162]).
    ★ 파일이 없으면 = 회차가 아직 안 돌았다. '이상 없음'이 아니지만 첫날부터
      경보를 내면 아무도 안 보므로 조용히 빈 것을 돌려준다([247] 과 같은 규칙)."""
    try:
        import camp_contacts
        return camp_contacts.stale_read() or {}
    except Exception:
        return {}


# ★ 진단이 **다른 감시자에게서 빌려 온** 소식은 여기서 두 번 싣지 않는다 (2026-08-19).
#   `system_audit` 이 회차 경보를 담는 것은 옳다 — 앱 실행 화면에서는 그것만 읽는다.
#   그러나 인계 문서는 그 **원본**(`schedule_watch`)도 같이 읽으므로, 그대로 두면 회차
#   경보가 **언제나 두 줄**이 된다(실측 2026-08-19 '쿠팡업무_원본자료자동정리' 가
#   `[P0] 자동 회차 실패 경보…` 한 줄 + `회차 [중단됨] …` 한 줄). 한 사건이 두 목소리로
#   울면 목록이 길어지고, 목록이 길면 **진짜 경보가 묻힌다**(`[170]`).
#   남기는 것은 **원본 줄**이다 — 갈래(`[중단됨]`)까지 말해 조치가 갈린다(`[289]`).
_AUDIT_DUP_IDS = ("scheduled-round-alert",)


def _notice_pair(x, fallback):
    """감시자가 준 소식 하나를 `(말, 조치)` 로 편다 (2026-08-20 · `[325]` 와 같은 모양).

    ★ **컨테이너를 문자열처럼 찍지 않는다.** `org_watch.notices()` 는 `(말, 조치)`
      **튜플**을 주는데 인계가 `str(line)` 으로 통째로 찍고 있었다 — 그래서 형님이
      보는 '먼저 처리할 것'에 파이썬 튜플 repr 이 그대로 나갔다:
      `('[조직도] 조직도 코드가 바뀌었는데…', 'python webapp/restart_server.py')`.
    ★ 보기 나쁜 것보다 나쁜 것은 **조치를 잃는 것**이다(`[289]`). 원본은 이미
      `python webapp/restart_server.py` 라고 정확히 적어 뒀는데 인계는 그것을 버리고
      `--print`(리포트를 한 번 더 찍는 것)를 시켰다 — 사람이 고칠 자리에 못 간다.
    ★ **모양을 하나로 못 박지 않는다.** 같은 값이 살아 있는 호출에서는 **튜플**이고
      스냅샷 JSON 을 거치면 **리스트**다(실측 둘 다 확인). 하나만 받으면 다른 쪽이
      조용히 repr 로 샌다(`[165]` — 오류는 안 난다).
    ★ **모르는 모양도 조용히 버리지 않는다**(`[169]`) — 그대로 글자로 싣는다.
      빼 버리면 그 경보가 통째로 사라진다."""
    if isinstance(x, dict):
        return (str(x.get("말") or x.get("무엇") or x), str(x.get("조치") or fallback))
    if isinstance(x, (tuple, list)):
        head = str(x[0]) if x else ""
        fix = str(x[1]).strip() if len(x) > 1 and str(x[1]).strip() else fallback
        return (head, fix)
    return (str(x), fallback)


# ★ 관문이 오래 걸렸을 때 **왜인지**를 가른다 (2026-08-27 실사고).
#   예전 조치는 언제나 "그 검사를 나눈다" 하나였다. 그런데 실측으로 `t31_tech` 가
#   3569초로 적힌 그 검사가 실제로 하는 일은 **16.3초**였다 — 코드가 아니라 그때
#   Z: 를 무는 다른 일이었다(09:35 원본정리가 40분을 물고 있었다).
#   **틀린 지목은 못 잡는 것보다 나쁘다**([172]) — 사람이 멀쩡한 검사를 쪼개러 간다.
#   ★ 조치가 갈린다([289]): 공유폴더가 붐볐으면 **회차 겹침**을 보고, 멀쩡했으면
#     그때야 검사를 나눈다. **못 쟀으면 아무 말도 안 한다**([169]).
#   ★ 여기서 Z: 를 다시 묻지 않는다([168]) — 관문이 그때 재 둔 값을 읽기만 한다.
_GATE_SHARE_SLOW_S = 5.0        # 한가할 때 1초 미만(실측 0.98초)이라 5초면 붐빈 것이다


def gate_share_note(rows):
    """관문 자국의 '공유폴더' 를 읽어 (덧붙일 말, 조치) 를 돌려준다.

    못 가르면 ("", None) 이다 — 지어내지 않는다([169])."""
    if not isinstance(rows, list) or not rows:
        return "", None                 # 안 쟀다
    try:
        got = [r for r in rows if isinstance(r, dict)]
        if not got:
            return "", None
        worst = max(float(r.get("공유초") or 0) for r in got)
        dead = any(r.get("닿음") is False for r in got)
    except Exception:
        return "", None
    look = "python schedule_watch.py --print   # 그 시각에 무엇이 돌았는지 본다"
    if dead:
        return (" · 그때 공유폴더에 **닿지 못했다** — 그 검사가 느린 것이 아니다."
                " 네트워크 드라이브 연결부터 확인한다", look)
    if worst >= _GATE_SHARE_SLOW_S:
        return (" · 그때 공유폴더 응답이 **%.1f초**였다(한가할 때 1초 미만) —"
                " 그 검사가 느린 것이 아니라 **다른 일이 공유폴더를 물고 있었다**."
                " 검사를 나누기 전에 회차가 겹치는지 본다" % worst, look)
    # ★ 공유폴더가 멀쩡해도 **그 검사가 공유폴더를 안 쓰면 아무것도 못 가른다.**
    #   2026-08-31 실측: 관문에서 `t272` 가 **225.8초**인데 그 검사는 Z: 를 한 글자도
    #   안 만지고, 한가할 때 혼자 돌리면 **37.2초**다(6배).  그때 무엇이 기계를
    #   먹었는지는 **모른다 — 모른다고 적는다**([169]).  여기서 "정말 느리다"고
    #   확언하면 사람이 **37초짜리를 쪼개러 간다**([172] 틀린 지목).
    # ★ 조치(둘째 값)는 그대로 둔다 — 좁히는 것도 고장이다([172]).
    return (" · 그때 공유폴더 응답은 %.1f초로 멀쩡했다 — 다만 **공유폴더를 안 쓰는"
            " 검사라면 이것으로는 못 가른다**(실측 t272: 관문 225.8초 · 혼자"
            " 돌리면 37.2초). 나누기 전에 **그 검사만 따로 다시 재 본다**"
            % worst, None)


def watchdog_round():
    """워치독 회차가 **주기(30분)를 넘겼나** — 회차가 써 둔 자국을 읽기만 한다.

    ★ 왜 필요한가(2026-08-28 실측): 회차 555개 중 중앙값은 7.0분인데 **19개(3%)가
      30분을 넘겼고** 최대 **354분**이었다.  그 사이 서버·회차·인계 감시가 통째로
      멈추는데 스케줄러는 '이미 실행 중'으로 건너뛰고 **성공이라 적는다**([175]).
    ★ 판정은 `watchdog.slow_note` **한 곳**이다([162]) — 여기서 다시 세면 두 답이 난다.
    ★ **끝난 회차만** 본다.  도는 중인 것을 '오래 걸렸다'고 하면 정상 회차마다 뜨고,
      멈춤 자체는 `[385]` 가 로그 나이로 이미 말한다 — 두 목소리를 만들지 않는다([325]).
    ★ 못 읽으면 아무 말도 안 한다([169]) · 안 넘겼으면 조용하다([170]).
    """
    p = os.path.join(BASE, "reports", ".워치독_진행.json")
    try:
        d = json.load(open(p, encoding="utf-8"))
    except Exception:                                            # noqa: BLE001
        return {}
    if not isinstance(d, dict) or d.get("상태") != "종료":
        return {}                                    # 아직 도는 중이거나 모양이 다르다
    try:
        분 = float(d.get("경과분") or 0)
    except Exception:                                            # noqa: BLE001
        return {}
    try:
        import watchdog as _W
        warn = float(_W.ROUND_WARN_MIN)
    except Exception:                                            # noqa: BLE001
        warn = 30.0
    if 분 <= warn:
        return {}
    느린 = [x for x in (d.get("느린단계") or []) if isinstance(x, dict)][:3]
    return {"분": round(분, 1), "한도분": warn, "느린단계": 느린,
            "끝낸단계": d.get("끝낸단계"), "언제": d.get("회차시작") or ""}


def gate_budget():
    """관문(합성검증)이 한도에 얼마나 붙었나 — 회차가 써 둔 자국을 **읽기만** 한다.

    ⚠ 이 값을 `blockers()` 안에서 직접 읽으면 **합성 스냅샷으로 부르는 검증이
      통째로 막힌다** — 실측 2026-08-24: 관문이 36.4분(한도 25분)이 된 날
      `t380` 이 죽었다(빈 스냅샷인데 이 경보가 끼어들었다). [291] 이 `t111` 에서,
      [404] 가 자율복구에서 이미 겪은 자리다 — **인계는 스냅샷만 읽는다**([162]).
    ★ 못 읽으면 지어내지 않는다({}) · 사흘 넘게 낡은 값으로는 확언하지 않는다([169]).
    """
    try:
        p = os.path.join(BASE, "reports", "합성검증_시간.json")
        with open(p, encoding="utf-8") as fh:
            g = json.load(fh)
        age_h = (datetime.now().timestamp() - os.path.getmtime(p)) / 3600.0
        m, tot, lim = g.get("여유율"), g.get("총초"), g.get("한도초")
        if m is None or not tot or not lim or age_h > 72:
            return {}
        slow = (g.get("오래걸린것") or [{}])[0]
        return {"여유율": m, "총초": tot, "한도초": lim,
                "가장오래": (slow.get("무엇") or "?")[:60], "그초": slow.get("초"),
                # 왜 오래 걸렸나 — 없으면 안 담는다([169])
                "공유폴더": (g.get("공유폴더") or None),
                # 초록이었나([372]) - `None` 은 **안 물어본 옛 자국**이다.
                #   `False` 일 때만 '초록이 아니었다'고 말한다([247]).
                "초록": g.get("초록")}
    except Exception:
        return {}


# ★ 카톡 반영이 **보류한 것**을 인계로 올린다(2026-08-26 · 분담판 [251]).
#   실측: 형님이 주신 카톡에서 새 접수 4건을 정확히 뽑고도 02_돌발AS접수 여유가
#   2행뿐이라 전량 보류였는데, 끝 줄은 '③ 엑셀 반영 (성공)' 이고 exit 0 이었다.
#   사람이 옆에 있으면 보류 줄을 읽지만 **무인 회차·앱 단추로 부르면 '성공' 만
#   남는다**([169] 의 그 모양: 숫자도 나오고 오류도 안 난다).
#   ★ 전량 보류 자체는 옳다 — 반쯤 넣는 것이 더 나쁘다. 고칠 것은 **말 안 하는 것**이다.
#   ★ 여기서 다시 판정하지 않는다([162]) — kakao_apply 가 자국에 적어 둔 것을 읽기만 한다.
#   ★ 못 읽으면 '없다'가 아니라 '못 읽음' 이다([169]).
#   ⚠ blockers() 에서 직접 부르지 않는다 — 합성 스냅샷 검증이 막힌다([291]·[404]·[407]).
def _kakao_held():
    p = os.path.join(REPORT_DIR, '카톡_반영회차.json')
    if not os.path.exists(p):
        return {}                       # 한 번도 안 돌았다 — 아무 말도 안 한다([247])
    try:
        with open(p, encoding='utf-8') as f:
            recs = json.load(f)
    except Exception as e:
        return {'못읽음': '%s: %s' % (type(e).__name__, str(e)[:80])}
    if not isinstance(recs, list) or not recs:
        return {}
    rec = recs[0] if isinstance(recs[0], dict) else {}   # _save 가 insert(0,..) — 맨 앞이 최신
    held = []
    for s in (rec.get('단계') or []):
        if isinstance(s, dict) and s.get('보류'):
            held.extend([str(x) for x in s['보류']])
    if not held:
        return {}                       # 보류 없음 — 조용하다([170])
    return {'보류': held, '때': rec.get('시각') or '',
            '파일': [str(x) for x in (rec.get('받은파일') or [])]}


# ★ 보류 문구가 **스스로 적어 둔 명령**을 그대로 쓴다([162] · 2026-08-27).
#   실측: 인계 조치가 `--sheet 02_돌발AS접수 --add 12` **고정 문자열**이었는데
#   그날 보류된 것은 **05_신규납품설치** 였다. 붙여넣으면 엉뚱한 시트를 늘리고
#   보류는 그대로 남는다([172] 틀린 지목). 게다가 `--apply` 가 빠져 있어
#   **아무것도 안 늘어난다** — 막기만 하는 안내는 없는 안내다([408]).
#   kakao_extract 는 어느 시트를 몇 행 늘려야 하는지 **이미 알고** 괄호 안에
#   그 명령을 넣어 둔다. 읽는 쪽이 다시 지으면 그것이 사본이고 언젠가 갈린다.
_HELD_CMD = re.compile(r"\(\s*(python\s+[^)]+?)\s*(?:후\s*재실행)?\s*\)")


def held_fix(line):
    """보류 한 줄에서 붙여넣어 도는 명령을 뽑는다. 못 뽑으면 None([169])."""
    m = _HELD_CMD.search(str(line or ""))
    if not m:
        return None
    cmd = " ".join(m.group(1).split())
    parts = cmd.split()
    # 없는 파일을 가리키면 그것도 틀린 조치다([448]) — 지어내느니 안 준다.
    if len(parts) < 2 or not parts[1].endswith(".py"):
        return None
    if not os.path.exists(os.path.join(BASE, parts[1].replace("/", os.sep))):
        return None
    return cmd



def _cloud_snapshot_gap():
    """폰이 1순위로 읽는 D1 최신 사본이 막혔나 (2026-08-26 지시).

    ★ **경보가 아니라 알림**이다([170]). 폰이 읽는 순서는
      `D1 최신 → GitHub Pages → 기기 사본`([271])이라, D1 이 없으면 Pages 로
      떨어지는 것이 **설계다**. 실제로 Pages 사본은 10분마다 올라간다.
      그런데 이 사실이 **아무 화면에도 안 뜨면**([328]) 폰에서 등록 예약이
      늦게 반영되는 이유를 아무도 모른다.
    ★ 키가 아예 없으면 아무 말도 안 한다([247]) — 한 번도 안 돌았을 뿐이다.
    """
    p = os.path.join(REPORT_DIR, 'cloud_continuity.json')
    if not os.path.exists(p):
        return {}
    try:
        with open(p, encoding='utf-8') as f:
            d = json.load(f)
    except Exception as e:
        return {'못읽음': '%s: %s' % (type(e).__name__, str(e)[:80])}
    if not isinstance(d, dict) or d.get('ok'):
        return {}                       # 되고 있으면 조용하다([170])
    return {'왜': str(d.get('error') or '')[:200],
            '언제': str(d.get('checked_at') or d.get('generated_at') or '')[:19]}


def _autopilot_stuck():
    """자율복구가 오래 못 푸는 일 — 판정은 `autopilot.stuck()` 한 곳이다([162]).

    ★ 못 읽으면 **빈 것으로 넘기지 않는다**([169]·[355]) — `except: pass` 는
      오타·import 실패까지 삼켜 그 기능을 통째로 없애고, 화면은 '이상 없음'으로 보인다.
    """
    try:
        import autopilot
        d = autopilot.stuck()
        d["한도"] = autopilot.STUCK_TRIES
        return d
    except Exception as exc:
        return {"굳음": [], "자원회복": [], "예산밖": [],
                "못읽음": "%s: %s" % (type(exc).__name__, exc), "한도": 0}


def blockers(st, for_sol=False):
    """다음 세션이 **먼저 처리해야** 하는 것 — 안 하면 조용히 어긋난다."""
    out = []
    # ★ 크레딧 5시간 창이 **지금 막혀 있나**(2026-08-22 형님 지시). 소진 중에만 적는다 —
    #   여유 있는 날까지 적으면 아무도 안 읽는다([170]). 사람이 할 일은 없지만
    #   **왜 AI 인계가 안 만들어지는지**를 모르면 멀쩡한 자동화를 고치러 간다([172]).
    #   ⚠ 여기서 기록을 다시 훑지 않는다 — 회차가 적어 둔 것을 읽는다([168]).
    try:
        import credit_window
        _cw = credit_window.load()
        if _cw.get("갈래") == "소진":
            out.append(("[크레딧] 5시간 창이 찼다 — %d분 뒤 충전된다. 그때까지 AI "
                        "인계 표를 만들지 않는다(충전되면 다음 회차가 스스로 잇는다). "
                        "무인 회차·엑셀·수집은 크레딧과 무관하게 그대로 돈다"
                        % (_cw.get("남은분") or 0),
                        "python credit_window.py --print"))
    except Exception:
        pass
    # ★ 자율복구가 **오래 못 푸는 일**은 인계까지 온다 (2026-08-23 형님 지시).
    #   실측: `reports/자율자동화_상태.md` 가 다섯 건을 시도 횟수까지 적고 있었는데
    #   **그 파일을 읽는 코드가 한 곳도 없어**([328]) 폰 클라우드 사본이 9일째
    #   한 번도 안 올라간 사실이 인계 17건 어디에도 없었다([169]).
    # ★ 판정은 `autopilot.stuck()` 한 곳이다([162]) · 여기서 큐를 다시 세지 않는다.
    # ★ **한 줄로 묶는다**([170]) — 원인이 저마다 달라도 사람이 볼 자리는 한 곳이고,
    #   갈래·시도·마지막 오류를 **그대로** 실어 조치가 갈리게 한다([289]).
    # ⚠ **키가 아예 없는 것**(옛 스냅샷 — 회차가 아직 안 물었다)과 **빈 것**을
    #   가른다([247]) — 안 물어본 것을 '걸린 것 없음'으로 세면 그것이 거짓이다.
    # ★ 카톡 반영 보류([251]) — 회차는 "성공" 으로 끝나므로 여기서 말하지 않으면
    #   형님이 주신 접수가 **조용히 사라진다**. 판정은 kakao_apply 가 한다([162]).
    # 폰이 1순위로 읽는 D1 사본이 막혔나 — **알림**이다(Pages 가 대신 돌고 있다).
    _cs = st.get("클라우드사본")
    if isinstance(_cs, dict) and _cs:
        if _cs.get("못읽음"):
            out.append(("[클라우드·알림] 폰 사본 상태를 못 읽었다 — %s"
                        % _cs["못읽음"], "type reports/cloud_continuity.json"))
        elif _cs.get("왜"):
            out.append((
                "[클라우드·알림] 폰이 1순위로 읽는 D1 최신 사본이 막혀 있다"
                " — GitHub Pages 사본으로 떨어져 **폰은 그대로 열린다**(설계 · [271])."
                " 다만 폰에서 넣은 등록 예약이 늦게 반영될 수 있다. 왜: %s"
                % _cs["왜"][:150],
                "python cloud_publish.py --push   # Pages 는 이것으로 즉시 갱신된다"))
    _kh = st.get("카톡보류")
    if isinstance(_kh, dict):
        if _kh.get("못읽음"):
            out.append(("[카톡] 반영 자국을 **못 읽었다** — 보류된 접수가 있는지 확인하지 "
                        "못했다(없다는 뜻이 아니다): " + str(_kh["못읽음"]),
                        "python kakao_apply.py --find"))
        elif _kh.get("보류"):
            _hl = list(_kh["보류"])
            _more = (" · 그 밖 %d건" % (len(_hl) - 1)) if len(_hl) > 1 else ""
            out.append(("[카톡] 반영이 **%d건을 보류했다 — 그만큼은 원장에 안 들어갔다**"
                        "(회차는 '성공' 으로 끝난다): %s%s · 자국 %s"
                        % (len(_hl), _hl[0][:110], _more, _kh.get("때") or "?"),
                        held_fix(_hl[0]) or "type reports\\카톡_반영회차.json"))

    _st = st.get("자율복구굳음")
    if isinstance(_st, dict):
        if _st.get("못읽음"):
            out.append(("[자율복구] 대기열을 **못 읽었다** — 오래 굳은 일이 있는지 "
                        "확인하지 못했다(걸린 것이 없다는 뜻이 아니다): " + str(_st["못읽음"]),
                        "python autopilot.py --status"))
        elif _st.get("굳음"):
            _g = list(_st["굳음"])
            _head = " · ".join("%s(%s회·%s)" % (r.get("이름"), r.get("시도"),
                                                r.get("갈래") or "?") for r in _g[:4])
            # ★ 오류를 실을 때 **누구 것인지** 같이 적는다 — 이름이 없으면
            #   넷 중 어느 건의 이유인지 몰라 엉뚱한 것을 고치러 간다([172]).
            _hit = next((r for r in _g if r.get("왜")), None)
            _why = ("%s — %s" % (_hit.get("이름"), _hit.get("왜"))) if _hit else ""
            # ★ **"AI 인계까지 실패했다"는 갈래를 봐야 참이 된다**(2026-09-05 실사고).
            #   실측: `ERP 공식 API 자료 수집` 67회 · 갈래 `resource` 가 매일 그
            #   문구로 인계 맨 위에 떴는데, `autopilot._escalate` 는 그 갈래를
            #   **애초에 AI 인계에서 뺀다** — 곧 시도조차 안 했다. 틀린 지목은
            #   못 잡는 것보다 나쁘고([172]) 가짜가 맨 위를 차지하면 진짜 경보가
            #   묻힌다([170]).
            # ★ 목록은 `autopilot.NO_AI_KINDS` **한 곳**에서 빌린다([162]) —
            #   여기 손으로 적으면 갈래가 늘어난 날 이 문구만 옛 표를 본다([165]).
            # ★ **못 읽으면 예전 문구 그대로**다([169]) — 모름을 근거로
            #   "AI 를 안 불렀다"고 확언하지 않는다.
            try:
                from autopilot import NO_AI_KINDS as _NOAI
            except Exception:
                _NOAI = None
            _kinds = [str(r.get("갈래") or "") for r in _g]
            if _NOAI and _kinds and all(k in _NOAI for k in _kinds):
                _tail = ("**코드로는 못 푼다** — 자원·인증·차선·설정이 풀려야 하는 "
                         "갈래라 AI 인계를 **애초에 안 만든다**(시도조차 안 했다는 "
                         "뜻이지 AI 가 실패한 것이 아니다)")
            elif _NOAI and any(k in _NOAI for k in _kinds):
                _tail = ("일부는 **코드로 못 푸는 갈래**(자원·인증·차선·설정)라 AI "
                         "인계를 안 만든다 — 아래 갈래를 보고 갈라 본다")
            else:
                _tail = "AI 인계까지 실패했다는 뜻이다(회차가 고장 난 것은 아니다)"
            out.append(("[자율복구] **%d건이 %s회 넘게 재시도해도 안 풀린다** — %s: %s%s"
                        % (len(_g), _st.get("한도") or "여러", _tail, _head,
                           (" · 마지막 오류: " + _why) if _why else ""),
                        "python autopilot.py --status"))
        # ★ **지나간 자원 실패는 경보가 아니라 알림이다**([424]) — 그러나
        #   조용히 빼지 않는다([169]). 실측 2026-08-25: `고정 주소 사본 올리기` 는
        #   14:59 에 죽고 **같은 일이 15:05 에 다른 길로 성공**했다. 그것을 P0 로
        #   올리면 사람이 멀쩡한 코드를 고치러 가고([172]) 진짜 경보가 묻힌다([170]).
        # ★ '고쳐졌다'고 말하지 않는다([322]) — 다음 재시도가 답한다.
        _bk = _st.get("자원회복") if isinstance(_st, dict) else None
        if _bk:
            _bh = " · ".join("%s(%s회)" % (r.get("이름"), r.get("시도")) for r in _bk[:4])
            out.append(("[자율복구·알림] 자원이 끊겨 실패했던 %d건 — 그 자원(공유폴더 등)이 "
                        "**지금은 살아 있다**. 코드 문제가 아니라 바쁜 시각에 못 잡은 것이며 "
                        "다음 재시도가 저절로 풀 수 있다(그렇다고 '고쳐졌다'는 뜻은 아니다): %s"
                        % (len(_bk), _bh),
                        "python autopilot.py --status"))
        # ★ **여기 예산 밖**인 일은 굳은 것이 아니다([436]) — 조치가 다르다([289]).
        #   그 시도 횟수는 '실패'가 아니라 **부르자마자 끊긴 횟수**라, 한 통에 담으면
        #   사람이 멀쩡한 코드를 고치러 간다([172]). 조용히 빼지도 않는다([169]).
        _ob = _st.get("예산밖")
        if _ob:
            _oh = " · ".join("%s(%s초 선언·%s회)" % (r.get("이름"), r.get("선언") or "?",
                                                     r.get("시도")) for r in _ob[:4])
            _o1 = next((r.get("예산밖") for r in _ob if r.get("예산밖")), "")
            out.append(("[자율복구·알림] %d건은 **워치독 회차 예산보다 오래 걸리는 일**이라 "
                        "여기서는 안 돌린다 — 코드가 깨진 것이 아니고, 다시 불러도 매번 "
                        "같은 자리에서 끊겨 진도가 0이다: %s%s"
                        % (len(_ob), _oh, (" · " + _o1) if _o1 else ""),
                        "python autopilot.py --status"))
        # ★ **사람이 설정을 고쳐야 풀리는 일**은 굳은 것이 아니다([289]) —
        #   조치가 다르다. 실측 2026-08-31: ERP 공식 API 가 55회를 `code` 갈래로
        #   재시도했는데 진짜 원인은 `config/ecount_config.json` 의 요청 칸
        #   이름이었다. 그 조치는 사람을 **멀쩡한 코드로 보내고**([172]) 매일
        #   P0 로 인계 맨 위를 차지해 **진짜 경보를 덮는다**([170]).
        # ★ **재시도를 멈추지 않는다** — 형님이 설정을 고치시면 다음 회차가
        #   저절로 성공해야 한다. 여기서 하는 것은 이름을 바로 붙이는 것까지다.
        # ★ 조용히 빼지 않는다([169]) — 무엇을 고치면 되는지 그대로 싣는다.
        # 크레딧 5시간 창이 막혀 **표를 아예 못 만든 것**은 굳은 것이 아니다([289]).
        #   `_escalate` 가 `ai_paused()` 에서 물러나 티켓을 안 만들기 때문이고
        #   **그것이 옳다** - 만들어 두면 `ai_ticket` 이 채워져 충전이 돼도
        #   영영 안 도는 자리가 된다(2026-08-22 지시).
        # 그러니 조치는 **아무것도 안 하는 것**이다 - 충전되면 다음 회차가
        #   그대로 이어받는다. "AI 인계까지 실패했다"고 적으면 사람이 멀쩡한
        #   코드를 뒤진다([172]).
        # 조용히 빼지 않는다([169]) - 몇 건이 기다리는지 숫자로 적는다.
        _cw = _st.get("크레딧대기") if isinstance(_st, dict) else None
        if _cw:
            _wh = " · ".join("%s(%s회)" % (r.get("이름"), r.get("시도"))
                             for r in _cw[:4])
            out.append(("[자율복구·알림] %d건은 **크레딧 5시간 창이 차서 AI 에게 "
                        "아직 안 보낸 것**이다 - 보내고 실패한 것이 아니라 **아예 "
                        "안 보냈다**. 코드에는 고칠 것이 없고 충전되면 다음 회차가 "
                        "저절로 이어받는다: %s" % (len(_cw), _wh),
                        "python credit_window.py"))
        _cf = _st.get("설정대기") if isinstance(_st, dict) else None
        if _cf:
            _ch = " · ".join("%s(%s회)" % (r.get("이름"), r.get("시도"))
                             for r in _cf[:4])
            _c1 = next((r.get("왜") for r in _cf if r.get("왜")), "")
            out.append(("[자율복구·알림] %d건은 **사람이 설정을 고쳐야 풀린다** — "
                        "코드에는 고칠 것이 없고, 설정이 그대로면 다시 불러도 같은 "
                        "자리에서 막힌다(설정을 고치시면 다음 회차가 저절로 "
                        "성공한다): %s%s"
                        % (len(_cf), _ch, (" · " + _c1) if _c1 else ""),
                        "python autopilot.py --status"))
    # ★ 관문(합성검증)의 **여유가 좁아지면 죽기 전에 말한다** (2026-08-23 형님 지시
    #   "앱 구동에 문제되는 거 전부 찾아서 · 다시는 반복되지 않게").
    #   이 관문은 `daily_run` 의 **0단계**라 시간을 넘기면 **그날 회차가 통째로 안 돈다**
    #   (대조·접수취소·오기입·사실대조·캠프 담당자가 전부 빠진다). 실측 2026-08-23:
    #   **1,350초 / 한도 1,500초 = 여유 10%**. 그런데 그 여유를 보는 눈이 **한 곳도
    #   없었다** — 죽고 나서야 `exit 1` 다섯 글자로 알았다([169]).
    # ★ **회차가 재 둔 것을 읽는다**([168]) — 여기서 관문을 돌리면 인계 한 장에 22분이 든다.
    # ★ 넉넉하면 **아무 말도 안 한다**([170]) · 못 읽으면 지어내지 않는다([169]).
    # ★ 답은 늘리는 것이 아니라 **나누는 것**이다 — 검증은 계속 늘고, 한도만 늘리면
    #   같은 사고가 더 큰 값에서 되풀이된다(지시문이 [318] 에서 정해 둔 그대로다).
    _g = st.get("관문시간") or {}
    if isinstance(_g, dict) and _g:
        _m, _tot, _lim = _g.get("여유율"), _g.get("총초"), _g.get("한도초")
        if _m is not None and _tot and _lim:
            _slow = {"무엇": _g.get("가장오래"), "초": _g.get("그초")}
            # ★ **왜** 오래 걸렸는지까지 말한다([289] — 조치가 갈린다).
            #   못 쟀으면 예전 그대로다([169] — 모름을 확언하지 않는다).
            _why, _act = gate_share_note(_g.get("공유폴더"))
            # ★ 그 관문이 **초록이 아니었으면** 이 숫자는 전체가 아니다 -
            #   검사 도중 멈춘 값이라 '여유 N%' 를 그대로 믿으면 안 된다.
            #   `None`(안 물어본 옛 자국)에는 **아무 말도 안 한다**([247]).
            if _g.get("초록") is False:
                _why = (_why or "") + (" · ★ 그 관문은 **초록이 아니었다**"
                                       " - 검사 도중 멈춘 값이라 이 숫자는"
                                       " 전체가 아니다")
            _fix = _act or "python tests/synthetic_check.py   # 그 검사를 나눈다(한도를 먼저 늘리지 말 것)"
            if _m < 0:
                out.append((
                    "[관문] 합성검증이 한도를 **넘겼다** — %.1f분 / 한도 %.1f분. 이 관문은 "
                    "일일대조의 0단계라 여기서 끊기면 **그날 회차가 통째로 안 돈다**. "
                    "가장 오래 걸린 것: %s (%s초)%s"
                    % (_tot / 60.0, _lim / 60.0, (_slow.get("무엇") or "?")[:60],
                       _slow.get("초"), _why),
                    _fix))
            elif _m < 0.20:
                out.append((
                    "[관문] 합성검증 여유가 %d%% 뿐이다 — %.1f분 / 한도 %.1f분. 바쁜 아침에는 "
                    "넘겨 **그날 회차가 첫 줄에서 죽는다**. 가장 오래 걸린 것: %s (%s초)%s"
                    % (round(_m * 100), _tot / 60.0, _lim / 60.0,
                       (_slow.get("무엇") or "?")[:60], _slow.get("초"), _why),
                    _fix))
    # ★ 워치독 회차가 **주기를 넘기면** 그 사이 감시가 통째로 멈춘다 — 그런데
    #   스케줄러는 '이미 실행 중'으로 건너뛰고 성공이라 적는다([175]).  어느 단계가
    #   먹었는지는 회차가 남긴 자국만 안다([228] 을 워치독으로 · 2026-08-28).
    _wd = st.get("워치독회차") or {}
    if isinstance(_wd, dict) and _wd.get("분"):
        _slow = ", ".join(
            "%s %s분" % (str(x.get("단계"))[:40], x.get("분"))
            for x in (_wd.get("느린단계") or []))
        out.append((
            "[워치독] 지난 회차가 **%s분** 걸렸다(주기 %s분) — 그 사이 서버·회차·인계 "
            "감시가 멈췄고 스케줄러는 '이미 실행 중'으로 건너뛰며 성공이라 적는다. "
            "오래 걸린 단계: %s"
            % (_wd.get("분"), int(_wd.get("한도분") or 30), _slow or "1분 넘긴 단계 없음"),
            "type reports" + chr(92) + ".워치독_진행.json"
            "   # 단계별 초가 들어 있다 (제한시간을 먼저 늘리지 말 것 · 분담판 [38])"))
    # ★ **원본이 실제로 실렸을 때만** 건너뛴다(`[169]`). 스케줄러 감시를 못 읽은 날
    #   빼 버리면 그 경보가 통째로 사라진다 — '다른 데 있겠지'는 근거가 아니다.
    borrowed_ok = bool((st.get("스케줄러") or {}).get("경보"))
    for row in (st.get("시스템진단") or []):
        if borrowed_ok and row.get("id") in _AUDIT_DUP_IDS:
            continue
        # ★ 낡은 진단을 **지금 사실처럼** 적지 않는다(`[449]`) — 싣되 나이를 말한다.
        _age = row.get("진단나이분")
        if _age is None:
            _note = " · ⚠ 이 진단이 언제 것인지 못 읽었다"
        elif _age > _AUDIT_STALE_MIN:
            _note = (" · ⚠ 이 진단은 %d분 전 것이다 — 그 뒤에 풀렸을 수 있다"
                     % int(_age))
        else:
            _note = ""
        out.append(("[%s] %s — %s%s" % (row.get("priority", ""), row.get("title", ""),
                                        str(row.get("evidence") or "")[:150], _note),
                    row.get("action") or "python system_audit.py --print"))
    # ★ 화면이 조용히 틀린 값을 보여 주는 것 (2026-08-13 지시: "위 같은 문제 잡아내는
    #   기능 AI 추가해"). 그날 둘 다 **사람이 전화로 지적하고 나서야** 알았다 —
    #   캘린더 미처리의 31%가 이미 취소된 건이었고([243]), 다녀온 현장이 원장 완료일이
    #   비어 미처리로 서 있었다([244]). 오류가 안 나는 종류라 아무 화면에도 안 떴다.
    #   판정은 `truth_watch` 회차가 미리 해 둔다 — 여기서 다시 세지 않는다([168]).
    tw = st.get("사실대조") or {}
    for q in (tw.get("경보") or []):
        out.append(("화면과 원본이 어긋난다 — %s: %s" % (q.get("무엇"), q.get("값")),
                    "python truth_watch.py --print"))
    for m in (tw.get("못물음") or []):
        # '못 물어봄'을 '이상 없음'으로 치지 않는다([169]) — 감시자가 눈먼 것이다.
        out.append(("사실대조가 확인하지 못한 것이 있다 — %s" % m,
                    "python truth_watch.py --print"))
    # ★ 대표 지시가 담긴 원본을 기계가 못 읽었다 (2026-08-14 실사고). 파일은 멀쩡히
    #   있고 회차도 성공이고 오류도 안 난다 — 그래서 여기 안 적히면 **아무 화면에도
    #   안 뜬다**. 조치 칸은 '사람이 그 파일을 읽는다'이지 자동 추출이 아니다([172]):
    #   녹취에는 누가 대표인지 적혀 있지 않아 잘못 뽑으면 **없는 지시가 캡처에 뜬다**.
    for row in (st.get("대표대화못읽음") or []):
        out.append(("대표 대화 원본을 기계가 한 건도 못 읽었다(%s) — %s"
                    % (row.get("갈래", "모름"), row.get("파일", "")),
                    "python ceo_events.py --sync   # 그 뒤 파일은 사람이 직접 읽는다"))
    # ★ 조직도가 바뀌었는데 **못 따라간 것** (2026-08-13 지시, [297]). 바뀐 것 자체는
    #   경보가 아니다 — 잘 따라갔으면 정상이고, 정상까지 경보하면 묻힌다([170]).
    for line in (st.get("조직도") or []):
        out.append(_notice_pair(line, "python org_watch.py --print"))

    # ★ 쓸데없는 파일이 쌓였다 — **알림이지 경보가 아니다**([170]).
    #   자동 갈래(파이썬캐시·회차산출물·찌꺼기)는 워치독이 매일 지운다.
    #   여기 걸리는 것은 대개 **보관본**이고, 몇 개를 남길지는 업무 판단이라
    #   기계가 안 정한다([463]) — 그래서 말하는 것까지가 몫이다.
    #   ★ 키가 아예 없으면(옛 스냅샷) 아무 말도 안 한다([247]).
    _cl = st.get("파일정리")
    if _cl:
        out.append(_notice_pair(_cl, "python cleanup_files.py"))
    # ★ 밴드 앱 DB 등록에서 **사람이 정해야 하는 것**([188] · [359] 의 나머지 반쪽).
    #   같은 프로젝트에 앱 DB 행이 여럿이면 어느 행을 고칠지 원본이 말해 주지 않는다 —
    #   실패가 아니므로 회차는 초록이고, 그래서 여기 안 적히면 **아무 화면에도 안 뜬다**.
    #   ★ 모호가 없으면 조용하다([170]) — 정상까지 경보하면 진짜 경보가 묻힌다.
    ba = st.get("밴드등록모호") or {}
    if ba.get("확인못함"):
        out.append(("밴드 앱 DB 등록의 모호 자국을 못 읽었다(%s) — '모호가 없다'는 뜻이 아니다"
                    % ba["확인못함"], "python band_canonical.py --ambiguous"))
    elif ba.get("모호"):
        rows = ba["모호"]
        out.append(("밴드 앱 DB 등록에 **사람이 정할 것 %d건** — 같은 프로젝트에 앱 DB 행이"
                    " 여럿이라 어느 행인지 원본이 말해 주지 않는다(회차 실패가 아니다): %s"
                    % (len(rows), " · ".join(rows[:3])),
                    "python band_canonical.py --ambiguous"))
    # ★ 정기점검 스케줄 원본이 새로 왔는데 앱 담당자 자료가 안 따라갔다 (2026-08-19
    #   실사고 · [328]). 이 사고는 **빈칸이 아니라 '틀린 사람'** 으로 나타나므로
    #   화면만 봐서는 영영 모른다([165]) — 실측 754칸 중 236칸이 옛 담당자였다.
    #   ★ 여기서 자동으로 다시 만들지 않는다([168]) — 밴드 전체 파싱이라 비싸고,
    #     무엇을 덮을지는 사람이 결정한다.
    _cs = st.get("캠프원본") or {}
    if _cs.get("갈래") in ("밀림", "없음", "모름"):
        out.append((("캠프 담당자 자료 — %s" % (_cs.get("말") or "확인 못 함")),
                    _cs.get("조치") or "python camp_contacts.py --stale"))
    # ★ 같은 원본이 **달력의 정기점검 예정**도 먹인다 (2026-08-20 · 분담판 `[168]`).
    #   그런데 감시자는 담당자 자료만 봤다 — 그래서 달력이 **옛 예정을 조용히**
    #   보여 주는 동안 어느 화면도 그 말을 안 했다. 형님은 "자료를 안올렸나?" 를
    #   물으셨고 **자료는 올라와 있었다**(실측 2026-08-19: 원본 09:11 · 그 파일
    #   08-18 14:04 · 손으로 돌리자 확정 50 → 53건).
    #   ★ **갈래를 합치지 않는다**(`[289]`) — 고치는 명령이 다르다. 합치면 한쪽만
    #     낡은 날 사람이 엉뚱한 명령을 돌리고 원인은 그대로 남는다(`[172]`).
    #   ★ 키가 **아예 없으면** 아무 말도 안 한다(`[247]`) — 옛 회차가 쓴 스냅샷은
    #     그 칸을 안 물었을 뿐이지 '정상'도 '고장'도 아니다.
    _cal = _cs.get("달력") or {}
    if _cal.get("갈래") in ("밀림", "없음", "모름"):
        out.append((("달력 정기점검 예정 — %s" % (_cal.get("말") or "확인 못 함")),
                    _cal.get("조치") or "python pm_schedule_sync.py --apply"))
    # ★ 사람이 앱에서 막힌 자리 (2026-08-13 지시). 오류가 기록되는 것과 누가 그것을
    #   보는 것은 다른 말이다 — `/api/originals` 권한거부 222건이 최근 3일에 쌓이는
    #   동안 어느 화면에도 안 떴다. 회귀(막았다는데 또 남)가 새 오류보다 앞에 온다.
    #   판정은 `error_book` 회차가 해 둔 것을 읽기만 한다([168]).
    #   ★ 여기서 디스크를 직접 읽지 않는다 — 읽었더니 합성검증이 실기계 상태를 타서
    #     `t111` 이 즉시 빨개졌다. 판단 재료는 언제나 st 를 거친다.
    for line in (st.get("오류사전") or []):
        out.append((line.split(" → ")[0], "python error_book.py --print"))
    # ★ 워크트리가 본체 상태와 끊겨 있으면 **합성검증부터 못 돈다**(config 가 없다).
    #   즉 "ALL GREEN 확인 후 실작업" 관문 자체를 통과할 수 없으니 맨 앞에 둔다.
    #   판단은 st 에 담긴 것만 쓴다 — 여기서 기계 상태를 직접 보면 합성검증이
    #   실행 환경에 따라 결과가 달라진다(테스트가 기계를 타면 안 된다).
    wt = st.get("워크트리")
    if wt:
        cut = [i["대상"] for i in wt.get("항목", []) if i.get("상태") == "없음"]
        if cut:
            out.append(("워크트리가 본체 상태와 끊겨 있다 — %s (설정·큐를 못 읽어 "
                        "합성검증부터 막힌다)" % ", ".join(cut[:4]),
                        "python worktree_state.py --apply"))
    # ★ **풀린 것도 말한다** (2026-08-11). 막힌 것만 적어 두면, 막고 있던 것이 사라진
    #   순간 그 일은 아무도 모르게 세워진 채 남는다 — 이날 실측으로 사람이 "하던 작업
    #   진행" 을 두 번 쳤다. 근거는 `worksplit_auto` 가 회차마다 새로 쓴다.
    #   ★ 그리고 **어느 창에서 하는 일인지**까지 적는다 (2026-08-12, `[238]`). 자원 점유가
    #     풀려도 **차선**(`lanes`)이 다른 창에서는 `--take` 가 거부된다 — 실측으로 인계가
    #     [39][48][50] 을 '가져가라'로 올렸는데 수집 차선 창에서는 안 됐다. 그러면 '풀렸다'가
    #     두 뜻이 되어, 풀리지 않은 것을 풀렸다고 말한 셈이 된다(`[169]` 와 같은 모양).
    for row in ((st.get("세션자동화") or {}).get("풀린일") or [])[:3]:
        lane = (row.get("차선") or "").strip()
        out.append(("세워 둔 일 **[%s]** 의 막힘이 풀렸다 — %s%s"
                    % (row.get("id"), (row.get("title") or "")[:60],
                       (" (`%s` 차선 창에서)" % lane) if lane else ""),
                    "python worksplit.py --take %s --who claude" % row.get("id")))
    # ★ 회차가 **아예 안 돌았거나 죽었다** — 스케줄러만 아는 사실이다 (2026-08-12, `[228]`).
    #   아래 '일일대조' 판정은 완주 기록을 보므로 "안 끝났다"까지만 안다. **왜**인지는
    #   여기가 말한다: 제한시간에 걸려 끊겼는지, 앞 회차에 막혀 거부됐는지, 등록조차
    #   안 됐는지. 실측 2026-08-12 — 일일대조·원본정리가 매일 강제 종료되고 정오회차는
    #   등록이 안 돼 한 번도 안 돌았는데 **어느 화면에도 안 떴다.**
    sw = st.get("스케줄러") or {}
    if sw.get("조회실패"):
        # 못 본 것을 정상이라 하지 않는다(`[169]`) — 감시자가 눈멀었다는 것도 소식이다.
        out.append(("스케줄러 상태를 **확인 못 했다** — %s. 이것은 '이상 없음'이 아니다"
                    % str(sw["조회실패"])[:80], "python schedule_watch.py"))
    # ★ 알림은 경보와 갈라 싣는다(`[288]`) — '고친 뒤 아직 안 돌았다' · '꺼져 있다' 는
    #   지금 고칠 고장이 아니지만 **아무 데도 안 적히면 하루 넘게 안 도는 회차가 생긴다**
    #   (실측 2026-08-16: 09:50 일일대조가 08-15 부터 꺼져 있었고 어느 화면에도 없었다).
    for a in (sw.get("알림") or [])[:5]:
        out.append(("회차 [%s] %s" % (a.get("갈래", ""),
                                    re.sub(r"\*\*", "", str(a.get("무엇") or ""))[:130]),
                    a.get("어떻게") or "python schedule_watch.py --print"))
    # ★ 크레딧이 떨어진 창은 **스스로 인계를 못 남긴다**(훅이 없다). 그 침묵을
    #   여기 올려야 다른 계정이 "이어받아도 된다"는 것을 안다(2026-08-17, `[291]`).
    # ⚠ 여기서 **다시 재지 않는다** — `st` 가 준 것만 읽는다(`sw` 와 같은 규칙).
    #   살아 있는 기계를 직접 물으면 합성 상태로 부르는 자리가 늘 막힌다(t111 실측).
    for a in (st.get("이어받기") or [])[:4]:
        out.append(("이어받기 [%s] %s" % (a.get("갈래", ""),
                                      re.sub(r"\*\*", "", str(a.get("무엇") or ""))[:130]),
                    a.get("어떻게") or "python takeover.py"))
    # ★ 겹침도 같은 자리다 (2026-08-17, `[293]`). 양보는 실패가 아니라서 스케줄러도
    #   회차 감시도 아무 말을 안 한다 — 그런데 주인마저 안 끝냈으면 그 일은 아무도 안 했다.
    for a in (st.get("조율") or [])[:3]:
        out.append(("조율 [%s] %s" % (a.get("갈래", ""),
                                    re.sub(r"\*\*", "", str(a.get("무엇") or ""))[:130]),
                    a.get("어떻게") or "python coordinate.py --print"))
    for a in (sw.get("경보") or [])[:5]:
        # 바깥에서 한 번 더 굵게 감싸므로 여기서는 `**` 를 쓰지 않는다(겹치면 안 굵어진다).
        out.append(("회차 [%s] %s" % (a.get("갈래", ""),
                                    re.sub(r"\*\*", "", str(a.get("무엇") or ""))[:110]),
                    a.get("어떻게") or "python schedule_watch.py --print"))
    # ★ 크롬 전용 수집도 같은 자리다 — 스케줄러가 모르는 축이다 (2026-08-13, `[247]`).
    #   회차는 '돌았다'고 말하는데 정작 긁는 것은 브라우저 안 유저스크립트다. 그것이
    #   꺼져 있으면 **어느 회차도 실패하지 않으면서 수집만 0건**이 된다.
    #   ★ 세 가지를 가른다 — 뭉치면 없는 경보가 난다:
    #     · 키가 **아예 없다** = 부르는 쪽이 안 물었다(부분 상태) → 아무 말도 안 한다
    #     · `None`            = 물어봤는데 **실패**했다            → '확인 못 함'
    #     · 빈 목록           = 물어봤고 **정상**이다              → 아무 말도 안 한다
    us = st.get("크롬수집", ())
    if us is None:
        # ★ 못 읽은 것과 '정상(빈 목록)'은 다르다(`[169]`). 뭉치면 감시자가 눈먼 채
        #   "이상 없음"을 말한다 — 이 기능이 막으려던 바로 그 모양이다.
        out.append(("크롬 전용 수집 상태를 **확인 못 했다** — 이것은 '이상 없음'이 아니다",
                    "python band/userscript_watch.py --print"))
    elif us:
        # ★ 조치 칸은 **붙여넣어 도는 명령**이어야 한다 — 안내 문장을 거기 넣으면
        #   명령처럼 보여 사람이 그대로 붙여넣는다. 사람이 할 일은 설명 쪽에 적는다.
        head = re.sub(r"\*\*", "", str(us[0]))
        fix = str(us[1]).strip().lstrip("→ ").strip() if len(us) > 1 else ""
        if fix:
            head = "%s → %s" % (head, fix)
        out.append((head[:230], "python band/userscript_watch.py --print"))
    # ★ 스케줄러가 '성공'이라 말해도 완주하지 않았을 수 있다 — 잠금을 못 잡은 회차가
    #   조용히 exit 0 으로 끝나기 때문이다. 그 사이 자료현황·대조 리포트가 통째로 멈춘다.
    dr = st.get("일일대조") or {}
    run_h = dr.get("진행중")
    if dr.get("밀림"):
        hrs = dr.get("경과시간")
        # 2026-08-28 실사고 — **한 줄이 서로 다른 두 회차를 말하고 있었다.**
        #   `중단` 은 회차가 **끝날 때** 쓰는 `agent_status.json` 에서 오고(= 앞 회차),
        #   바로 뒤에 붙는 `_step_hint()` 는 **지금 도는** 회차의 진행 자국을 읽는다.
        #   그래서 실측 인계 문구가 `마지막 회차가 **중단**으로 끝났다 · 지금 단계:
        #   **Z폴더 원장 누락·금액 공백 스캔**(시작) 22분째` 였다 — 「끝났다」와
        #   「22분째」를 같이 말한다([338]·[325] 와 같은 모양, 자리가 다르다).
        #   조치는 맞았지만(기다린다) 사람은 "끝났는데 왜 22분째지"로 읽는다.
        # ★ **도는 회차가 있을 때만** 갈라 말한다 — 없으면 예전 그대로다([172] ·
        #   좁히는 것도 고장이다). `완주없음`·`N시간째` 갈래는 한 글자도 안 건드린다:
        #   "한 번도 완주하지 않았다"와 "지금 22분째"는 서로 어긋나지 않는다.
        # ⚠ `_step_hint()` 에 인자를 늘리지 않는다 — 옛 검사가 `def _step_hint():`
        #   **서명을 글자로 얼려** 두었다([219]). 두 조각이 만나는 여기서 고친다.
        why = (("**앞** 회차가 **중단**으로 끝났다 — **지금 도는 것은 그다음 회차다**"
                if run_h is not None else "마지막 회차가 **중단**으로 끝났다")
               if dr.get("중단")
               else "%s 완주하지 않았다" % ("한 번도" if hrs is None else "%.0f시간째" % hrs))
        bad = dr.get("실패단계") or []
        # ★ 지금 돌고 있으면 **띄우라고 하면 안 된다** — 잠금에 막혀 조용히 건너뛴다.
        #   기다리라고 말해야 한다. 조치가 정반대라 이 한 줄이 갈림길이다.
        act = ("(지금 %.0f시간째 돌고 있다 — 새로 띄우지 말고 끝나기를 기다린다)" % run_h
               if run_h is not None
               else "python daily_run.py    # 먼저 tasklist 로 앞 회차가 도는지 확인")
        why += _step_hint()
        # ★ 여기서 오늘의 스케줄러 결과를 **확언하지 않는다** (2026-08-19).
        #   예전 문구는 "스케줄러는 '성공'으로 보고한다"였는데, 실측으로 그날
        #   스케줄러는 `코드 1`(실패)이라 적고 있었다 — 바로 두 줄 위 `회차 [고침대기]`
        #   가 그렇게 말한다. 한 문서 안에서 두 줄이 서로 어긋나면 사람은 **없는 것을
        #   찾아 나선다**(`[172]`). 이 괄호는 '이 검사가 왜 있나'라는 일반론이지
        #   오늘에 대한 주장이 아니므로, 주장이 아닌 말로 적는다.
        out.append(("일일자동대조 — %s. 스케줄러가 '성공'이라 적어도 완주하지 않았을 수 있다"
                    "(앞 회차가 도는 동안 다음 회차가 조용히 건너뛴다)%s"
                    % (why, (" · 실패단계: " + ", ".join(bad[:4])) if bad else ""), act))
    elif run_h is not None and run_h >= DAILY_SLOW_H:
        # 완주 기록은 아직 싱싱한데 **지금 회차가 비정상적으로 길다.** 이대로 두면
        # 20시간을 넘겨서야 위 경보가 뜬다 — 그때는 이미 하루를 잃은 뒤다.
        out.append(("일일자동대조가 **%.0f시간째** 돌고 있다 (보통 25분)%s — Z: 를 훑는 다른"
                    " 작업과 겹쳤을 수 있다. 이 회차가 끝날 때까지 다음 회차는 조용히 건너뛴다"
                    % (run_h, _step_hint()),
                    "python session_handoff.py --check   # 끝나면 저절로 사라진다"))
    # ★ 앱 서버가 옛 코드로 돌면 **고쳐도 화면이 안 바뀐다.** 서버는 200 을 주고
    #   화면은 숫자를 보여 주므로 아무도 옛 서버인 줄 모른다(2026-08-08 반나절).
    ap = st.get("앱서버") or {}
    if ap.get("확인못함"):
        out.append(("앱 서버가 **옛 코드인지 확인하지 못했다** — %s"
                    % str(ap.get("확인못함"))[:110],
                    "python webapp/restart_server.py --status"))
    if ap.get("옛코드"):
        # ★ 담당자가 쓰는 중이면 그냥 부르는 것은 **또 미뤄진다** — 조치가 고장을
        #   안 고치면 없는 조치다([172]). 몇 번 미뤘는지 세어 말하고, 그때는
        #   "조용한 때에" 라는 조건을 붙여 `--force` 를 알려 준다.
        #   **자동으로 뺏지 않는다** — 언제 내릴지는 사람이 정한다(2026-08-20 지시:
        #   "항상 류지영 업무가 우선이야").
        held = ap.get("보류")
        many = isinstance(held, int) and held >= 2
        out.append(("앱 서버가 **옛 코드**로 돌고 있다 (pid %s · 뜬 시각 %s) — "
                    "%s 가 서버보다 새것이다. 고쳐도 화면이 안 바뀐다%s"
                    % (ap.get("pid"), ap.get("뜬시각"),
                       ", ".join((ap.get("더새로운파일") or [])[:4]),
                       (" · 담당자가 쓰는 중이라 **%d번 미뤘다**(류지영 업무가"
                        " 우선이라 안 내렸다) — 아무도 안 쓰는 때에 내린다" % held)
                       if many else ""),
                    "python webapp/restart_server.py --force" if many
                    else "python webapp/restart_server.py"))
    if st["큐잔량"]:
        out.append(("입력 큐에 %d건이 반영되지 않았다" % st["큐잔량"],
                    "python ledger_db.py --intake  # Excel은 다음 11:00·15:00 회차"))
    # ★ 날짜 없는 글은 **신선도 판정에 안 잡힌다** — band_latest_days() 는 날짜 있는
    #   글만 보기 때문이다. 그래서 "밴드 최신 = 오늘" 인데도 그 밑에 수백 건이
    #   대조 밖에 있을 수 있다. 조용한 사고라 여기서 말해 준다.
    dl = st.get("밴드날짜없음") or {}
    if sum(dl.values()) >= 50:
        out.append(("밴드에 **본문은 있는데 날짜가 없는 글**이 %d건 (%s) — "
                    "날짜가 없으면 어떤 작업과도 대조되지 않는다"
                    % (sum(dl.values()),
                       ", ".join("%s %d" % (k, v) for k, v in sorted(dl.items()))),
                    "크롬 창을 **앞으로 꺼낸 뒤** python band/recheck_plan.py 로 "
                    "'날짜없음' 목록을 뽑아 재수집 (숨은 탭에서는 수집기가 시작을 거절한다)"))
    # 가짜(오염) 글 기록은 **할 일이 아니라 상태다** (2026-08-07 확정). 표본 3/3 이
    # 피드로 리다이렉트됐다 — 밴드가 그 번호를 못 보여 준다, 즉 삭제된 글이다.
    # 되살릴 내용이 없으므로 여기 '먼저 처리할 것'에 올리지 않는다(올리면 다음 세션이
    # 또 풀려고 한 시간을 쓴다). 표시는 신선도 표 밑에 한 줄로 남는다.
    if st["임시파일"]:
        out.append(("원장 임시파일이 남았다(만들다 끊김): %s" % ", ".join(st["임시파일"][:3]),
                    "내용 확인 후 삭제 — 정식 vN+1 로 승격되지 않은 파일이다"))
    if st.get("옛버전편집"):
        out.append(("**사람이 옛 버전을 열어 편집 중이다**: %s (최신본이 아니다)"
                    % ", ".join(st["옛버전편집"][:3]),
                    "그 창에 적는 것은 다음 회차로 넘어가지 않는다 — 지금 적은 것을 최신본에 "
                    "옮기고 옛 창을 닫도록 사람에게 알린다. 남의 엑셀은 대신 닫지 않는다"))
    # ★ 엑셀 손입력 감지(2026-08-11 지시 — 입력 창구는 앱 하나). 역수입 금지라 그
    #   값은 정본에 안 들어간다 — 말없이 버리면 그 사람의 입력이 소리 없이 사라진다.
    he = st.get("손입력감지")
    if he:
        last = he.get("마지막") or {}
        # ★ **조용히 빼지 않는다**([169]) — 뺀 것이 있으면 숫자로 말한다.
        빠짐 = he.get("뺀건수") or 0
        꼬리 = (" · 기계 도구가 만든 것 %d건은 뺐다(경보 아님)" % 빠짐) if 빠짐 else ""
        out.append(("엑셀 **손입력 감지** %d건(24시간) — 마지막: %s %s%s. "
                    "손으로 적은 값은 정본(DB)에 반영되지 않는다"
                    % (he.get("최근24h", 0), last.get("종류", ""),
                       last.get("파일") or last.get("잠금") or "", 꼬리),
                    "적은 사람을 찾아 **앱으로 다시 입력**하도록 안내한다 "
                    "(자동 반영 금지 — 역수입 금지)"))
    for c in st["점유"]:
        m = c["mins"] if c["mins"] is not None else "?"
        if c["stale"]:
            why = "주인 프로세스 %s 가 없다" % c.get("pid") if c.get("alive") is False \
                else "%s분 경과" % m
            # 내 것이면 --free 로 놓이지만, **남의 죽은 세션 것은 --free 가 거부한다**
            # (세션 단위 규칙). 그때 통하는 것은 --adopt 다 — 될 리 없는 명령을 적어
            # 두면 사람이 그대로 해 보고 막힌다(2026-08-06 실측).
            fix = ("python ai_claim.py --who %s --free %s" % (c["who"], c["lock"])
                   if c.get("mine") else
                   "python session_handoff.py --adopt   (죽은 세션 것이라 --free 는 거부된다)")
            out.append(("★ '%s' 점유가 **죽은 세션의 잔재**로 보인다 — %s (%s · %s)"
                        % (c["lock"], why, c["who"], c["why"]), fix))
        else:
            out.append(("%s 가 '%s' 를 잡고 있다(%s분, 살아 있음) — 배타 작업은 피할 것"
                        % (c["who"], c["lock"], m),
                        "조회·분석으로 돌리거나 상대가 놓을 때까지 기다린다"))
    # 수집이 밀린 것은 **조용한 사고**다. 화면은 멀쩡히 숫자를 보여 주는데 그 숫자가
    # 원본을 못 따라간 값이다(2026-08-06 8/5 돌발AS 1건 사건). 막힌 것 맨 앞에 둔다.
    for f in st.get("수집신선도") or []:
        if not f.get("밀림"):
            continue
        # ★ 색인이 그만큼 안 돌았으면 **밀렸다고 확언하지 않는다**(분담판 [116]).
        #   조치가 정반대다 — 한쪽은 '내보내기를 또 하라'이고 실제로 필요한 것은
        #   '회차를 돌려라'다. 틀린 지목은 못 잡는 것보다 나쁘다([172]).
        if f.get("색인탓") is True:
            out.append(("%s 이 %d일 밀린 것으로 **보이는데** 원본 색인이 %.1f일째 "
                        "안 돌았다 — 그 침묵이 밀림으로 읽혔을 수 있다(확인 못 함). "
                        "파일이 이미 들어와 있어도 이렇게 나온다"
                        % (f["이름"], f["밀린일"], f.get("색인나이") or 0),
                        "python source_index.py    # 또는 09:35 원본정리 회차"))
            continue
        tail = ""
        # ★ **`.get()` 은 칸이 없어도 None 을 준다** — 그래서 '못 읽음'과 '읽었고
        #   색인 탓이 아님'이 한 낱말로 뭉쳐 있었다([247]). 칸이 **있는지**로 먼저 가른다.
        #   ⚠ 옛 스냅샷에는 이 칸이 없다 — 그때는 아무 말도 안 한다(쓰는 쪽만 고치면
        #   옛 기록이 여전히 읽는 쪽을 속인다 · [212]).
        if "색인탓" in f:
            if f["색인탓"] is None:
                # ★ 못 읽은 것을 '이상 없음'으로도 '밀림 확정'으로도 치지 않는다([169]).
                tail = " (원본 색인 나이를 못 읽어 색인 탓인지 못 갈랐다)"
            elif f["색인탓"] is False:
                # ★ 조치가 갈린다([289]) — 색인은 최신이라 돌려 봐야 안 바뀐다.
                tail = (" (원본 색인은 %.1f일 전에 돌았다 — 색인 탓이 아니라 원본이 "
                        "안 들어온 것이다)" % (f.get("색인나이") or 0))
        out.append(("★ %s 수집이 밀렸다 — 최신 %s (%d일 전, 한도 %d일). "
                    "지금 화면·보고 숫자는 그만큼 **적게** 나온다%s"
                    % (f["이름"], f["최신"], f["밀린일"], f["한도"], tail),
                    f["되살리는법"]))
    # ★ **살아 있는 파일이 아니라 스냅샷에서 읽는다**(2026-08-20). 여기만 홀로
    #   `daily_step_now()` 로 디스크를 직접 봤는데, 그러면 `blockers(st)` 가
    #   **그날 그 기계가 무엇을 했느냐**에 따라 달라져 합성 st 로 부르는 검증
    #   셋(t111·t125·t138)이 회차가 단계를 건너뛴 날 한꺼번에 죽었다. 나머지
    #   경보는 전부 st 에서 읽는다 — 여기도 같게 맞춘다([162]). 담는 것은
    #   snapshot() 이고, 이 함수는 판단만 한다.
    skipped = st.get("건너뜀") or []
    if skipped:
        # ★ '완주'와 '다 했다'는 다른 말이다(분담판 [82]). 예산을 넘으면 남은 단계를
        #   건너뛰고 완주시키는데([180] — 그것이 옳다), 그 사실이 리포트 깊숙한 한 줄로만
        #   남아 아무 화면에도 안 떴다. 실측 2026-08-10·08-11 세 회차가 그렇게 지나갔다.
        out.append(("일일자동대조가 '완주'로 끝났지만 **%d단계를 건너뛰었다**(회차 예산 초과) — "
                    "그 단계들은 그 회차에 **안 돌았다**: %s%s"
                    % (len(skipped), ", ".join(skipped[:4]),
                       " 외 %d개" % (len(skipped) - 4) if len(skipped) > 4 else ""),
                    "다음 회차가 이어서 한다 — 급하면 python daily_run.py"))
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
    # ★ 맨 위 — 밴드 글이 **고쳐졌다**는 소식 (2026-08-08 지시).
    #   밀림('못 받은 글이 있다')과 다른 종류의 사고다. 글은 다 받았는데 **받은 뒤에
    #   내용이 바뀐** 것이라, 개수·날짜 어디를 봐도 티가 안 난다. 그래서 아래 어느
    #   칸도 아니고 맨 위다. 사람이 `--ack` 로 내리기 전까지 남는다.
    rc = st.get("밴드재수집") or {}
    if rc:
        chg, new = rc.get("바뀐글") or [], rc.get("새글") or []
        _rg = rc.get("되돌아감")
        if _rg:
            out.append(("[P1] 밴드 글 %d건이 **완료 -> 완료 아님으로 되돌아갔다** — 밴드 쪽 완료 근거가 사라졌다(원장·앱DB가 받쳐 주는지 따로 본다): %s"
                        % (len(_rg), ", ".join(r.get("글", "") for r in _rg[:6])),
                        "python band/recollect.py --print"))

        L += ["## ★ 밴드 글이 바뀌었다 — 최근 %d일 재수집 (%s)"
              % (rc.get("창일수", 30), rc.get("회차", "")), "",
              "- **고쳐진 글 %d건** · 새로 들어온 글 %d건. "
              "밴드는 상태가 바뀌면 **같은 번호의 글을 고쳐** 다시 올린다 —"
              " 개수도 날짜도 그대로라 여기 말고는 티가 안 난다." % (len(chg), len(new)), ""]
        if chg:
            L += ["| 밴드 | 글번호 | 작성일 | 무엇이 바뀌었나 |", "|---|---:|---|---|"]
            det = {c.get("글"): c.get("어떻게", "") for c in (rc.get("변경상세") or [])}
            for c in chg[:15]:
                key = "%s/%s" % (c.get("밴드ID"), c.get("글번호"))
                L.append("| %s | %s | %s | %s |"
                         % (c.get("밴드", ""), c.get("글번호", ""), c.get("작성일", ""),
                            det.get(key) or (c.get("요약") or "본문 바뀜")))
            if len(chg) > 15:
                L.append("| … | | | 그 밖 %d건 |" % (len(chg) - 15))
            L.append("")
        L += ["> 대조·보고서가 이 글들을 근거로 쓰고 있었다면 **다시 뽑아야 한다**.",
              "> 전체는 `python band/recollect.py --print` · "
              "확인했으면 `python band/recollect.py --ack` 로 이 칸을 내린다.", ""]
    wf = st.get("업무흐름") or {}
    if wf.get("바뀜"):
        L += ["## ★ 업무 단계 정의가 바뀌었다 (%s)" % (wf.get("시각") or ""), "",
              "- 돌발AS·정기점검 화면의 **단계 선택지**가 이 정의를 읽는다. "
              "바뀌면 화면이 조용히 따라간다 — 그래서 여기 올린다.", ""]
        L += ["- %s" % x for x in (wf.get("무엇이") or [])]
        L += ["", "> 지금 정의는 `python work_flow.py` · "
                  "확인했으면 `python work_flow.py --ack` 로 이 칸을 내린다.", ""]
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
    fr = st.get("수집신선도") or []
    if fr:
        L += ["## 원본 수집이 어디까지 들어왔나", "",
              "| 원본 | 최신 | 밀린 일수 | 한도 |", "|---|---|---:|---:|"]
        for f in fr:
            late = "?" if f["밀린일"] is None else str(f["밀린일"])
            # ★ 중단은 '조용함'보다 **센 사실**이라 먼저 본다 — 조용함은 "그 위에
            #   새 글이 없음을 확인했다" 이고 중단은 "앞으로 안 받는다" 다([289]).
            tag = " ★밀림" if f.get("밀림") else (
                " (수집 중단)" if f.get("수집중단")
                else (" (조용함)" if f.get("조용함") else ""))
            L.append("| %s%s | %s | %s | %d |"
                     % (f["이름"], tag, f["최신"], late, f["한도"]))
        offs = [f for f in fr if f.get("수집중단")]
        quiet = [f for f in fr if f.get("조용함")]
        L += ["", "> 밴드·이카운트는 **사람 로그인**이 있어야 긁힌다(절대규칙 3).",
              "> 밀려 있으면 화면 숫자가 그만큼 적게 나온다 — 숫자를 의심하기 전에 여기부터 본다."]
        if quiet:
            # 날짜만 보면 밀린 것처럼 보이지만 받을 것이 없는 밴드 — 왜 안 긁어도 되는지 적는다.
            # 이 줄이 없으면 다음 세션이 또 없는 번호를 긁는다(2026-08-07 사고).
            L += ["> **조용함**: 최신 글이 오래됐지만 그 위로 새 글이 없음을 확인한 것이다 —"
                  " 긁을 것이 없다. 없는 번호를 긁으면 쓰레기가 캐시에 들어간다."]
            L += ["> · %s — %s" % (f["이름"], f["조용함"]) for f in quiet]
        if offs:
            # ★ **조용히 빼지 않는다**([169]) — 밀림 경보에서는 내렸지만 왜 안 받는지는
            #   여기 남는다. 안 적으면 다음 사람이 '밴드가 왜 이렇게 오래됐지' 를 물을
            #   근거가 없고, 되돌리는 법도 모른다.
            L += ["> **수집 중단**: 사람이 멈춘 것이라 안 들어오는 것이 정상이다 —"
                  " 밀림 경보에서 뺐다. 다시 켜려면"
                  " `python band/collect_switch.py --resume`."]
            L += ["> · %s — %s" % (f["이름"], f["수집중단"]) for f in offs]
        why = [f for f in fr if f.get("밀림") and f.get("근거")]
        if why:
            # ★ '밀림'만 적고 이유를 안 적으면 사람이 없는 번호를 긁으러 간다([217]).
            #   근거가 낡았다·추월됐다는 것은 "그 위에 몇 개가 있다"는 뜻이 아니다 —
            #   존재 확인용 몇 건만 찔러 보면 된다(붙여넣기 파일이 이미 그렇게 나온다).
            L += ["> **밴드 근거 상태** — 아래는 '없음 확인' 근거를 못 믿는 이유다."
                  " 그 위 번호가 몇 개인지는 **모른다**; 계획은 존재 확인용 몇 건만 담는다."]
            L += ["> · %s — %s" % (f["이름"], f["근거"]) for f in why]
        ct = st.get("밴드오염") or {}
        if sum(ct.values()):
            # 상태 한 줄 — 할 일이 아니다. 재수집하지 말 것(삭제된 글, 표본 3/3 리다이렉트).
            L += ["> **오염(삭제 판정)**: %s — 없는 번호를 열면 밴드가 피드를 돌려줘 남의"
                  " 본문이 잡혔던 기록. 삭제된 글이라 되살릴 내용이 없다. 대조 대상이"
                  " 아니며 **재수집하지 않는다**(검증 [135])."
                  % ", ".join("%s %d건" % (k, v) for k, v in sorted(ct.items()))]
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


def _swap_in(tmp, dst):
    """tmp 를 dst 로 **원자적으로** 갈아끼운다 — 새로 만들지 않고 [171] 의 그 도구를 빌린다([162]).

    ★ 2026-09-04 실사고: 09:50 회차가 0단계 관문에서 죽었다 —
      `json.decoder.JSONDecodeError: Expecting , delimiter: line 386 column 301`.
      범인은 이 스냅샷을 **쓰는 도중에 읽은 것**이다. `open(...,"w")` 는 먼저 통째로
      비우고 쓰므로 그 사이에 읽는 쪽은 반쯤 쓰인 파일을 본다. 그때 오류가 가리킨
      char 18881 은 그 순간 파일의 line 386 col 301 에 **글자까지 맞아떨어졌다**.

      값이 관문 하나가 아니다 — 관문은 `daily_run` 의 **0단계**라 여기서 죽으면
      그날 대조가 통째로 안 돈다(접수취소·객관완료·청구상태·오기입·사실대조·캠프 담당자).
      관문은 이 파일을 **열여섯 곳**에서 실측 증거로 읽고([247]) 워치독은 30분마다
      다시 쓰므로 25분 도는 관문과 **반드시 겹친다.** 그런데 겹친다고 늘 죽는 것이
      아니라 **읽는 그 순간**과 겹쳐야 죽는다 — 그래서 더 조용하고 더 나쁘다([169]).

    ⚠ **맨몸 `os.replace` 로는 모자란다.** 임시폴더 재현(진짜 파일은 한 글자도 안
      건드렸다 · [247])에서 옛 방식은 1분에 **토막 읽기 186회**였고, 그것을 그냥
      `os.replace` 로 바꾸자 이번에는 **`PermissionError [WinError 5]`** 가 났다 —
      윈도우는 읽는 쪽이 물고 있는 파일을 갈아끼우지 못한다([171] 이 밴드 캐시에서
      겪은 그 자리다). 그래서 물러서며 다시 거는 `swap_in` 을 빌린다.

    ★ 못 갈아끼우면 **옛 스냅샷을 그대로 둔다** — 낡았어도 온전한 것이 반쯤 쓰인
      것보다 낫다. 대신 조용히 넘어가지 않고 말한다([169]). `.tmp` 도 안 지운다
      ([171]) — 애써 만든 새 내용을 버리는 것보다 사람이 그것을 보게 두는 편이 낫다.
    """
    try:
        sys.path.insert(0, os.path.join(BASE, "band"))
        from convert_dump import swap_in
    except Exception:
        swap_in = None
    try:
        if swap_in:
            swap_in(tmp, dst)
        else:
            os.replace(tmp, dst)
        return True
    except Exception as e:
        print("  ! 인계 스냅샷을 갈아끼우지 못했습니다(%s: %s) — 옛 것을 그대로 둡니다. "
              "새 내용은 %s 에 있습니다" % (type(e).__name__, str(e)[:80], os.path.basename(tmp)))
        return False

def write_snapshot(st, for_sol=False):
    """reports/세션인계.md|json 갱신 — 워치독(--snapshot)과 --adopt 가 같은 것을 남긴다."""
    os.makedirs(REPORT_DIR, exist_ok=True)
    # ★ 비우고 쓰지 않는다 — tmp 에 다 쓴 뒤 한 번에 갈아끼운다(위 _swap_in 참조).
    #   .md 도 같이 한다: 반쯤 쓰인 인계 문서는 사람이 "인계가 사라졌다"로 읽는다.
    md = os.path.join(REPORT_DIR, "세션인계.md")
    with open(md + ".tmp", "w", encoding="utf-8") as f:
        f.write(to_md(st, for_sol=for_sol))
    _swap_in(md + ".tmp", md)
    js = os.path.join(REPORT_DIR, "세션인계.json")
    with open(js + ".tmp", "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=1, default=str)
    _swap_in(js + ".tmp", js)
    return blockers(st, for_sol=for_sol)


def adopt(who="claude"):
    """다른 계정·다른 창이 이 작업을 **이어받을 때** 한 번 부른다 (2026-08-06 지시).

    계정이 바뀌면 세션 식별자(sid)도 프로세스도 전부 바뀐다. 그때 남아 있는 것들은
    새 세션 입장에서 "남의 것"으로 보여, 규칙대로면 손을 못 댄다 — 그래서 아무도
    풀지 못하는 점유가 원장을 영원히 막는다. 반대로 무턱대고 다 풀면 **정말 살아
    있는 옆 세션**의 작업을 가로챈다.

    그래서 기준은 하나다: **주인 프로세스가 죽었다는 증거(pid)가 있을 때만** 손댄다.
      0. 워크트리면 본체 상태와 먼저 잇는다 — **큐를 읽기 전에** 해야 한다.
         순서가 뒤집히면 워크트리의 빈 큐를 보고 "0건" 이라 답한다(조용히 틀린 답).
      1. 죽은 세션의 점유만 회수 — 살아 있는 것은 그대로 두고 알려만 준다
      2. 입력 큐 → DB 흡수 (엑셀은 열지 않는다. 반영은 11:00·15:00 회차 그대로)
      3. 수집이 밀렸는지 확인 — 밴드·이카운트는 사람 로그인이 먼저다
      4. 인계 문서 갱신
    사람 판단이 필요한 것(미푸시·임시파일·로그인)은 고치지 않고 목록으로 남긴다.
    """
    import ai_claim as C
    steps, freed, kept = [], [], []
    try:
        import worktree_state as W
        if W.is_worktree():
            r = W.apply()
            made = ["%s(%s)" % (rel, how) for rel, how in r.get("이음", [])]
            kept_wt = ["%s — %s" % (rel, why) for rel, why in r.get("건너뜀", [])]
            steps.append(("워크트리 → 본체 잇기",
                          ", ".join(made) if made else "이미 다 이어져 있음"))
            if kept_wt:
                steps.append(("워크트리에 따로 있어 그대로 둔 것", ", ".join(kept_wt)))
    except Exception as exc:
        steps.append(("워크트리 → 본체 잇기", "실패: %s" % str(exc)[:80]))
    with C.state_guard():
        d = C._load_unlocked()
        for lock, claim in list(d.items()):
            if C._is_mine(claim, who) or C._is_dead(claim):
                d.pop(lock, None)
                freed.append("%s(%s)" % (lock, (claim or {}).get("who", "?")))
            else:
                kept.append("%s(%s · %s)" % (lock, (claim or {}).get("who", "?"),
                                             (claim or {}).get("why", "")))
        C._save_unlocked(d)
    steps.append(("죽은 세션 점유 회수", ", ".join(freed) if freed else "없음"))
    if kept:
        steps.append(("살아 있어 그대로 둔 점유", ", ".join(kept)))

    try:
        import ledger_db
        n = ledger_db.intake_json()
        steps.append(("입력 큐 → DB", "%d건 흡수 (엑셀 반영은 11:00·15:00 회차)" % n))
    except Exception as exc:
        steps.append(("입력 큐 → DB", "실패: %s" % str(exc)[:80]))

    st = collect()
    late = [f for f in st.get("수집신선도") or [] if f.get("밀림")]
    steps.append(("수집 신선도",
                  " · ".join("%s %s(%d일 밀림)" % (f["이름"], f["최신"], f["밀린일"])
                             for f in late) if late else "밀린 원본 없음"))
    write_snapshot(st)
    steps.append(("인계 문서", os.path.join(REPORT_DIR, "세션인계.md")))

    print("# 이어받기 준비 — 다른 계정/다른 창에서 시작할 때", "")
    for name, detail in steps:
        print("- %s: %s" % (name, detail))
    left = blockers(st)
    print("\n## 아직 사람이 해야 하는 것 (%d)" % len(left))
    for why, how in left:
        print("- %s\n  → %s" % (why, how))
    if not left:
        print("- 없음. 바로 새 작업을 시작해도 된다.")
    return st


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--for-sol", action="store_true", help="Terra -> Sol 검토 관문도 함께 확인")
    ap.add_argument("--snapshot", action="store_true", help="상태를 파일로 남긴다(워치독용)")
    ap.add_argument("--check", action="store_true", help="새 세션 시작 시 읽는다")
    ap.add_argument("--adopt", action="store_true",
                    help="다른 계정·다른 창이 이어받을 때 — 죽은 점유 회수·큐 흡수·인계 갱신")
    ap.add_argument("--who", default="claude", help="--adopt 주체(claude|codex)")
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
    if a.adopt:
        adopt(a.who)
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
        bl = write_snapshot(st, for_sol=a.for_sol)
        print("세션인계 갱신 — 걸린 것 %d건 · 관리대장 v%s" % (len(bl), st["원장"].get("버전", "?")))
        return 0
    print(md)
    return 0


if __name__ == "__main__":
    sys.exit(main())
