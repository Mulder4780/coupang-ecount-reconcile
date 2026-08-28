# -*- coding: utf-8 -*-
"""
daily_run.py — 쿠팡 업무 자동대조 에이전트 (일일 오케스트레이터)
==================================================================
매일 1회(09:50, 류지영 입력 종료 후) 전체 파이프라인을 안전한 순서로 실행하고
reports/종합리포트_*.md 한 장으로 요약한다. Windows 작업 스케줄러에 daily_run.bat 등록 시 완전 자동.

원칙:
  - 0단계 합성검증(ALL GREEN) 실패 시 전체 중단 (사용자 상시 지시)
  - ERP 쓰기(--post)는 절대 자동 실행하지 않음 — 전송 대기 건수만 보고
  - 각 단계는 데이터가 없으면 조용히 건너뜀(스킵 사유 기록) — 있는 데이터만큼 검증
"""
import sys, os, glob, json, re, subprocess, time, uuid, hashlib
from datetime import datetime
from operation_window import input_window_label, is_input_window

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
REPORT_DIR = os.path.join(ROOT, "reports")
# ★ 관문(합성검증)은 이 표시를 보고 **회차가 부른 것인지** 안다([412]).
#   사람이 손으로 돌린 관문은 회차가 도는 중이면 물러난다 — 둘이 같이
#   Z: 를 훑으면 서로를 25분 제한 밖으로 밀어내 그날 대조가 통째로 안 돈다
#   (2026-08-24 실사고). **이 표시가 빠지면 회차의 관문이 스스로 물러나
#   매일 아침 대조가 안 돈다** — 고치려던 것보다 나쁘다.
ENV = {**os.environ, "PYTHONIOENCODING": "utf-8",
       "COUPANG_GATE_OWNER": "daily_run"}
RUN_LOCK = os.path.join(REPORT_DIR, ".daily_run.lock")
# 조율 표에서 이 회차를 부르는 이름. **양보한 쪽과 완주한 쪽이 같은 이름을 써야**
# `coordinate.audit()` 이 "양보했는데 주인이 끝냈나"를 이을 수 있다(둘이 갈리면
# 모든 양보가 영영 '헛양보'로 보인다 — 낱말이 어긋나면 한 건도 안 걸린다).
COORD_JOB = "일일대조"

# collect_all과 autopilot의 "한 회차 진척·아직 남음" 계약. 0(완료)이나
# 1(실패)로 바꾸면 큐가 일찍 닫히거나 정상 증분을 실패로 센다([217]).
INCREMENTAL_RETURN_CODE = 75


def _pid_alive(pid, born_before=None, pid_started_at=None):
    """같은 PC의 프로세스가 살아 있는가. 죽은 잠금만 안전하게 회수한다.

    ★ 판정은 pid_alive.py 한 곳에서 한다 (2026-08-06 실사고 · 검증 [121]).
      여기 있던 옛 판정은 윈도우에서 **이미 끝난 프로세스도 살아 있다고** 했다
      (OpenProcess 는 종료된 프로세스에도 핸들을 준다). 그 탓에 잠금이 스스로 풀릴
      길이 없어져 daily_run 이 밤새 한 번도 못 돌았다.
      모르면(None) '살아 있다'로 본다 — 남의 회차를 밀어내는 쪽이 더 위험하다.
      다만 pid 자체가 없거나 망가진 잠금 파일은 회수 대상이다.

    ★ `born_before` 는 잠금이 쓰인 시각이다 (2026-08-11 실사고 · 검증 [210]).
      회차가 죽은 뒤 윈도우가 그 pid 를 **다른 프로그램에 재사용**하면 번호만 보고
      '살아 있다'가 된다 — 그날 quick_share_server 가 죽은 회차의 pid 를 물려받아
      잠금이 다섯 시간 동안 안 풀렸다. 주인은 잠금을 쓰기 전에 떠 있었을 수밖에
      없으므로, 잠금 시각보다 뒤에 태어난 프로세스는 주인이 아니다.
    """
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False
    try:
        import pid_alive
    except Exception:
        return True                      # 판정 못 하면 건드리지 않는다
    return pid_alive.owner_alive(
        pid, pid_started_at=pid_started_at, born_before=born_before
    ) is not False


def _process_identity():
    """One shared owner fingerprint shape; unknown creation time stays null."""
    try:
        import pid_alive
        return pid_alive.identity()
    except Exception:
        return {"pid": os.getpid(), "pid_started_at": None}


def acquire_run_lock(path=RUN_LOCK):
    """프로세스 간 단발 잠금. 성공 시 소유 토큰, 중복 실행이면 None."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    token = f"{os.getpid()}:{time.time_ns()}:{uuid.uuid4().hex}"
    payload = {
        **_process_identity(),
        "token": token,
        "started_at": datetime.now().astimezone().isoformat(timespec="seconds"),
    }
    for _attempt in range(2):
        try:
            fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            try:
                with open(path, encoding="utf-8") as f:
                    owner = json.load(f)
            except (OSError, ValueError, TypeError):
                owner = {}
            # O_EXCL로 파일을 만든 직후 JSON을 쓰는 아주 짧은 동안에는 다른 실행이
            # 빈/부분 파일을 볼 수 있다. 최근의 손상 파일은 선점 중으로 보고 건드리지 않는다.
            if not owner:
                try:
                    if time.time() - os.path.getmtime(path) < 60:
                        return None
                except OSError:
                    return None
            # 잠금이 쓰인 시각(started_at → 없으면 파일 mtime)보다 뒤에 태어난
            # 프로세스는 주인이 아니다 — pid 재사용 오판 방지(검증 [210]).
            born = None
            try:
                born = datetime.fromisoformat(str(owner.get("started_at"))).timestamp()
            except (TypeError, ValueError):
                try:
                    born = os.path.getmtime(path)
                except OSError:
                    pass
            if _pid_alive(
                owner.get("pid"),
                born_before=born,
                pid_started_at=owner.get("pid_started_at"),
            ):
                return None
            try:
                os.unlink(path)
            except OSError:
                return None
            continue
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
            f.write("\n")
        return token
    return None


def release_run_lock(token, path=RUN_LOCK):
    """자기가 만든 잠금만 놓는다. 다른 실행의 잠금을 지우지 않는다."""
    try:
        with open(path, encoding="utf-8") as f:
            owner = json.load(f)
        if owner.get("token") == token:
            os.unlink(path)
    except (OSError, ValueError, TypeError):
        pass


def has_inbox_kind(kind):
    """무작위 파일명도 놓치지 않도록 엑셀 머리글 내용으로 원천을 찾는다."""
    try:
        from inbox_scan import pick
        return bool(pick(kind))
    except Exception:
        return False


def _source_files(dir_fn, extensions, name_prefixes=()):
    """source_dirs의 정본 폴더를 재귀 탐색한다. 투입함은 분류 후에만 읽는다."""
    try:
        import source_dirs
        folders = getattr(source_dirs, dir_fn)()
    except Exception:
        return []
    out, seen = [], set()
    for folder in folders:
        for base, _dirs, files in os.walk(folder):
            for name in files:
                low = name.lower()
                if extensions and not low.endswith(extensions):
                    continue
                if name_prefixes and not low.startswith(name_prefixes):
                    continue
                path = os.path.join(base, name)
                key = os.path.normcase(os.path.abspath(path))
                if key not in seen:
                    seen.add(key)
                    out.append(path)
    return out


# 되풀이하면 안 되는 단계 — 큐에 넣거나 파일을 옮기거나 원장을 건드리는 것들.
# ★ 호출부마다 retry=0 을 적게 하지 않는다 — 새 단계를 넣는 사람이 반드시 잊는다.
#   여기서 **인자를 보고** 자동으로 정한다.
NO_RETRY_FLAGS = ("--queue", "--apply", "--intake", "--force")
NO_RETRY_NAMES = ("cloud_queue_sync.py", "ledger_db.py", "workbook_patch.py",
                  "ledger_writer.py", "expand_rows.py",
                  # 게시는 git commit·push 다. 실패한 자리를 모르는 채 다시 하면
                  # 빈 커밋이나 반쯤 올라간 상태가 겹친다 — 다음 회차에 맡긴다.
                  "cloud_publish.py")


def _retryable(args):
    joined = " ".join(str(a) for a in args)
    if any(f in args for f in NO_RETRY_FLAGS):
        return False
    return not any(n in joined for n in NO_RETRY_NAMES)


PROGRESS = os.path.join(REPORT_DIR, ".daily_run.progress.json")
ROUND_BUDGET_MIN = int(os.environ.get("COUPANG_ROUND_BUDGET_MIN", "150"))
_ROUND_T0 = [None]          # 회차 시작 시각 — main() 이 채운다
_OVER_BUDGET = [False]


#: 이 회차가 예산 초과로 **건너뛴 단계 이름들**.  진행 파일에 같이 실려
#: 인계가 읽는다 — 리포트 깊숙한 한 줄은 아무도 안 본다(분담판 [82]).
_SKIPPED = []


def note_progress(step, state, extra=None):
    """**단계마다** 어디까지 왔는지 디스크에 남긴다 (2026-08-09 지시).

    ★ 32시간 미완주의 진짜 문제는 '느리다'가 아니라 **어디서 멈췄는지 아무도 모른다**는
      것이었다. 종합리포트는 **맨 끝에 한 번** 쓰이므로 회차가 완주하지 못하면
      기록이 **한 줄도 안 남는다.** 스케줄러는 그동안 '성공'이라 적는다.
      그래서 화면은 '08-08 01:38 — 32시간째 미완주'만 보여 주고 이유를 못 댔다.
    이제 이 파일 하나로 "지금 몇 번째 단계 · 무엇 · 언제 시작"이 항상 남는다.
    죽어도 남는다 — 그게 요점이다.
    """
    try:
        os.makedirs(REPORT_DIR, exist_ok=True)
        previous = {}
        if os.path.exists(PROGRESS):
            try:
                with open(PROGRESS, encoding="utf-8") as fh:
                    previous = json.load(fh)
            except (OSError, ValueError):
                previous = {}

        # ★ 이전 단계의 자유 필드를 update()로 이어받지 않는다. 예전 구현은
        # `시간초과:true`·`결과:false`·`명령:...`가 다음 정상 단계와 회차 끝에도
        # 남아, 화면이 이미 끝난 실패를 현재 실패처럼 표시했다. 회차 전체에서
        # 이어갈 값은 끝난 단계 목록뿐이고, 나머지는 매 상태의 새 사실이어야 한다.
        done = list(previous.get("끝난단계") or [])[-60:]
        now_dt = datetime.now().astimezone()
        now_iso = now_dt.isoformat(timespec="seconds")
        # ★ **단계마다 얼마나 걸렸나**를 같이 남긴다 (2026-08-12, `[228]` 이 드러낸 것).
        #   `[180]` 이 "어디서 멈췄나"를 남기게 했지만 **"어디가 오래 걸렸나"** 는 아직
        #   아무 데도 없었다. 그래서 회차가 292분을 쓰고 스케줄러 제한(PT3H)에 매일
        #   강제 종료되는데도 **범인 단계를 짐작으로밖에 말할 수 없었다.**
        #   회차 예산(`over_budget`)은 단계 **사이**에서만 보므로 한 단계가 길면 그냥
        #   지나친다 — 그 한 단계가 누구인지 모르면 예산을 조여도 소용이 없다.
        _new_round = (step == "(회차 시작)")
        t0 = previous.get("단계시작") if previous.get("단계") == step else None
        t0 = t0 or now_iso
        try:
            spent = int((now_dt - datetime.fromisoformat(t0)).total_seconds())
        except (TypeError, ValueError):
            t0, spent = now_iso, 0
        cur = {
            **_process_identity(),
            "단계": step,
            "상태": state,
            "시각": now_iso,
            "단계시작": t0,
            "단계경과초": max(0, spent),
            "끝난단계": done,
            # 문자열 목록(`끝난단계`)은 읽는 쪽이 있으므로 모양을 안 바꾼다 — 시간은 옆에 따로 쌓는다.
            # ★ 새 회차면 지난 회차의 시간 기록을 **비운다**. 안 비우면 `[-60:]` 상한에
            #   걸려 **여러 회차가 섞이고**(실측: 한 파일에 합성검증이 5회차분 5번 찍혀
            #   있었다) '이 회차에서 무엇이 오래 걸렸나'에 남의 회차 값이 답한다.
            #   판단을 부르는 쪽에 두지 않는 이유: 인자를 빠뜨리면 **조용히** 안 비워진다(`[381]`).
            "단계기록": [] if _new_round else list(previous.get("단계기록") or [])[-60:],
            # ★ `느린단계` 도 **이어받는다**. 예전엔 `state == "끝"` 일 때만 담아서,
            #   회차의 마지막 기록인 `(회차 끝)`(상태 '완주'/'실패')에서 이 키가 통째로
            #   사라졌다 — 그런데 **인계는 회차가 끝난 뒤에 읽는다.** 그래서 `[228]` 이
            #   만든 '범인 단계를 댄다'가 한 번도 안 돌았고, 40단계를 건너뛴 회차조차
            #   무엇이 예산을 먹었는지 말하지 못했다(`[169]` — 오류도 안 나고 조용하다).
            "느린단계": [] if _new_round else list(previous.get("느린단계") or []),
            # ★ 예산 초과로 **건너뛴 단계**(분담판 [82]) — 회차는 '완주'로 끝나지만
            #   그 단계들은 오늘 안 돌았다. 인계가 이것을 읽어 말한다.
            "건너뜀": list(_SKIPPED)[-40:],
        }
        if _ROUND_T0[0]:
            cur["회차시작"] = _ROUND_T0[0].astimezone().isoformat(timespec="seconds")
            cur["경과분"] = round((datetime.now() - _ROUND_T0[0]).total_seconds() / 60, 1)
        cur["예산분"] = ROUND_BUDGET_MIN
        if extra:
            cur.update(extra)
        if state == "끝":
            cur["끝난단계"] = (list(cur.get("끝난단계") or []) + [step])[-60:]
            cur["단계기록"] = (list(cur.get("단계기록") or [])
                             + [{"단계": step, "초": cur["단계경과초"]}])[-60:]
            # 오래 걸린 순으로 다섯 — 화면·인계가 이 한 줄만 읽어도 범인을 댈 수 있다.
            cur["느린단계"] = sorted(cur["단계기록"], key=lambda r: -int(r.get("초") or 0))[:5]
        tmp = f"{PROGRESS}.{os.getpid()}.tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(cur, fh, ensure_ascii=False)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp, PROGRESS)
    except Exception:
        pass          # 진행 기록을 남기려다 회차를 세우지 않는다


def over_budget():
    """회차가 예산을 넘었나 — 넘으면 **남은 단계를 건너뛰고 완주시킨다**.

    ★ 왜 중단이 아니라 완주인가: 회차가 끝나야 종합리포트가 써지고 잠금이 풀리고
      **다음 회차가 돈다.** 예산 없이 무한정 끌면 다음 회차가 "이미 실행 중"으로
      조용히 건너뛰고, 그것이 이틀·사흘 쌓여 '32시간 미완주'가 된다.
      반쯤이라도 완주한 회차가, 영원히 안 끝나는 회차보다 낫다.
    """
    if not _ROUND_T0[0]:
        return False
    if (datetime.now() - _ROUND_T0[0]).total_seconds() / 60 >= ROUND_BUDGET_MIN:
        _OVER_BUDGET[0] = True
        return True
    return False


STEP_FLOOR_SEC = int(os.environ.get("COUPANG_STEP_FLOOR_SEC", "120"))


def budget_left_sec():
    """회차 예산이 몇 초 남았나. 예산 시계가 없으면 None(그때는 줄이지 않는다)."""
    if not _ROUND_T0[0]:
        return None
    return ROUND_BUDGET_MIN * 60 - (datetime.now() - _ROUND_T0[0]).total_seconds()


def fit_timeout(timeout):
    """단계의 시간 제한을 **남은 예산 안으로** 줄인다 (2026-08-12 · 분담판 [38]).

    ★ 예산이 있는데도 회차가 292분을 돌아 **작업 스케줄러의 3시간 제한에 나무째
      끊겼다**(0xC000013A · `일일자동대조`·`원본자료자동정리` 둘 다). 예산은 150분인데
      어떻게 292분인가 — `over_budget()` 은 단계 **사이**에서만 보기 때문이다.
      145분째에 시작한 단계가 제 시간 제한 60분을 그대로 들고 가면 205분이 되고,
      그 뒤 정리 단계까지 붙으면 292분이 된다. 예산이 **경계가 아니라 권고**였다.
    ★ 그래서 **범인 단계를 지목하지 않고** 고칠 수 있다. 어느 단계가 오래 걸리는지는
      아직 모르고(단계별 시간 기록은 어제 붙었다), 모르는 채로 제한시간을 손대는 것은
      짐작이다. 그러나 "어떤 단계도 남은 예산보다 오래 받을 수 없다"는 **모든 단계에
      참인 규칙**이라 짐작이 아니다. 범인이 누구든 회차는 예산 안에서 끝난다.
    ★ 끊기는 것과 예산 초과는 결과가 다르다. 끊기면 잠금이 남아 **다음 회차가 조용히
      건너뛰고**(스케줄러는 '성공'이라 적는다) 리포트가 한 줄도 안 써진다. 예산으로
      끝내면 완주로 남고 못 한 몫은 다음 회차가 이어서 한다 — 그게 `[180]` 의 뜻이다.
    ★ **바닥을 둔다**(STEP_FLOOR_SEC). 예산이 다 됐다고 3초를 주면 그 단계는 시작하자마자
      죽어 '실패'로 적히는데, 그건 사실이 아니다(시간을 안 준 것이다). 예산이 바닥나면
      애초에 `over_budget()` 이 단계를 건너뛰므로 여기 오는 것은 자투리가 남은 때뿐이다.
    """
    left = budget_left_sec()
    if left is None:
        return timeout, False
    room = max(STEP_FLOOR_SEC, int(left))
    if timeout <= room:
        return timeout, False
    return room, True


def run(name, args, timeout=600, retry=None):
    """한 단계를 돌린다. 실패하면 **한 번만** 쉬었다 다시 해 본다.

    ★ 2026-08-05 — 14:02 회차에서 5단계가 실패했는데, 손으로 다시 돌리니 전부 정상이었다.
      원인은 결함이 아니라 **경합**이다(Z: 대량 보관 작업·엑셀 점유와 겹친 순간).
      이런 것은 한 번 쉬었다 하면 지나간다. 그런데 그동안은 그대로 '실패'로 남아,
      진짜 결함(오늘 잡은 stmt_link 회귀 같은 것)이 잡음에 묻혔다.
      ※ 되풀이해도 되는 것만 재시도한다 — 큐·반영·파일 이동 단계는 자동으로 제외된다.
    """
    if over_budget():
        # 예산을 넘었으면 **남은 단계를 건너뛰고** 회차를 끝낸다. 이유를 적어 남긴다 —
        # 조용히 건너뛰면 "돌았는데 왜 결과가 없나"가 된다(그게 지금까지의 증상이었다).
        # ★ **건너뛴 것을 세어 남긴다**(분담판 [82]). [180] 은 리포트에 이유를
        #   적게 했는데 **인계는 리포트를 안 읽는다** — 그래서 회차가 뒤쪽
        #   17단계를 건너뛰고 '완주'로 끝나도 아무 화면에도 안 떴다(실측
        #   2026-08-10·08-11 세 회차). 리포트에만 적으면 아무도 안 본다([169]).
        _SKIPPED.append(name)
        note_progress(name, "건너뜀(예산초과)")
        return {"name": name, "ok": None,
                "out": f"건너뜀 — 회차 예산 {ROUND_BUDGET_MIN}분 초과. "
                       f"다음 회차가 이어서 합니다(완주를 우선합니다)"}
    if retry is None:
        retry = 1 if _retryable(args) else 0
    note_progress(name, "시작", {"명령": os.path.basename(str(args[0])) if args else ""})
    for attempt in range(retry + 1):
        # 시도마다 다시 잰다 — 첫 시도가 예산을 먹었으면 재시도는 더 짧게 받는다.
        fitted, cut = fit_timeout(timeout)
        got = _run_once(name, args, fitted)
        if cut:
            got["out"] = ((got.get("out") or "") +
                          f"\n[예산 맞춤] 시간 제한을 {timeout // 60}분 → {fitted // 60}분으로 줄였습니다"
                          f" — 회차 예산 {ROUND_BUDGET_MIN}분을 넘기면 작업 스케줄러가"
                          f" 회차를 통째로 끊습니다. 못 끝낸 몫은 다음 회차가 이어서 합니다").strip()
        if got.get("returncode") == INCREMENTAL_RETURN_CODE:
            # 실패 재시도나 resolve로 보내지 않는다. 같은 멱등 명령을 영속 큐에 두고
            # watchdog이 다음 30분 회차에 이어 간다.
            deferred = _autopilot_defer(name, args, timeout, got.get("out", ""))
            got["ok"] = None
            got["deferred"] = bool(deferred)
            if deferred:
                got["out"] += "\n자율복구 증분 대기 — 다음 회차가 저장된 자리부터 이어서 합니다"
            note_progress(name, "끝", {"결과": None, "증분계속": True})
            return got
        # ★ **시간 초과는 재시도하지 않는다** (2026-08-09 실측).
        #   재시도는 원래 **경합**을 위한 것이다(Z: 대량 작업·엑셀 점유와 겹친 순간) —
        #   그건 금방 실패하고 한 번 쉬면 지나간다. 그런데 시간 초과는 다르다:
        #   "이 단계는 준 시간보다 오래 걸린다"는 뜻이라 다시 해도 또 넘긴다.
        #   실측: 09:50 회차가 2시간째일 때 collect_all 이 1시간 만에 초과로 죽고
        #   **똑같이 한 시간을 더 쓰는 중**이었다. 회차 예산을 그 하나가 두 번 먹는다.
        if "시간초과" in str(got.get("out", "")):
            got["out"] = (got["out"] + " — 재시도하지 않습니다"
                          "(다시 해도 또 넘깁니다. 다음 회차가 이어서 합니다)")
            deferred = _autopilot_defer(name, args, timeout, got["out"])
            if deferred:
                got["ok"] = None
                got["deferred"] = True
                got["out"] += "\n자율복구 대기열에 저장 — 워치독이 제한 재시도합니다"
            note_progress(name, "끝", {"결과": False, "시간초과": True})
            return got
        if got["ok"] or attempt >= retry:
            if not got["ok"] and attempt:
                got["out"] = (got["out"] + "\n[재시도 후에도 실패]").strip()
            if got["ok"]:
                _autopilot_resolve(name, args)
            else:
                deferred = _autopilot_defer(name, args, timeout, got["out"])
                if deferred:
                    got["ok"] = None
                    got["deferred"] = True
                    got["out"] += "\n자율복구 대기열에 저장 — 자원 복구 뒤 자동 재개합니다"
            note_progress(name, "끝", {"결과": bool(got.get("ok"))})
            return got
        note_progress(name, "재시도")
        time.sleep(20)          # 상대가 파일을 놓을 시간을 준다
    note_progress(name, "끝", {"결과": False})
    return got


def _autopilot_defer(name, args, timeout, output):
    """자율복구 기록 실패가 본 업무 회차를 세우지 않게 하는 얇은 경계."""
    try:
        import autopilot
        return autopilot.defer(name, list(args), int(timeout), str(output or ""))
    except Exception:
        return None


def _autopilot_resolve(name, args):
    try:
        import autopilot
        autopilot.resolve(name, list(args))
    except Exception:
        pass


def _kill_tree(p):
    """Windows 는 `kill()` 로 **자식의 자식까지 안 죽인다** — 나무째 끊는다."""
    if os.name == "nt":
        try:
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(p.pid)],
                           capture_output=True, timeout=30,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception:
            pass
    try:
        p.kill()
    except Exception:
        pass


def _tidy(out):
    # 토큰 절약: 노이즈 제거 후 요약만 보존(상세는 각 모듈이 reports/에 파일로 남긴다)
    keep = [ln for ln in out.splitlines()
            if ln.strip() and not any(x in ln for x in
                ("UserWarning", "warn(msg)", "  [예정]", "  [건너뜀]", "i 관리대장 최신본"))]
    return "\n".join(keep[-12:]).strip()


def _run_once(name, args, timeout):
    """★ `subprocess.run(timeout=)` 을 쓰면 안 된다 — **윈도우에서 영원히 멈춘다**
       (2026-08-08 실사고).

    CPython 의 `subprocess.run` 은 시간이 넘으면 `process.kill()` 을 부른 뒤
    윈도우에서만 **시간제한 없는** `process.communicate()` 를 한 번 더 부른다.
    자식이 SMB(Z:) 읽기처럼 **끊기지 않는 대기**에 걸려 있으면 TerminateProcess 가
    먹지 않고, 그 두 번째 드레인이 끝나지 않는다. 30분 제한을 걸어 둔 단계가
    **13시간 30분**을 매달려 있었다(부모 CPU 0.4초 · 자식 0.5초 — 둘 다 그냥 서 있었다).

    그리고 그동안 회차는 **락을 쥔 채**라 다음 회차가 "이미 실행 중"으로 조용히
    건너뛴다. 스케줄러는 '성공'이라 적는다. 즉 **하루치 대조가 통째로 안 돌면서
    아무 데도 티가 안 난다** — 09:50 이 하는 일(접수취소·객관완료·청구상태)이 전부 멈춘다.

    그래서: ① 나무째 죽이고 ② 드레인에도 **제한을 건다** ③ 그래도 안 죽으면
    포기하고 **다음 단계로 넘어간다**. 안 죽은 자식은 리포트에 pid 로 남긴다 —
    회차 하나를 살리자고 회차 전체를 세우지는 않는다.
    """
    p = subprocess.Popen([PY] + args, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                         text=True, encoding="utf-8", errors="replace", cwd=ROOT, env=ENV,
                         creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    try:
        so, se = p.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        _kill_tree(p)
        try:
            p.communicate(timeout=60)
        except subprocess.TimeoutExpired:
            return {"name": name, "ok": False,
                    "out": f"시간초과({timeout}s) — 죽이려 해도 안 죽는다(pid {p.pid}). "
                           "Z: 대기에 걸린 것으로 보고 다음 단계로 넘어갔다"}
        return {"name": name, "ok": False, "out": f"시간초과({timeout}s)"}
    # 예외 이름과 원인은 traceback 끝에 있다. 앞 500자만 남기면 호출 위치만 보이고
    # 실제 TimeoutError/PermissionError가 잘려, 2026-08-01 큐 실패를 진단하지 못했다.
    out = (so or "") + (("\n[stderr] " + se[-2000:]) if p.returncode != 0 and se else "")
    return {"name": name, "ok": p.returncode == 0, "returncode": p.returncode,
            "out": _tidy(out)}


# 0단계 관문(합성검증)에 주는 시간. 실측 395.7초(2026-08-19, 한가한 기계) 대비
# 약 3.8배. `run()` 기본값 600초로는 바쁜 아침에 관문이 회차를 죽였다.
GATE_TIMEOUT_S = int(os.environ.get("COUPANG_GATE_TIMEOUT_S") or 1500)
GATE_PROOF_SCHEMA = 1
GATE_PROOF_NAME = "합성검증_통과증명.json"
GATE_SOURCE_EXTENSIONS = {
    ".py", ".js", ".html", ".css", ".ps1", ".bat", ".vbs", ".sql",
    ".toml", ".yaml", ".yml",
}
GATE_SOURCE_EXCLUDED_DIRS = {
    ".git", "__pycache__", "reports", "tmp", "outputs", "inbox",
    "archive_spool", ".pytest_cache", ".mypy_cache",
}


def _gate_proof_path(root=ROOT):
    return os.path.join(os.fspath(root), "reports", GATE_PROOF_NAME)


def _gate_source_files(root=ROOT):
    """검증 대상 코드 목록. 보고서·수집자료는 빼고 코드와 테스트만 센다.

    `git ls-files --others` 도 같이 보므로 아직 커밋하지 않은 새 기능도 빠지지 않는다.
    git 을 못 쓰는 설치본에서는 같은 확장자를 직접 훑되, 실데이터 폴더는 건드리지
    않는다. 파일 내용만 지문에 넣고 수정시각은 넣지 않는다 — 복사만 다시 했다고
    긴 검사를 재실행하지 않기 위해서다.
    """
    root = os.path.abspath(os.fspath(root))
    names = None
    try:
        # 공용 무창 실행기를 쓴다. `git.exe` 를 그냥 띄우면 담당자 화면에 검은 창이
        # 번쩍일 수 있고, `subprocess.run(timeout=)` 은 Windows 에서 종료 뒤에도
        # 매달릴 수 있다([175]·[272]).
        from proc_guard import run_tree
        got = run_tree(
            ["git", "-C", root, "ls-files", "--cached", "--others",
             "--exclude-standard", "-z"],
            timeout=15,
        )
        if got.returncode == 0:
            names = [x for x in got.stdout.split("\0") if x]
    except (OSError, subprocess.SubprocessError):
        names = None

    if names is None:
        names = []
        for base, dirs, files in os.walk(root):
            dirs[:] = [d for d in dirs if d not in GATE_SOURCE_EXCLUDED_DIRS]
            for name in files:
                names.append(os.path.relpath(os.path.join(base, name), root))

    out = []
    for rel in names:
        rel = rel.replace("\\", "/").lstrip("./")
        parts = [p for p in rel.split("/") if p]
        if not parts or any(p in GATE_SOURCE_EXCLUDED_DIRS for p in parts[:-1]):
            continue
        if os.path.splitext(parts[-1])[1].lower() not in GATE_SOURCE_EXTENSIONS:
            continue
        path = os.path.join(root, *parts)
        if os.path.isfile(path):
            out.append((rel, path))
    return sorted(set(out))


def _gate_fingerprint(root=ROOT, detail=False):
    """코드·테스트 내용의 결정적 SHA-256과 근거 수를 돌려준다.

    `detail=True` 면 `map: {상대경로: sha256}` 을 **더한다** — 검증 도중 코드가
    바뀌었을 때 **어느 파일이 바뀌었는지**를 대기 위해서다. 기존 반환 키
    (`fingerprint`·`files`·`bytes`)는 한 톨도 안 바뀐다([172]) — 합격증도 옛 검사도
    그대로다. 실측 0.37초(파일 296개·10.3MB)라 비싸지 않다.
    """
    digest = hashlib.sha256()
    per = {} if detail else None
    count = 0
    size = 0
    for rel, path in _gate_source_files(root):
        try:
            with open(path, "rb") as fh:
                body = fh.read()
        except OSError:
            # 목록을 만든 뒤 사라진 파일도 '검증 중 코드 변경'으로 잡히게 이름은 남긴다.
            body = b"<missing>"
        raw_name = rel.encode("utf-8", "surrogateescape")
        digest.update(len(raw_name).to_bytes(4, "big"))
        digest.update(raw_name)
        digest.update(len(body).to_bytes(8, "big"))
        digest.update(body)
        if per is not None:
            per[rel] = hashlib.sha256(body).hexdigest()
        count += 1
        size += len(body)
    out = {"fingerprint": digest.hexdigest(), "files": count, "bytes": size}
    if per is not None:
        out["map"] = per
    return out


def _gate_same(a, b):
    """두 지문이 **같은 판**인가. `map` 유무에 안 흔들리게 키를 명시한다.

    ⚠ `a != b` 로 dict 를 통째로 비교하면 한쪽에만 `map` 이 있을 때 **언제나 다르다** —
      그러면 관문이 매번 "검증 도중 코드가 바뀌었다"로 죽는다. 만들면서 그 자리를
      먼저 막았다. `fingerprint` 가 같으면 `bytes` 도 같다(같은 내용이다).
    """
    return (a.get("fingerprint") == b.get("fingerprint")
            and a.get("files") == b.get("files"))


def _gate_changed(before, after):
    """검증 도중 **바뀐 파일 이름**을 준다. 못 가르면 빈 목록이다([169] — 지어내지 않는다)."""
    a = before.get("map") or {}
    b = after.get("map") or {}
    if not a or not b:
        return []
    return sorted(k for k in set(a) | set(b) if a.get(k) != b.get(k))


def _load_gate_proof(root=ROOT):
    try:
        with open(_gate_proof_path(root), encoding="utf-8") as fh:
            proof = json.load(fh)
        return proof if isinstance(proof, dict) else {}
    except (OSError, ValueError):
        return {}


def _gate_proof_matches(proof, stamp):
    return bool(
        isinstance(proof, dict)
        and proof.get("schema") == GATE_PROOF_SCHEMA
        and proof.get("result") == "ALL GREEN"
        and proof.get("fingerprint") == stamp.get("fingerprint")
        and proof.get("files") == stamp.get("files")
    )


def _save_gate_proof(stamp, duration_s, root=ROOT):
    """실제 ALL GREEN 뒤에만 합격증을 원자적으로 저장한다."""
    path = _gate_proof_path(root)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    proof = {
        "schema": GATE_PROOF_SCHEMA,
        "result": "ALL GREEN",
        "fingerprint": stamp["fingerprint"],
        "files": stamp["files"],
        "bytes": stamp["bytes"],
        "verified_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "duration_s": round(float(duration_s), 3),
        "command": "python tests/synthetic_check.py",
    }
    tmp = "%s.tmp-%s-%s" % (path, os.getpid(), uuid.uuid4().hex)
    try:
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(proof, fh, ensure_ascii=False, indent=1)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp, path)
    finally:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
    return proof


def _run_gate(root=ROOT, runner=None):
    """코드가 같으면 최근 합격증을 쓰고, 바뀌었을 때만 전체 합성검증을 돈다.

    실패·시간초과는 옛 합격증을 덮어쓰지 않는다. 검증 도중 코드가 바뀌면 그 결과는
    어느 판을 검사한 것인지 알 수 없으므로 합격으로 인정하지 않는다.
    """
    runner = runner or run
    started = time.monotonic()
    before = _gate_fingerprint(root, detail=True)
    proof = _load_gate_proof(root)
    if _gate_proof_matches(proof, before):
        age = proof.get("verified_at") or "시각 없음"
        return {
            "name": "합성검증",
            "ok": True,
            "returncode": 0,
            "cached": True,
            "out": ("ALL GREEN — 검증서 재사용 · 코드 %d개 · 최근 전체검증 %s"
                    % (before["files"], age)),
        }

    step = runner(
        "합성검증", [os.path.join(os.fspath(root), "tests", "synthetic_check.py")],
        timeout=GATE_TIMEOUT_S, retry=0,
    )
    # ★ **실패 갈래도 같은 것을 묻는다** (2026-08-28 실사고).
    #   성공 갈래는 오래전부터 "검증 도중 코드가 바뀌면 어느 판을 검사한 것인지 알 수
    #   없다"고 묻는데 **실패 갈래는 안 물었다.** 그래서 반쯤 고친 코드를 시험한 실패가
    #   `code`(= 검증이 실제로 빨갛다)로 확언되고, 조치는 *"그 검증부터 본다"* 로
    #   나간다 — 사람을 **멀쩡한 검증으로** 보낸다([172]·[289]).
    #   실측 2026-08-28 09:50: 회차가 관문을 도는 동안 사람이 `app_server.py` 를
    #   고치는 중이었고, 그 반쪽을 시험한 `t49` 실패가 그대로 P0 가 됐다. 커밋 시각
    #   (10:37)이 실패(09:50)보다 **뒤**라는 것을 손으로 대 보고서야 갈렸다.
    # ⚠ 지문은 **한 번만** 잰다 — 성공·실패 어느 쪽으로 가든 같은 사실이다.
    after = _gate_fingerprint(root, detail=True)
    edited = not _gate_same(after, before)
    changed = _gate_changed(before, after) if edited else []

    if step.get("ok") and "ALL GREEN" in str(step.get("out") or ""):
        if edited:
            step = dict(step)
            step["ok"] = False
            step["소스바뀜"] = True
            step["바뀐파일"] = changed
            step["out"] = (str(step.get("out") or "")
                           + "\n검증 도중 코드가 바뀌었습니다 — 바뀐 판을 다시 검사합니다.")
            return step
        # ⚠ 합격증에는 `map` 을 안 담는다 — 파일만 커지고 읽는 쪽이 없다.
        _save_gate_proof({k: v for k, v in after.items() if k != "map"},
                         time.monotonic() - started, root)
        return step

    if edited:
        step = dict(step)
        step["소스바뀜"] = True
        step["바뀐파일"] = changed
    return step

def _run_pipeline():
    steps = []

    # 0. 합성검증 — 실패 시 전체 중단
    #
    # ★ **관문 자신이 회차를 죽이고 있었다** (2026-08-19 실사고). 09:50 회차가
    #   `시간초과(600s)` 로 죽었는데, 그 600 은 `run()` 의 **기본값**이었다 — 이 단계에
    #   맞춰 정한 값이 아니다. 실측(한가한 기계) **395.7초**이므로 여유가 1.5배뿐이고,
    #   09:50 은 하필 제일 바쁜 시각이다(09:35 원본정리 · 증분 파이프라인 5분마다 ·
    #   둘 다 Z:(SMB)를 훑는다). 그래서 **아무 코드도 안 고쳤는데** 바쁜 아침마다
    #   회차 전체가 첫 줄에서 죽는다. 스케줄러에는 `exit 1` 만 남는다.
    # ★ 검증은 계속 늘어난다(지금 [318]). 그러니 **한도를 실측에 맞춰 다시 잡는다** —
    #   실측이 이 값의 절반을 넘으면 그때는 늘릴 것이 아니라 **검증을 나눌 때**다.
    # ⚠ 회차 예산(`ROUND_BUDGET_MIN` 150분)보다 훨씬 작아야 한다. 관문 하나가 예산을
    #   다 먹으면 완주해도 남는 단계가 없다([180]).
    s = _run_gate()
    steps.append(s)
    if not s["ok"] or "ALL GREEN" not in s["out"]:
        # ★ **왜 막혔는지를 남긴다** — 이 단계는 자율복구 대기열에 안 들어가므로
        #   자국이 없으면 exit 1 만 남고 아무도 이유를 못 찾는다(2026-08-18 실사고).
        _leave_gate_trace(s)
        finish(steps, aborted=True)
        sys.exit("합성검증 실패 — 전체 중단")
    _clear_gate_trace()

    # 합성검증 뒤, 모든 대조보다 먼저 단일 투입함을 정본 폴더로 분류한다.
    # 예전에는 다운로드 흡수가 파이프라인 끝이라 새 자료가 다음 날까지 반영되지 않았다.
    # ★ 보고일 자동 갱신(2026-08-06 지시) — 맨 앞에 둔다. 뒤 단계들이 이 날짜를
    #   기준으로 브리핑·캡처를 만들기 때문이다. 이미 오늘 것이면 큐를 늘리지 않는다.
    steps.append(run("보고일·집계기준일 자동 갱신",
                     [os.path.join(ROOT, "report_dates.py")], timeout=300))

    steps.append(run("업로드 투입함 원본 분류",
                     [os.path.join(ROOT, "upload_intake.py"), "--apply"]))

    # 1. 판매·세금계산서 inbox 대조 (inbox 없으면 원장 준비표)
    steps.append(run("판매·세금계산서 대조", [os.path.join(ROOT, "ecount_reconcile.py")]))

    # 2. ERP 계정별원장 4유형 대조 (inbox에 '원장' 파일 있을 때만)
    # PC가 꺼진 동안 휴대폰에서 예약한 코드를 영구 큐에서 안전하게 가져온다.
    # 반영 성공 확인 전에는 서버 항목을 지우지 않으며 입력 보호시간에는 이 단계도 멈춘다.
    steps.append(run("휴대폰 클라우드 예약 반영", [os.path.join(ROOT, "cloud_queue_sync.py")]))

    if has_inbox_kind("ledger"):
        steps.append(run("ERP원장 4유형 대조", [os.path.join(ROOT, "erp_ledger_check.py")]))
    else:
        steps.append({"name": "ERP원장 4유형 대조", "ok": None,
                      "out": "스킵 — 내용 판별 가능한 계정별원장 파일 없음"})

    # 2.44 카톡·밴드 교차 — 같은 사건이 두 군데로 나뉘어 온다. 한쪽만 보면 반쪽만 본다.
    steps.append(run("카톡·밴드 교차 확인", [os.path.join(ROOT, "cross_signal.py")]))

    # 2.45 오류 사전·재발 감시 — 사람이 막힌 자리를 매일 센다(2026-08-13 지시).
    # 읽기 전용이다. 세 갈래로 가른다: **회귀**(막았다는데 또 났다) · ★새 오류 ·
    # 아는 오류. 실측 2026-08-13 — `/api/originals` 권한거부가 최근 3일 222건인데
    # 어느 화면에도 안 떴다. 오류가 기록되는 것과 누가 그것을 보는 것은 다른 말이다.
    steps.append(run("오류 사전·재발 감시", [os.path.join(ROOT, "error_book.py")], timeout=300))

    # 2.45 오기입 확인 — 손으로 적는 칸의 오타를 찾는다(읽기 전용, 아무것도 안 고친다)
    #   오기입은 매일 새로 생기고, **틀린 값은 비어 있지 않아서 어느 화면에도 안 띈다.**
    steps.append(run("오기입 확인", [os.path.join(ROOT, "typo_watch.py")]))

    # 2.44a 화면 ↔ 원본 사실대조 (2026-08-13 지시: "위 같은 문제 잡아내는 기능 AI 추가해").
    #   화면이 오류 없이 틀린 값을 보여 주는 것을 매일 기계가 먼저 묻는다.
    #   읽기 전용 — 아무것도 안 고치고 큐에도 안 넣는다. 결과는 인계 '먼저 처리할 것'.
    steps.append(run("화면 사실대조", [os.path.join(ROOT, "truth_watch.py")]))

    # 2.44b 대표보고 검증 (2026-08-13 지시: "대표 보고 화면 캡처 정말 중요함 /
    #   잘못된 데이터 들어가지 않게 에이전트 만들어서 검증해 매번").
    #   `truth_watch` 가 화면 전반에 대해 하는 일을 **대표보고 한 장**에 대해 한다.
    #   ★ 이 한 장은 값이 다르다 — 잘못된 원장 값은 고치면 되지만, **잘못된 대표보고는
    #     이미 사람의 판단을 바꾼 뒤다.** 실측 2026-08-13: '잔여 미수금액 0' 은
    #     아무도 안 채운 열에서 나온 확언이었다(750행 중 채워진 행 0개).
    #   읽기 전용 — 아무것도 안 고치고 큐에도 안 넣는다. 결과는 reports/대표보고_검증.md.
    steps.append(run("대표보고 검증", [os.path.join(ROOT, "exec_report_guard.py")]))

    # 2.45a 캠프명 ERP 표준화 (2026-08-11 지시: "ERP가 기준이야 … 기존 캠프명을 ERP
    #   기준으로 변경"). 새로 들어오는 접수·점검에도 비표준 이름이 계속 생기므로 회차다.
    #   유일 매칭만 승인 경로(앱 DB 감사로그 + Excel 보관 큐)로 바꾸고, 후보 여럿·
    #   ERP 없음·입력오류는 reports/캠프명_표준화.md 에 사람 몫으로 남는다.
    steps.append(run("캠프명 ERP 표준화",
                     [os.path.join(ROOT, "camp_standardize.py"), "--queue"], timeout=1500))

    # 2.45b 정기점검 점검내용 호기 분류·깨진 문자 조사 (같은 지시). 원문 불변 —
    #   파생 DB(db/pm_content.db)와 reports/정기점검_호기분류.md 만 만든다. 깨진
    #   내용(?? 호기)은 밴드 원본 근거가 설 때만 --queue 가 교정하므로 여기서는
    #   조사만 한다(근거 없는 교정은 사람 몫).
    steps.append(run("정기점검 호기 분류",
                     [os.path.join(ROOT, "pm_content.py")], timeout=1500))

    # 2.46 보험 건·보험사 입금 — 보험금은 **쿠팡이 아니라 보험사가** 낸다.
    #   입금을 찾는 receipt_fill 은 쿠팡 입금내역만 보므로 보험금은 어느 화면에도 안 떴다.
    #   받을 돈인데 안 받아도 티가 안 나는 자리다(2026-08-09 지시).
    steps.append(run("보험 입금 확인",
                     [os.path.join(ROOT, "insurance_watch.py"), "--queue"]))

    # 2.47 업무 단계 정의 — 관리대장 드롭다운·앱 흐름도가 바뀌었나 (2026-08-10 지시).
    #   흐름도가 바뀌었는데 **아무 화면에도 티가 안 나는 것**이 제일 위험하다.
    #   화면은 이 정의를 읽어 단계 선택지를 만들므로, 바뀌면 조용히 따라가 버린다.
    steps.append(run("업무 단계 정의 확인",
                     [os.path.join(ROOT, "work_flow.py"), "--check"]))

    # 2.5 쿠팡 PO 대조 (inbox에 'PO' 파일 있을 때만)
    if has_inbox_kind("po"):
        steps.append(run("쿠팡 PO 대조", [os.path.join(ROOT, "po_reconcile.py")]))
    else:
        steps.append({"name": "쿠팡 PO 대조", "ok": None, "out": "스킵 — inbox/에 쿠팡 PO 목록 파일 없음(파일명에 PO 포함)"})

    # 2.7 입금(수금) 자동입력 — 자료가 들어오면 사람 손 없이 채운다(사용자 지시 2026-07-28).
    #     파일이 없으면 receipt_fill 이 스스로 안내만 하고 조용히 끝나므로 조건 없이 돌린다.
    #     (계정별원장은 '0. 원본 자료' 에도 들어오므로 파일명 조건을 걸면 놓친다 — pick 이 내용으로 찾는다)
    steps.append(run("입금 대조·자동입력", [os.path.join(ROOT, "receipt_fill.py"), "--queue"]))

    # 2.8 카톡 신규 접수 등록 — 대화 내보내기가 들어오면 02·04 에 새 행으로 올린다.
    #     유형이 확정된 것(돌발·정기)만 올리고 철거·납품은 보류한다 — 대상 시트가 없다.
    kakao_sources = _source_files("kakao_dirs", (".txt",))
    if kakao_sources:
        #     ★ --days 7 로 **최근분만** 올린다. 과거분(현재 미등록 82건)은 파싱 품질을 사람이
        #       한 번 보고 넣어야 해서 자동에 태우지 않는다: python kakao_extract.py --new
        steps.append(run("카톡 신규 접수 등록",
                         [os.path.join(ROOT, "kakao_extract.py"), "--new", "--days", "7", "--queue"]))
    else:
        steps.append({"name": "카톡 신규 접수 등록", "ok": None, "out": "스킵 — kakao/inbox/에 대화 내보내기 없음"})

    # 2.81 유수비 대표 대화 판정 — 프로젝트번호 배정 전 쿠팡 접수도 대표보고·캡처에
    #      보이게 한다. 원장·입력 큐에는 쓰지 않고 쿠팡 건/무관/모름 수와 읽기 전용
    #      이벤트만 갱신한다. 내부 업무지시서와 사진·동영상은 접수로 세지 않는다.
    steps.append(run("대표 대화 쿠팡 접수 추출",
                     [os.path.join(ROOT, "ceo_events.py"), "--sync"]))

    # 2.9 작업 내용 자동 기입 — 원문(카톡·밴드 **완료 글**)에 있으면 03시트에 채운다.
    #     사용자 지시(2026-07-29): "밴드나 카톡에서 확인되면 자동 기입. 모든 데이터가 마찬가지."
    #     ★ '신청내용'(요청)은 쓰지 않는다 — 그건 무엇을 했나가 아니다. 빈 양식도 거른다.
    steps.append(run("작업내용 자동기입(03시트)", [os.path.join(ROOT, "fill_work_detail.py"), "--apply"]))

    # 2.95 구글 캘린더 대조 — 사용자 지시(2026-07-29): "이 캘린더 추가하고 항상 대조해서
    #      엑셀과 앱에 반영해줘". 캘린더는 **예정**이므로 예정일 칸만 채우고 완료 칸은 안 건드린다.
    #      원천(비공개 iCal 주소·.ics 파일)이 없으면 gcal_sync 가 스스로 조용히 끝난다 —
    #      주소 하나 없다고 일일 파이프라인 전체가 실패로 물들면 안 된다.
    steps.append(run("구글 캘린더 대조", [os.path.join(ROOT, "gcal_sync.py"), "--queue"]))

    # 2.955 청구(거래명세서) 근거 갱신 — ERP 판매조회의 새 매출을 06시트 빈 행에 올린다.
    #       ★ 같은 내용의 판매조회가 여러 벌 있으면 금액이 배수로 합산된다(2026-07-30 3배 사고).
    #         billing_fill 이 SHA256 으로 같은 파일을 한 번만 읽으므로 자동에 태워도 안전하다.
    steps.append(run("청구 근거 갱신(06시트)", [os.path.join(ROOT, "billing_fill.py"), "--queue"]))

    # 2.9555 청구상태(06시트 AH)를 ERP 수금확인으로 맞춘다 — 사용자 지시(2026-08-08)
    #        "ERP 기준으로 확정하고 객관적으로 입증되면 엑셀에 완료처리해".
    #        ★ 회차로 만드는 이유: 수금은 **매일 새로 생긴다.** 한 번 손으로 맞춰 두면
    #          그날 이후로는 다시 낡는데, 낡은 단계 딱지는 비어 있지 않아서
    #          **아무 화면에도 티가 안 난다** — 이 프로젝트가 계속 당해 온 종류다.
    steps.append(run("청구상태 ERP 수금확인 반영",
                     [os.path.join(ROOT, "billing_status.py"), "--queue"]))

    # 2.96 재계산 대기 세기 — 원장엔 올라왔는데 엑셀이 아직 계산 안 해 앱에 안 나오는 건.
    #      숫자가 틀린 게 아니라 대기 중이라는 걸 앱이 스스로 말하게 한다(사용자 오해 방지).
    steps.append(run("재계산 대기 확인", [os.path.join(ROOT, "recalc_pending.py")]))

    # 2.98 엑셀 재계산은 11:00·15:00 일괄반영 회차 안에서만 수행한다.
    #      09:50에는 승인·대기 상태만 읽어 보고, Excel COM을 열거나 저장하지 않는다.
    steps.append(run("엑셀 재계산 상태 확인", [os.path.join(ROOT, "excel_recalc.py")]))

    # 2.97 ERP 접속 IP 확인 — 공인 IP가 바뀌면 이카운트 OAPI가 통째로 막힌다.
    #      사용자 지시(2026-07-30): "IP가 변경되면 이 화면에서 등록해서 진행".
    #      자동 등록은 하지 않는다(회사 ERP 보안 설정) — 바뀐 사실과 넣을 값을 알린다.
    steps.append(run("ERP 접속 IP 확인", [os.path.join(ROOT, "erp_ip_guard.py")]))

    # 이카운트가 **조회 API로 공식 제공하는 것만** 가져온다. 품목·발주서는 API,
    # 판매/매입전표·계산서·수금은 기존 XLSX 도착 확인 경로다. 6시간 성공 캐시가 있어
    # 일일 회차가 같은 7천여 품목을 매번 다시 요청하지 않는다.
    steps.append(run("ERP 공식 API 자료 수집",
                     [os.path.join(ROOT, "erp_api_collect.py")], timeout=300))

    # 3. 밴드 수집·대조 — 공식 API 토큰이 있으면 수집+대조, 브라우저 수집 캐시만 있으면 대조만
    band_dumps = _source_files("band_dump_dirs", (".json",))
    if band_dumps:
        steps.append(run("밴드 업로드 원본 변환", [os.path.join(ROOT, "band", "ingest.py")]))

    band_cache = [f for f in glob.glob(os.path.join(ROOT, "band", "cache", "*.json"))
                  if not os.path.basename(f).startswith(("raw_", "dump_"))]
    if os.path.exists(os.path.join(ROOT, "band", ".band_token.json")):
        steps.append(run("밴드 수집", [os.path.join(ROOT, "band", "band_sync.py")]))
        steps.append(run("밴드 대조", [os.path.join(ROOT, "band", "band_reconcile.py")]))
    elif band_cache:
        steps.append(run("밴드 대조(캐시)", [os.path.join(ROOT, "band", "band_reconcile.py")]))
    else:
        steps.append({"name": "밴드 수집·대조", "ok": None, "out": "스킵 — 밴드 미인증·캐시 없음(앱 심사 대기)"})

    # 4. 카톡 대조 (kakao/inbox에 txt 있을 때만)
    if kakao_sources:
        steps.append(run("카톡 대조", [os.path.join(ROOT, "kakao", "kakao_reconcile.py")]))
    else:
        steps.append({"name": "카톡 대조", "ok": None, "out": "스킵 — kakao/inbox/에 대화 내보내기 txt 없음"})
    # 카톡 내보내기가 1일 넘게 오래되면 대시보드·아침 브리핑에 경고(2026-08-04 지시).
    steps.append(run("카톡 내보내기 경과 감시", [os.path.join(ROOT, "kakao", "export_watch.py")]))

    # 5. 확정 업데이트 자동 반영 — 빈 칸만·근거 보유·항상 새 버전(vN+1) 생성이라 안전
    # ERP 매출서류(계산서·명세서 현황) — 있으면 대조 + 25시트 반영
    _has_tax = has_inbox_kind("tax")
    if _has_tax:
        steps.append(run("ERP 매출서류 대조(반영 대기)", [os.path.join(ROOT, "erp_docs_check.py")]))
    else:
        steps.append({"name": "ERP 매출서류 대조", "ok": None,
                      "out": "스킵 — inbox/에 매출(세금)계산서현황 없음"})

    # 밴드 문서 사진(거래명세서·세금계산서) OCR → 확실한 건은 빈칸 입력 큐에 적재
    document_images = _source_files("doc_photo_dirs",
                                    (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".heic"))
    if document_images:
        # 사진 1장에 OCR 1초 남짓. 첫 회는 1,458장에 25분이 걸려 600초 제한에 걸렸다(2026-07-26).
        # doc_ocr가 결과를 band/ocr_cache/에 저장하므로 두 번째 실행부터는 몇 초로 끝난다.
        steps.append(run("밴드 문서 이미지 대조·입력", [os.path.join(ROOT, "band", "doc_ocr.py"),
                                                       "--scan", "--apply"], timeout=1800))
        # 그 위에 교차검증을 한 겹 더 얹는다(2026-08-06 지시). 엔진 하나의 답만 믿으면
        # 틀린 값이 조용히 원장에 들어간다 — 서로 다른 엔진 둘이 **같은 값**을 낸 항목만
        # 빈칸에 넣고, 갈리면 사람에게 넘긴다. 전량을 두 번 읽지는 않는다:
        # 금액 정합성이 깨졌거나 핵심 항목을 못 읽은 건, 그리고 원장에 쓰려는 건만 재검한다.
        steps.append(run("문서 OCR 교차검증", [os.path.join(ROOT, "band", "ocr_crosscheck.py"),
                                               "--scan", "--apply"], timeout=1800))
    else:
        steps.append({"name": "밴드 문서 이미지 대조", "ok": None,
                      "out": "스킵 — band/docs_inbox/에 사진 없음"})
        steps.append({"name": "문서 OCR 교차검증", "ok": None, "out": "스킵 — 사진 없음"})

    # 미수집 원본·사진·텍스트를 **매일 조금씩** 굳힌다 (2026-08-08 지시:
    # "미수집 데이터들 싹 다 긁어모아 … 두번 일 안하게"). 한 번에 다 하려 들면
    # 회차가 시간 제한에 걸려 통째로 실패한다 — 상한을 두고 매일 이어 간다.
    # 이미 있는 것은 파일을 열지도 않으므로 다 끝난 뒤에는 몇 초로 끝난다.
    # 로그인이 필요한 수집(밴드·ERP)은 여기에 없다 — 그건 사람 몫이다(절대규칙 3).
    # ★ 시간제한을 3600초(1시간)에서 1200초로 줄였다 (2026-08-09 실측).
    #   이 한 단계가 **회차 전체를 먹고 있었다** — 09:50 회차가 2시간째일 때 계속
    #   이 단계에 있었고(한 단계에 65분+), 그 뒤 단계들은 시작조차 못 했다.
    #   그게 '32시간째 미완주'의 실제 모습이다.
    #   ※ 상한을 줄여도 **잃는 것이 없다.** 이 단계는 원래 "매일 조금씩 굳히는" 것이라
    #     이번에 못 한 몫은 내일 이어서 한다. 반면 회차가 안 끝나면 그 뒤의 대조·큐가
    #     **전부** 멈춘다. 하나를 다 하려다 나머지를 다 잃는 거래는 하지 않는다.
    steps.append(run("미수집 원본·사진·텍스트 보관", [os.path.join(ROOT, "collect_all.py"),
                                                     "--run", "--limit", "600"], timeout=1200))

    # 완료보고서와 문서발행 표시는 프로젝트NO가 정확히 일치하는 근거만 빈칸에 큐잉한다.
    # 실제 ZIP 패치는 바로 아래 ledger_writer 한 번으로 합쳐 버전 난립과 충돌을 막는다.
    steps.append(run("완료보고서 확정분 큐", [os.path.join(ROOT, "confirm_fill.py"), "--queue"]))
    # 밴드·카톡 완료보고 + ERP 판매조회 + 06 거래명세서 원장을 프로젝트NO로 함께 대조한다.
    # 02·03·04의 원인 열만 채우므로 검증결과 수식은 보존되고 Excel에서 정상/확인으로 재계산된다.
    steps.append(run("ERP·거래명세서·현장검증 확정분 큐",
                     [os.path.join(ROOT, "verification_sync.py"), "--queue"]))
    # 정기점검·돌발AS 일지(미실시건) — 완료근거는 원장 빈 칸만 보완 큐에 넣고,
    # 미실시 사유·취소·일자상이는 28_대조현황과 대표보고에 그대로 남긴다.
    steps.append(run("정기점검·돌발AS 일지 대조 안전입력 큐",
                     [os.path.join(ROOT, "work_log_sync.py"), "--queue"]))
    # 완료일·검증 정상 등 객관 요건을 충족한 행을 Excel 상태 셀 대신 DB 완료 정본에 기록한다.
    # confirm_fill의 밴드 완료글+날짜 판정과 함께 매일 돌아야 사람이 다시 지시하지 않는다.

    # ★ 2026-07-30 지시: 반영은 **DB에 모았다가 하루 두 번(11:00·15:00)만** 엑셀에 쓴다.
    #   전에는 여기서 바로 --apply 해서 하루에 관리대장 버전이 수십 개씩 늘었다(v311→v327).
    #   09:50 일일 대조는 적재까지만 하고, 별도 작업 스케줄러가 두 회차를 정확히 실행한다.
    # ★ ERP 등록 여부(02_돌발AS접수 ERP등록 · 04_정기점검 ERP판매전표)를 채운다.
    #   2026-08-07 에 알았다: 이 도구는 2026-07-28 에 만들어져 있었는데 **daily_run 에
    #   들어 있지 않았다.** 손으로 돌린 사람이 있을 때만 채워졌다는 뜻이다. 이날 그냥
    #   돌려 보니 채울 칸이 52개 남아 있었고, 그중 46건은 ERP 가 이미 완료를 입증한
    #   건이었다. 도구가 있는데 아무도 안 부르는 것이 가장 조용한 종류의 누락이다.
    #   판정 근거는 ERP **판매조회의 프로젝트코드(UJ번호)** 다 — 거래명세서번호가
    #   아니다(그 길은 2026-07-28 에 이미 막다른 길로 판명났다. 365건이 그렇게 남았었다).
    #   --queue 라 엑셀은 열지 않는다. 반영은 11:00·15:00 회차 그대로.
    #   (retry 는 적지 않는다 — `--queue` 라 NO_RETRY_FLAGS 가 알아서 막는다)
    steps.append(run("ERP 등록여부 판정(판매조회)",
                     [os.path.join(ROOT, "fill_erp_status.py"), "--queue"]))
    steps.append(run("입력 DB 적재", [os.path.join(ROOT, "ledger_db.py"), "--intake"]))
    # 방금 흡수한 신규행에도 실제 완료일+원천 근거가 있으면 Excel 회차를 기다리지 않고
    # SQLite 완료 정본에 즉시 기록한다. Excel 자체는 정해진 11:00·15:00 회차만 유지한다.
    steps.append(run("객관근거 완료 DB 동기화",
                     [os.path.join(ROOT, "complete_verified.py"), "--queue"]))
    # 금액은 PO 원본 견적서와 거래명세서가 프로젝트·PO·총액까지 유일 일치할 때만,
    # 계산서는 25/26시트 ERP 원본의 '확정' 프로젝트만 정산 완료 DB에 올린다.
    steps.append(run("정산 객관완료 DB 동기화",
                     [os.path.join(ROOT, "settlement_completion.py"), "--sync"]))
    # 세 사람의 담당 범위도 같은 객관 근거로 완료 이력을 남긴다. 상태 문구는
    # "류지영 완료"·"오종현 완료"·"유현민 완료"이며 사람 체크나 추정은 쓰지 않는다.
    steps.append(run("담당자별 객관완료 DB 동기화",
                     [os.path.join(ROOT, "staff_completion.py"), "--sync"]))
    # 오래된 세금계산서 미발행을 경과일 순으로 잡아낸다(2026-08-03 지시). 읽기 전용 감시 —
    # 발행일은 지어내지 않고, 발행 사실은 ERP 원본이 들어와야 객관완료가 잡는다.
    # ERP 거래명세서 **인쇄본**(ESD007E) 파싱·원장 대조. 현황표(erp_docs_check)와 레이아웃이
    # 달라 그 도구는 0건으로 봤다(2026-08-04 실측) — 별도 파서가 읽는다. 읽기 전용.
    steps.append(run("거래명세서 인쇄본 대조", [os.path.join(ROOT, "stmt_docs.py")],
                     timeout=1500))
    # ERP 판매조회 → 프로젝트 색인. **완료 판정·앱 금액의 정본**이라 명세서 색인보다 먼저.
    #   (2026-08-05: 이 색인이 없으면 settlement_completion 의 ERP 근거가 통째로 죽는다)
    steps.append(run("ERP 판매 프로젝트 색인", [os.path.join(ROOT, "erp_sales_index.py")],
                     timeout=900))
    # 명세서 ↔ 판매조회(UJ 프로젝트코드) 색인. 인쇄본에는 UJ 가 없어 금액·거래처·품목으로
    # 잇는다(2026-08-05). 이 색인이 건별 PDF 파일명의 프로젝트NO 가 된다.
    steps.append(run("명세서 ↔ 프로젝트 색인", [os.path.join(ROOT, "stmt_link.py")],
                     timeout=1500))
    # 거래처코드·캠프 마스터 — 앱이 CU코드를 표시하고, 캠프 공백(ERP에만/원장에만)을 드러낸다.
    steps.append(run("거래처코드 색인", [os.path.join(ROOT, "customer_index.py")], timeout=900))
    steps.append(run("캠프 마스터", [os.path.join(ROOT, "camp_master.py")], timeout=1200))
    # 캠프 **담당자**(현장책임·안전관리 이름·전화) — 유수비 대표 지시(2026-08-18).
    #   ERP 에는 캠프 전화가 거의 없다(CU 356개 중 18개). 원천은 밴드 접수 글이다.
    #   ★ 캠프 마스터 **뒤**에 온다 — 거래처코드를 그 결과에서 빌린다.
    steps.append(run("캠프 담당자", [os.path.join(ROOT, "camp_contacts.py"), "--write"],
                     timeout=900))
    # 캠프명 <-> ERP 거래처코드 대조(유니웍스 기초등록 · 2026-08-18 노승용 매니저 요청).
    # **캠프 담당자 뒤**에 온다 — 그 결과를 원천 하나로 읽는다. 읽기 전용이라
    # 어디에도 안 쓴다(찾아 주기만 한다, `typo_watch` 와 같은 자리).
    steps.append(run("캠프 거래처코드 대조", [os.path.join(ROOT, "camp_code_match.py")],
                     timeout=300))
    # ★ 사용자 지시(2026-08-05) "각 건의 PDF·이미지를 번호로 알아보게 저장":
    #   밴드는 글 단위(PDF+텍스트+사진), ERP 명세서는 전표번호 단위 PDF 로 굳힌다.
    #   회차마다 상한을 둬 daily_run 이 길어지지 않게 한다 — 남은 건 다음 회차가 잇는다.
    # 밴드 게시글 보관은 위의 `미수집 원본·사진·텍스트 보관`(`collect_all.py`)이
    # 이미 같은 `archive_posts.py`를 **7분 예산·체크포인트**로 실행한다. 여기서 다시
    # 150건을 돌리면 같은 회차가 최대 30분 더 멈췄다(2026-08-24 실측 1,804초).
    # 보관 기능을 없애는 것이 아니라, 시간 제한과 이어하기가 있는 한 경로로 합친다.
    steps.append(run("명세서 건별 PDF 보관",
                     [os.path.join(ROOT, "stmt_archive.py"), "--limit", "40"],
                     timeout=300))
    steps.append(run("세금계산서 건별 PDF 보관",
                     [os.path.join(ROOT, "tax_archive.py"), "--limit", "40"],
                     timeout=300))
    # ★ 원본이 늘면 분석 3종(미발행·불일치·확인필요현황)이 10분을 넘긴다 —
    #   2026-08-04 거래명세서 785건 흡수 후 기본 600초에 셋 다 타임아웃으로 FAIL했다.
    #   단독 재실행은 전부 성공(로직 문제 아님) → 시간만 넉넉히 준다.
    steps.append(run("세금계산서 미발행 경과 감시",
                     [os.path.join(ROOT, "tax_invoice_watch.py")], timeout=1500))
    # 금액 재계산 대기의 견적↔명세 교차·불일치 진단(2026-08-03 지시). 읽기 전용 —
    # 명세합계가 다른 프로젝트 견적과 일치하는 입력 밀림을 짝까지 찾아 보여 준다.
    steps.append(run("견적·명세 불일치 진단",
                     [os.path.join(ROOT, "quote_mismatch.py")], timeout=1500))
    # 09:50에는 읽기 전용 무결성 검사만 한다. 실제 복구도 11:00·15:00 회차 안에서만 한다.
    steps.append(run("워크북 무결성 검사", [os.path.join(ROOT, "fix_workbook.py")]))

    # 5.3 류지영 정기점검 스케줄 원본 → 27_정기점검원본일정.
    #     지정 폴더의 최신 xlsx를 매일 다시 읽으며, 내용이 같으면 새 버전을 만들지 않는다.
    #     원본에 UJ번호가 없으므로 임의 프로젝트를 만들지 않고 04시트와 근거가 있는 건만 연결한다.
    steps.append(run("정기점검 스케줄 원본 분석",
                     [os.path.join(ROOT, "pm_schedule_sync.py")]))

    # 신규 프로젝트 업무 흐름도는 관리대장·앱에 노출하지 않는다. 지정 원본 폴더의 최신본만
    # 내부 DB로 안전하게 교체해, 이후 신규 업무 기준을 추가해도 원본 구조를 그대로 보관한다.
    steps.append(run("신규 프로젝트 업무 흐름도 DB 동기화(앱 비표시)",
                     [os.path.join(ROOT, "new_project_flow_sync.py"), "--apply"]))

    # 5.5 밴드 업무 추출 보고. 24시트 반영은 11:00·15:00 회차에서만 한다.
    if band_cache:
        steps.append(run("밴드 업무추출(반영 대기)", [os.path.join(ROOT, "band_extract.py")]))
        # 5.6 접수했다가 취소된 건(2026-08-08 지시). 취소된 건이 원장에 접수 그대로
        #     남으면 'AS 미실시'가 영영 안 줄어든다 — 화면은 멀쩡한데 숫자가 거짓인
        #     조용한 사고다. 앱 DB 정본에 즉시 기록하고 Excel은 자동 보관본만 따라간다.
        steps.append(run("접수취소·원격해결 DB 동기화",
                         [os.path.join(ROOT, "cancel_watch.py"), "--sync"], timeout=1500))
        # 5.7 위 단계가 **반쪽으로 도는 것**을 막는다 (2026-08-09). 취소는 대부분 댓글로
        #     오는데 실측 7,475글이 댓글을 한 번도 안 들여다봤다. 그 상태에서도
        #     cancel_watch 는 오류 없이 끝난다 — 사각지대는 오류가 아니라 **없는 자료**다.
        #     여기서는 **긁지 않는다**(사람 로그인이 필요하다 — 절대규칙 3).
        #     사람 손에 갈 붙여넣기 파일만 최근 것부터 다시 만들어 둔다.
        steps.append(run("댓글 사각지대 붙여넣기 파일 갱신",
                         [os.path.join(ROOT, "band", "comment_backfill.py"),
                          "--days", "90", "--write"], timeout=900))

    # 6. 확인필요현황은 별도 보고서만 갱신한다. 관리대장 23시트 쓰기는 11:00·15:00 회차.

    # 6.5 확인필요현황 집계 갱신.
    #     ★ 2026-08-07 지시로 **별도 엑셀을 더는 만들지 않는다** ("앞으로도 별도의
    #       엑셀 파일은 만들지 말고 관리대장으로만 관리해"). 예전 주석은 *"23시트는
    #       평면 목록이라 유형별 상세열을 담지 못한다"* 였는데, 담지 못한 게 아니라
    #       안 나눴을 뿐이었다 — 이제 프로젝트NO·명세서번호·PO번호·확인방법이
    #       시트에 **열로** 있다. 시트 반영은 11:00·15:00 회차(`ledger_db`)가 한다.
    #     이 단계가 여전히 매일 도는 이유는 **집계 JSON** 때문이다
    #     (reports/확인필요_집계.json — 앱·리포트가 이 숫자를 읽는다).
    #     ★ 여기에 연결돼 있지 않아 2026-07-27 16:45 자로 하루 동안 멈춰 있었다 —
    #       읽는 사람은 멈춘 줄 모르고 어제 숫자를 오늘 숫자로 본다. 반드시 매일 돈다.
    steps.append(run("확인필요현황 집계 갱신", [os.path.join(ROOT, "findings_export.py")],
                     timeout=1500))

    # 6.6 Z: 상시 공백 스캔 — 파일명 기반 전수조사와 서류↔원장 1:1 대조는 읽기 전용이다.
    #     2026-07-28 뒤 리포트가 멈춰 자료현황이 오래된 숫자를 계속 보여 준 문제를 막는다.
    steps.append(run("Z폴더 원장 누락·금액 공백 스캔", [os.path.join(ROOT, "zscan.py")], timeout=1800))
    steps.append(run("Z폴더 서류 1:1 대조", [os.path.join(ROOT, "zscan.py"), "--docs"], timeout=1800))

    # 6.7 자료현황 한 장 — "밴드에서 뭘 얼마나 가져왔나 / 지금 뭘 갖고 있나 / 원장이 얼마나 찼나".
    #     같은 질문을 매번 다시 세지 않으려고 만든다(사용자 지시 2026-07-29).
    #     느린 것(Z: 2만 개 순회)은 다시 돌지 않고 앞 단계가 남긴 리포트에서 숫자만 읽는다 —
    #     그래서 이 단계는 위 대조들이 **끝난 뒤에** 와야 한다.
    # ★ 원본 색인·폴더 정리(2026-08-05 지시 "누가 봐도 깔끔하게 항상 정리").
    #   색인이 있어야 앱 '원본 자료' 화면과 바로가기가 최신을 가리킨다.
    # 업무센터 UX 점검은 3일마다 12:00 별도 스케줄러가 돈다(쿠팡업무_업무센터UX점검).
    #   daily_run 에서도 가볍게 한 번 더 갱신해 리포트가 오래되지 않게 한다.
    steps.append(run("업무센터 UX 점검", [os.path.join(ROOT, "ux_review.py")], timeout=600))
    steps.append(run("원본 색인 갱신", [os.path.join(ROOT, "source_index.py")], timeout=1800))
    steps.append(run("원본 폴더 정리·바로가기", [os.path.join(ROOT, "source_tidy.py")], timeout=1800))
    # 캐시가 어긋나면 Z: 원본 전체를 다시 훑는다 — 실측 콜드 222초, 경합 아래서는 600초를
    # 넘겨 매 회차 FAIL 이었다(2026-08-07). `_보관` 제외로 줄였지만 여유를 둔다.
    steps.append(run("자료현황 갱신", [os.path.join(ROOT, "data_status.py")], timeout=1200))
    # 미반영 목록(2026-08-05 지시 "반영 안된 자료들 목록 정리"). 앞 단계들이 남긴 리포트를
    # 모아 세므로 **맨 뒤**여야 한다 — 먼저 돌면 어제 숫자를 오늘 것으로 보여 준다.
    steps.append(run("미반영 목록 갱신", [os.path.join(ROOT, "pending_report.py")], timeout=300))
    # 6.75 하루치 정산분 보고자료 — 다음 날 아침 대표 보고에 그대로 쓴다(2026-08-05 지시).
    #      카톡·밴드·ERP 를 각각 그 날짜로만 훑어 만든다. 앞 단계(색인·밴드 반영)가
    #      끝난 뒤여야 ERP 내보내기와 밴드 글이 최신으로 잡힌다.
    steps.append(run("정산분 보고자료(어제)", [os.path.join(ROOT, "settle_report.py")],
                     timeout=900))

    # 6.85 다운로드 흡수 — 실행 도중 새로 내려받은 파일을 다음 회차용 투입함에 보존한다.
    #      떨어진다. 손으로 옮기면 세션이 바뀌는 순간 잊힌다(2026-07-31 실제로 그랬다).
    #      내용 판별로 '0. 원본 자료'에 이동 — Z: 라서 옮겨지는 순간 모든 PC·세션이 본다.
    steps.append(run("다운로드 흡수(원본 자료로)", [os.path.join(ROOT, "download_intake.py"), "--apply"]))

    # 6.87 ERP 엑셀 → PDF 사본 (사용자 지시 2026-08-04 "PDF 또는 이미지로 저장해 전부 반영").
    #      엑셀은 열 때마다 서식이 흔들리고 재조회 시점에 숫자가 달라진다. PDF 는 "그때
    #      ERP 가 이렇게 보여 줬다"를 고정한다 — 사후 대조·감사에 쓰는 건 이 고정본이다.
    #      새 파일만 변환하므로 매 회차 비용이 거의 없다(대상 0개면 Excel 을 띄우지도 않는다).
    # ★ 예산이 설계값과 어긋나 있었다 (2026-08-07 실측). 이 도구는 안에서 Excel 서브프로세스에
    #   1800초를 주는데(erp_pdf_export.py:116) 여기서는 600초만 줬다. 그래서 변환할 게
    #   4개뿐인 날에도 `시간초과(600s)` 로 매 회차 FAIL 이었다(단독 실행은 315초 exit 0).
    steps.append(run("ERP PDF 사본 만들기", [os.path.join(ROOT, "erp_pdf_export.py")],
                     timeout=1800))

    # 6.9 폰 원격 준비 상태 — 막히는 건 조용히 막힌다. 절전 설정이 되살아나거나 미푸시가
    #     쌓이면 정작 폰에서 붙으려 할 때 알게 된다(사용자 지시 2026-07-31). 아무것도 바꾸지 않고
    #     상태만 남긴다 — 전원·SSH 는 관리자 권한이 필요한 시스템 설정이라 사람이 실행한다.
    steps.append(run("폰 원격 준비 점검", [os.path.join(ROOT, "remote_ready.py")]))

    # 7. 전표 전송 대기 현황 (dry-run만 — 실전송은 절대 자동화하지 않음)
    steps.append(run("전표 전송대기(dry-run)", [os.path.join(ROOT, "ecount_upload.py")]))

    # 8. 클라우드 사본 갱신 — PC가 꺼져도 폰에서 오늘 자료를 볼 수 있게.
    #    config/cloud.json 이 없으면 조용히 건너뛴다(설정 전에는 아무 일도 안 함).
    if os.path.exists(os.path.join(ROOT, "config", "cloud.json")):
        steps.append(run("클라우드 사본 올리기", [os.path.join(ROOT, "cloud_export.py"), "--upload"]))
    else:
        steps.append({"name": "클라우드 사본 올리기", "ok": None,
                      "out": "스킵 — config/cloud.json 미설정(CLOUD_SETUP.md 참고)"})

    # 9. 폰용 사본 — PC가 꺼져 있어도 열리는 HTML 한 장(서버·인터넷 불필요).
    #    이동이 잦아 PC를 켜 둘 수 없을 때 이게 유일하게 확실한 방법이다.
    steps.append(run("폰용 사본 만들기", [os.path.join(ROOT, "mobile_snapshot.py")]))

    # 9-b. 대표 보고용 내용 브리핑 — 대표 지시(2026-07-28): 숫자가 아니라 '무슨 일이 있었나'
    steps.append(run("대표 브리핑(내용)", [os.path.join(ROOT, "daily_brief.py"), "--md"]))

    # 10. 고정 주소 사본 — PC를 꺼도 폰·태블릿이 이걸로 조회·자동채움을 한다(잠가서 올린다)
    # 공유 캘린더 꾸러미를 **먼저** 새로 잠근다 — 그래야 바로 아래 게시가 최신을 올린다.
    # (단톡방에 뿌린 링크는 이 파일 하나만 본다. 2026-08-06 지시)
    steps.append(run("공유 캘린더 갱신", [os.path.join(ROOT, "cal_share.py")], timeout=600))
    steps.append(run("고정 주소 사본 올리기", [os.path.join(ROOT, "cloud_publish.py"), "--push"]))

    # 11. 버전 파일 정리 — 최신본 하나만 작업 폴더에 두고 구버전은 사용자가 지정한
    #     OLD/ 한 곳으로 옮긴다. 같은 이름이 있어도 덮어쓰거나 삭제하지 않는다.
    # 9.9 복구용 보관 — 코드는 git bundle 한 파일로, 기록은 사실만. 서버(Z:)에 둔다.
    #     PC가 죽으면 PC 안의 백업은 같이 죽는다. 비밀키는 절대 담지 않는다(규칙 1).
    # ★ 재시도하지 않는다 (2026-08-07 실측). 이 단계는 Z: 로 1,721개·199.6MB 를 한 개씩
    #   복사하며 완주에 ~1,475초가 든다 — 산발적 경합이 아니라 **결정적으로** 한도를 넘는다.
    #   그런 단계를 재시도하면 실패에 시간을 두 배로 태울 뿐이다(세 단계 합쳐 회차당 ~81분).
    #   근본 해결(증분 복사·야간 이동)은 별도 과제로 남긴다 — reports/오류점검_20260807.md B1.
    steps.append(run("복구용 보관(서버)", [os.path.join(ROOT, "archive_keep.py")],
                     timeout=1800, retry=0))

    steps.append(run("관리대장 버전 정리", [os.path.join(ROOT, "ledger_versions.py"), "--prune"]))

    finish(steps)
    return steps


def main():
    if is_input_window():
        print(f"입력 보호시간({input_window_label()}) — 일일 자동대조를 시작하지 않습니다.")
        return
    token = acquire_run_lock()
    if not token:
        # ★ 겹쳐서 물러난 것을 **자국으로 남긴다** (2026-08-17 지시). 예전에는 한 줄만
        #   찍고 exit 0 이라, 밖에서 보면 **겹쳐서 안 돈 것과 다 한 것이 구별되지
        #   않았다**(`[169]`). 스케줄러는 '성공'이라 적는다.
        #   양보는 "저쪽이 그 일을 한다"는 **주장**이므로 주인 이름을 같이 남긴다 —
        #   `coordinate.audit()` 이 그 주인이 정말 끝냈는지 되묻는다.
        print("다른 daily_run 프로세스가 이미 실행 중 — 중복 실행을 시작하지 않습니다.")
        try:
            import coordinate
            coordinate.record_yield(COORD_JOB, COORD_JOB,
                                    "다른 daily_run 이 락을 쥐고 있다")
        except Exception:                 # noqa: BLE001 — 조율을 적으려다 회차를 막지 않는다
            pass
        return
    # ★ 진행 파일을 **덮기 전에** 앞 회차의 마지막 자리를 읽는다(분담판 [140]).
    #   순서가 뒤집히면 그 증거는 영영 사라진다 — 덮는 것은 바로 다음 줄이다.
    _note_prev_crash()
    _ROUND_T0[0] = datetime.now()
    _OVER_BUDGET[0] = False
    note_progress("(회차 시작)", "시작", {"끝난단계": []})
    final_state = "중단"
    final_extra = {"종료구분": "중단"}
    try:
        result = _run_pipeline()
        failed = [str(step.get("name") or "") for step in (result or [])
                  if isinstance(step, dict) and step.get("ok") is False]
        if failed:
            final_state = "실패"
            final_extra = {
                "종료구분": "실패",
                "끝까지실행": True,
                "실패단계": failed[-20:],
            }
        elif _OVER_BUDGET[0]:
            final_state = "완주(예산초과로 일부 건너뜀)"
            final_extra = {"종료구분": "완주", "일부건너뜀": True}
        else:
            final_state = "완주"
            final_extra = {"종료구분": "완주"}
        return result
    except KeyboardInterrupt:
        final_state = "중단"
        final_extra = {"종료구분": "중단", "오류유형": "KeyboardInterrupt"}
        raise
    except SystemExit as exc:
        if exc.code in (None, 0):
            final_state = "완주"
            final_extra = {"종료구분": "완주"}
        else:
            final_state = "실패"
            final_extra = {"종료구분": "실패", "오류유형": "SystemExit",
                           "오류": str(exc.code)[:300]}
        raise
    except BaseException as exc:
        final_state = "실패"
        final_extra = {"종료구분": "실패", "오류유형": type(exc).__name__,
                       "오류": str(exc)[:300]}
        raise
    finally:
        # ★ 정상 완주·단계 실패·예외 실패·사용자 중단을 서로 다른 사실로 남긴다.
        # 이 한 줄이 없으면 회차가
        #   중간에 죽었을 때 진행 기록이 '시작' 인 채로 굳어, 다음 회차가
        #   "아직 돌고 있나"와 "죽었나"를 구별하지 못한다.
        note_progress("(회차 끝)", final_state, final_extra)
        # 양보한 쪽이 "저쪽이 한다"고 적었다 — 그 주장이 지켜졌다는 증거를 여기서 남긴다.
        try:
            import coordinate
            coordinate.record_run(COORD_JOB, final_state,
                                  str(final_extra.get("오류") or ""))
        except Exception:                 # noqa: BLE001
            pass
        release_run_lock(token)


GATE_CRASH = os.path.join(REPORT_DIR, "일일대조_오류.json")
STEP_CRASH = os.path.join(REPORT_DIR, "일일대조_단계중단_오류.json")


def _gate_headline(out):
    """검증 출력에서 **사람이 읽을 한 줄**을 뽑는다 — 못 뽑으면 지어내지 않는다."""
    lines = [x.strip() for x in str(out or "").splitlines() if x.strip()]
    for ln in reversed(lines):                       # 마지막 예외 줄이 제일 정확하다
        if re.match(r"^[A-Za-z_.]*(Error|Exception)\b.*:", ln):
            return ln[:300]
    return (lines[-1][:300] if lines else "")


def _gate_which_test(out):
    """어느 검증에서 죽었나 — 트레이스백의 **마지막 t프레임**.

    2026-08-25 실사고: 예전에는 `tNNN...()` 라는 **호출 글자**만 찾았다. 그런데
    파이썬 트레이스백에서 호출 줄은 **바깥 프레임**이고 진짜 터진 자리는
    `File ..., line N, in tNNN...` 이라 **괄호가 없다.** 그래서 자국이
    `t202_layer_dialogs` 를 대는데 실제 범인은 `t201_upload_intake` 였다 —
    조치가 "그 검증부터 본다" 이므로 사람이 **멀쩡한 t202 를 뒤진다**([172]).
    ★ 마지막 프레임이 헬퍼(`in book`)면 그 바깥의 **마지막 t프레임**이 답이다.
    ★ 폴백은 남긴다([172]) — 트레이스백 없이 요약만 있는 자취도 있다.
    """
    text = str(out or "")
    frames = re.findall(
        r'^\s*File "[^"]*", line \d+, in (t\d+[A-Za-z0-9_]*)\s*$',
        text, re.M)
    if frames:
        return frames[-1]
    hits = re.findall(r"\b(t\d+[A-Za-z0-9_]*)\(\)", text)
    return hits[-1] if hits else ""


# 갈래마다 **조치가 다르다** — 'resource' 에 "코드를 고치세요"라고 적으면
# 사람이 멀쩡한 코드를 뒤진다(`[289]` 가 ERP API 에서 배운 자리).
_GATE_FIX = {
    "resource": ("Z:(SMB)·관리대장을 그 순간 못 읽었다 — **고칠 코드가 없을 수 있다**. "
                 "연결을 확인하고 `python tests/synthetic_check.py` 를 다시 돌린다."),
    "auth":     ("로그인(밴드·이카운트)이 필요하다 — 사람 인증 뒤 다시 돌린다."),
    "timeout":  ("시간을 넘겼다 — 기계가 바쁘면 `[6]`·`[192]` 처럼 그때만 빨갛다. 다시 돌려 본다."),
    "code":     ("자원·인증 표시가 없었다 — **검증이 실제로 빨갈 가능성이 높다**. "
                 "`python tests/synthetic_check.py > out.txt 2>&1` 로 돌려 ALL GREEN 글자를 눈으로 본다."),
    # ★ 이 갈래는 **검증이 빨갛다는 뜻이 아니다** — 그래서 조치도 다르다([289]).
    "편집중":    ("이 회차가 관문을 **도는 동안 소스가 바뀌었다** — 그 결과는 어느 판을 "
                 "검사한 것인지 알 수 없다([412]). **검증이 빨갛다는 뜻이 아니고 "
                 "고칠 코드가 없을 수 있다.** 지금 코드로 "
                 "`python tests/synthetic_check.py > out.txt 2>&1` 를 다시 돌려 "
                 "ALL GREEN 글자를 눈으로 본다 — 초록이면 이 자국은 지나간 것이다."),
}


def _gate_fix(kind, which):
    """갈래별 조치 — **검증 이름을 아는데 안 적으면** 사람이 그것을 다시 찾는다.

    `code` 는 "검증이 실제로 빨갛다" 는 뜻인데, 어느 검증인지는 `_gate_which_test`
    가 이미 뽑아 뒀다. 그것을 조치 맨 앞에 붙여 준다(`[169]` — 아는 것을 안 적지 않는다).
    """
    fix = _GATE_FIX.get(kind, "갈래를 못 가렸다 — 아래 자취를 그대로 읽는다.")
    if which and kind in ("code", ""):
        fix = "**%s** 가 막았다 — 그 검증부터 본다. " % which + fix
    return fix


# ★ 마지막 자국에서 이만큼 안에 부팅했을 때만 "꺼져서 죽었다"고 말한다.
#   가장 긴 단계가 30분(단계 제한시간)이고 종료·부팅에 몇 분이 든다.
REBOOT_WINDOW_MIN = 45


def _reboot_note(last_iso):
    """앞 회차 자국 **뒤에** 이 PC 가 부팅했으면 그 시각을 짧게 돌려준다(아니면 "").

    ★ 2026-08-28 실사고 — 일일대조가 140.5분째 61번째 단계에서 사라졌는데 자국은
      후보로 **워치독·이름으로 죽이는 자리**만 적었다. 실제로는 그 3.5분 뒤 사람이
      시작 메뉴에서 **전원 끄기**를 누른 것이었다(재부팅 13:59:22 · 빠른 시작이라
      40초 만에 돌아왔다). 그 조치를 따랐으면 **멀쩡한 코드를 뒤진다**([172]).

    ★ **부팅은 자는 것과 다르다** — 반드시 모든 프로세스를 죽인다. 그래서
      `[385]`·`[468]` 의 절전 '완화'와 달리 여기서는 원인이 **확정**된다.

    ★ **못 재면 아무 말도 안 한다**([169]) · **자국보다 앞선 부팅은 근거가 아니다**.
    ⚠ 재는 자리는 `system_audit._boot_time` **한 곳**이다([162]) — 여기서 다시
      물으면 같은 물음에 두 답이 생긴다. 순환을 피해 **늦게 들여온다**.
    """
    if not last_iso:
        return ""
    try:
        if ROOT not in sys.path:
            sys.path.insert(0, ROOT)
        import system_audit          # noqa: PLC0415  (늦게 들여온다 — 순환 방지)
        boot, _why = system_audit._boot_time()
    except Exception:                # noqa: BLE001
        return ""                    # 못 재면 지어내지 않는다([169])
    if boot is None:
        return ""
    try:
        last = datetime.fromisoformat(str(last_iso))
    except (TypeError, ValueError):
        return ""
    # ⚠ 자국은 `+09:00` 이 붙은 값이고 부팅 시각은 **로컬 naive** 다. 그냥 비교하면
    #   `TypeError` 로 죽는다 — [212] 가 시각 모양에서 겪은 그 자리이고, 그때는
    #   인계 문서가 통째로 안 나왔다. 둘 다 이 PC 시각이므로 tz 만 떼면 된다.
    if last.tzinfo is not None:
        last = last.replace(tzinfo=None)
    # ★ **부팅했다는 것만으로 원인이라 단정하지 않는다**([172]). 그 사실이 말해
    #   주는 것은 '언젠가 PC 가 꺼졌다'까지다 — 코드 고장으로 먼저 죽고 **몇 시간
    #   뒤** 사람이 PC 를 끈 경우까지 재부팅 탓으로 돌리면, 그것이 곧 이 고침이
    #   막으려던 **틀린 지목**이 된다.
    #   창은 실측에서 나온다: 이 회차의 가장 긴 단계가 **30분**(단계 제한시간)이고
    #   종료·부팅에 몇 분이 든다. 그러므로 종료가 죽인 경우는 마지막 자국에서
    #   `REBOOT_WINDOW_MIN` 안에 들어온다(실측 2026-08-28 은 **3.5분**).
    gap = (boot - last).total_seconds() / 60.0
    if gap <= 0:
        return ""                    # 자국보다 앞선 부팅은 이 죽음과 무관하다
    if gap > REBOOT_WINDOW_MIN:
        return ""                    # 한참 뒤 부팅 — 원인이라 말할 근거가 없다([169])
    return boot.strftime("%m-%d %H:%M")


def _note_prev_crash():
    """앞 회차가 **단계 도중에** 사라졌으면 그 단계 이름을 자국으로 남긴다 (분담판 [140]).

    ★ `[304]` 가 만든 자국은 **0단계 관문 전용**이다. 그래서 나머지 79단계는 죽어도
      왜인지가 아무 데도 안 남았다 — 실측 2026-08-19: 09:50 회차가 13:33 에 떠서
      13단계를 끝내고 14:03 '보험 입금 확인'(pid 26800)에서 **사라졌는데**
      `reports/` 에 `*_오류.json` 이 하나도 없었다. 그 단계를 직접 부르면 300초 안에
      exit 0 으로 멀쩡히 끝난다 — 즉 그 단계가 터진 것이 아니라 **밖에서 죽었다.**
      경과 29.6분이라 회차 예산(150분)도 작업 제한(PT3H)도 아니었다.

    ★ **근거는 지어낼 것이 없다** — `note_progress` 는 `finally` 에서 반드시
      `(회차 끝)` 을 찍는다. 그 표식이 **없으면** 그 회차는 끝을 못 본 것이다.
      있으면 아무 말도 안 한다(정상 완주든 단계 실패든 제 이름으로 이미 적힌다).

    ★ **왜인지는 지목하지 않는다**(`[172]`). 지금 댈 수 있는 것은 '어느 단계였나'
      까지다. 후보(워치독이 죽였나 · 이름으로 죽이는 자리가 남의 나무를 같이
      끊었나)를 나란히 적고 사람이 고른다 — 확언하면 멀쩡한 코드를 뒤지게 된다.

    ★ **여기는 락을 잡은 뒤**다 — 다른 daily_run 이 도는 중이면 위에서 이미 물러났다.
      그러므로 남아 있는 진행 파일은 **끝난 회차**의 것이다.

    ★ 못 읽으면 **아무 말도 안 한다** — 없는 사고를 지어내면 매 회차 거짓 자국이
      하나씩 쌓여 아무도 안 본다(`[170]`).

    ⚠ 지우는 자리는 **회차 시작**이지 끝이 아니다. 끝에서 지우면 이 회차가 만든
      자국을 이 회차가 도로 지워 **아무도 못 본다.** 자국이 사라지는 근거는
      '다음 회차가 `(회차 끝)` 까지 갔다' 하나다(`[228]` 의 '성공하면 지운다').
    """
    try:
        if os.path.exists(STEP_CRASH):
            os.remove(STEP_CRASH)
    except Exception:                     # noqa: BLE001
        pass
    prev = {}
    try:
        with open(PROGRESS, encoding="utf-8") as fh:
            prev = json.load(fh)
    except (OSError, ValueError):
        return                            # 첫 회차이거나 못 읽음 — 지어내지 않는다
    if not isinstance(prev, dict):
        return
    step = str(prev.get("단계") or "")
    if not step or step == "(회차 끝)":
        return                            # 끝을 봤다 — 여기서 할 말이 없다
    done = list(prev.get("끝난단계") or [])
    pid = prev.get("pid") or prev.get("주인pid") or "?"
    # ★ **부팅했으면 그것이 답이다 — 그리고 맨 앞에 세운다** (2026-08-28 실사고).
    #   `schedule_watch.traces()` 는 `무엇` 을 **120자에서 자른다** — 뒤에 붙이면
    #   정작 원인이 잘려 안 보인다([292]·[325] — 비지 않는 것을 맨 앞에 세운다).
    booted = _reboot_note(prev.get("시각"))
    if booted:
        무엇 = ("이 PC 가 %s 에 **부팅했다** — 앞 회차는 그때 꺼져서 죽었다"
                "(코드 고장이 아니다). '%s' 단계 · 끝낸 %d개 · %s분째 · pid %s"
                % (booted, step, len(done), prev.get("경과분"), pid))
        조치 = ("**코드를 뒤지지 않는다**([172]) — 워치독도 이름으로 죽이는 자리도"
                " 무관하다. 다음 예정 회차가 처음부터 다시 돈다."
                " 예정을 보려면 `python schedule_watch.py --print`.")
    else:
        무엇 = ("앞 회차가 '%s' 단계에서 사라졌다 — `(회차 끝)` 표식을 못 찍었다"
                " (끝낸 단계 %d개 · 그 단계에서 %d분 · 회차 %s분째 · pid %s)"
                % (step, len(done), int((prev.get("단계경과초") or 0) // 60),
                   prev.get("경과분"), pid))
        조치 = ("그 단계를 직접 돌려 본다 — 멀쩡히 끝나면 단계가 터진 것이 아니라"
                " **밖에서 죽은 것**이다. 그때 볼 후보 둘: 워치독 30분 회차가 무엇을"
                " 죽였나 · 이름으로 죽이는 자리가 남의 나무를 같이 끊었나."
                " 확언하지 말 것 — 지금 아는 것은 '어느 단계였나'까지다.")
    try:
        os.makedirs(REPORT_DIR, exist_ok=True)
        with open(STEP_CRASH, "w", encoding="utf-8") as fh:
            json.dump({"시각": datetime.now().isoformat(timespec="seconds"),
                       "명령": "daily_run.py (%s 단계)" % step,
                       # ★ 부팅이 확인되면 그것은 **모름이 아니다**. 이 칸을 읽는
                       #   코드는 없지만(실측), 사람이 JSON 을 열었을 때 '모름'
                       #   이라 적혀 있으면 그 자체가 틀린 말이 된다([169]).
                       "갈래": "재부팅" if booted else "모름",
                       "부팅": booted or "",
                       "단계": step,
                       "무엇": 무엇,
                       "조치": 조치,
                       "앞회차": {"시각": prev.get("시각"), "상태": prev.get("상태"),
                                "끝난단계": done[-8:],
                                "느린단계": prev.get("느린단계") or []}},
                      fh, ensure_ascii=False, indent=1)
    except Exception:                     # noqa: BLE001
        pass                              # 자국을 남기려다 회차를 막지 않는다


def _leave_gate_trace(step):
    """0단계 관문이 막았을 때 **왜인지를 디스크에 남긴다** (2026-08-18 실사고).

    ★ 실측: 09:50 회차가 **7일 중 6일**(8/10·12·14·15·16·18) 이 한 줄에서 죽었는데
      그 이유가 **어느 화면에도 안 떴다.** 스케줄러는 `exit 1`, `schedule_watch` 는
      "0 이 아닌 값으로 끝났다", 인계 문서도 같은 말이다. 진짜 이유
      (`FileNotFoundError: 관리대장을 찾을 수 없음: Z:/…`)는 **종합리포트 파일을
      열어야** 나왔다. 그래서 여섯 번 반복되는 동안 아무도 못 고쳤다 — `[228]` 이
      '어느 회차가 죽었나'를 보이게 만든 다음 질문('왜')에 답할 자국이 여기 없었다.
    ★ 이 단계는 **자율복구 대기열에 일부러 안 들어간다**(`autopilot.defer` 첫 줄) —
      안전문을 '나중에'로 돌리고 업무를 계속하면 안 되기 때문이다. 그 판단은 옳다.
      그러므로 남는 길은 **자국뿐**이다.
    ★ **갈래는 새로 만들지 않는다**(`[162]`) — `autopilot.classify_failure` 가 이미
      가른다(실측: 그 함수의 자기 시험이 바로 이 문자열을 `resource` 로 판정한다).
    ★ 성공하면 **지운다** — 옛 자국이 남으면 이미 지나간 고장을 계속 보고한다(`[228]`).
    """
    out = str(step.get("out") or "")
    # ★ **시간초과인지는 `run()` 이 이미 말해 준다** — 자취 글자로 다시 묻지 않는다
    #   (`[324]` 가 반대 방향에서 배운 그 자리다). `classify_failure` 는 출력에
    #   `timeout` 이라는 **낱말**만 있어도 timeout 이라 하는데, 합성검증 출력에는 그
    #   낱말이 **늘 들어간다**(이 프로젝트 검증이 `communicate(timeout=)`·
    #   `GATE_TIMEOUT_S` 를 재고 그 이름을 찍는다). 그래서 검증이 assert 로 죽어도
    #   갈래가 `timeout` 이 되고 조치는 *"다시 돌려 본다"* 로 나간다 — 조치는 갈래마다
    #   다르므로(`[289]`) 그 한 줄이 사람을 엉뚱한 데로 보낸다(`[172]`).
    #   실측 2026-08-19 17:17: 진짜 원인은 `t326` 이 목을 잃고 진짜 Z: 를 읽은 것인데
    #   자국은 `갈래=timeout`·조치 "다시 돌려 본다" 라고 적어 두고 있었다.
    # ★ **소스가 바뀐 것은 갈래를 뒤집는 사실이다** (2026-08-28 실사고 · 위 `_run_gate`).
    #   `code` 는 "검증이 실제로 빨갛다"는 **주장**인데, 도중에 소스가 바뀌었으면 그
    #   주장이 성립하지 않는다 — 반쪽을 시험한 것이다.
    # ⚠ **좁게 뒤집는다**([172]): `timeout`(시간을 넘긴 것은 소스와 무관한 사실이다) ·
    #   `resource`(Z: 를 못 읽었다) · `auth`(로그인이 필요하다)는 **한 글자도 안 건드린다.**
    #   넓히면 진짜 자원 실패를 "편집중"으로 덮어 못 잡는 것보다 나빠진다.
    edited = bool(step.get("소스바뀜"))
    if out.startswith("시간초과("):
        kind = "timeout"                           # `run()` 이 준 결정적 증거
    else:
        try:
            import autopilot
            kind = autopilot.classify_failure(out)
        except Exception:                          # noqa: BLE001
            kind = ""                              # 모르면 '모름' — 지어내지 않는다(`[169]`)
        if kind == "timeout":
            # 시간초과가 아닌 것이 확실하다(위에서 갈렸다) — 낱말에 걸린 것이다.
            # 자원·인증 표시가 없었다는 뜻이므로 그것이 곧 `code` 갈래의 뜻이다.
            kind = "code"
        if edited and kind in ("code", ""):
            kind = "편집중"
    head = _gate_headline(out)
    which = _gate_which_test(out)
    무엇 = "합성검증이 막았다"
    if which:
        무엇 += " · %s" % which
    if head:
        무엇 += " · %s" % head
    # ★ **조용히 갈래만 바꾸지 않는다**([169]) — 무엇이 바뀌었는지 이름을 댄다.
    #   못 가렸으면(빈 목록) 그 사실만 적고 이름을 지어내지 않는다.
    if edited:
        names = [str(x) for x in (step.get("바뀐파일") or [])]
        말 = (", ".join(names[:4]) + (" 외 %d개" % (len(names) - 4) if len(names) > 4 else "")
              ) if names else "어느 파일인지는 못 가렸다"
        무엇 += " · ⚠ 이 회차가 관문을 도는 동안 소스가 바뀌었다(%s)" % 말
    try:
        os.makedirs(REPORT_DIR, exist_ok=True)
        with open(GATE_CRASH, "w", encoding="utf-8") as fh:
            json.dump({"시각": datetime.now().isoformat(timespec="seconds"),
                       "명령": "daily_run.py (0단계 합성검증)",
                       "갈래": kind or "모름",
                       "검증": which,
                       "무엇": 무엇,
                       "조치": _gate_fix(kind, which),
                       "자취": out[-4000:]}, fh, ensure_ascii=False, indent=1)
    except Exception:
        pass                    # 자국을 남기려다 종료를 막지 않는다


def _clear_gate_trace():
    """관문을 통과했으면 옛 자국을 지운다 — 안 지우면 고쳐진 고장을 계속 보고한다."""
    try:
        if os.path.exists(GATE_CRASH):
            os.remove(GATE_CRASH)
    except Exception:
        pass


def finish(steps, aborted=False):
    os.makedirs(REPORT_DIR, exist_ok=True)
    # 동기화 백본: 앱(웹·워크벤치)이 읽는 기계 판독용 상태 파일 — 에이전트가 유일한 작성자
    import json
    json.dump({"time": datetime.now().isoformat(), "aborted": aborted,
               "steps": [{"n": s["name"], "s": ("deferred" if s.get("deferred") else
                                                    ("ok" if s["ok"] else
                                                     ("skip" if s["ok"] is None else "fail")))}
                          for s in steps]},
              open(os.path.join(REPORT_DIR, "agent_status.json"), "w", encoding="utf-8"), ensure_ascii=False)
    base = os.path.join(REPORT_DIR, f"종합리포트_{datetime.now():%Y%m%d_%H%M}.md")
    with open(base, "w", encoding="utf-8") as f:
        f.write(f"# 쿠팡 업무 자동대조 종합리포트 — {datetime.now():%Y-%m-%d %H:%M}\n\n")
        if aborted:
            f.write("**★ 합성검증 실패로 중단 — 아래 로그 확인 후 코드 수정 필요**\n\n")
        f.write("| 단계 | 결과 |\n|---|---|\n")
        for s in steps:
            mark = ("⏳ 자동복구 대기" if s.get("deferred") else
                    ("✅" if s["ok"] else ("⏭ 스킵" if s["ok"] is None else "❌ 실패")))
            f.write(f"| {s['name']} | {mark} |\n")
        f.write("\n---\n")
        for s in steps:
            f.write(f"\n## {s['name']}\n```\n{s['out']}\n```\n")
    # 종합리포트가 완성된 뒤 자율복구 현황도 같은 시각으로 맞춘다. 실패해도
    # 종합리포트와 회차 잠금 해제는 반드시 진행되어야 한다.
    try:
        import autopilot
        autopilot.write_status()
    except Exception:
        pass
    print(f"\n종합리포트: {base}")
    for s in steps:
        mark = ("DEFER" if s.get("deferred") else
                ("OK " if s["ok"] else ("SKIP" if s["ok"] is None else "FAIL")))
        print(f"  [{mark}] {s['name']}")


if __name__ == "__main__":
    main()
