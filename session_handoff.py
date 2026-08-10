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


def stranded_editor():
    """**사람이 옛 버전을 열어 놓고 편집 중인가.** (2026-08-07 실측)

    엑셀은 파일을 열면 옆에 `~$이름.xlsx` 잠금파일을 만든다. 그 잠금이 최신본이 아닌
    **옛 vN** 에 걸려 있으면 조용한 손실이 시작된 것이다 — 그 사람이 v538 에 손으로
    적어 넣는 동안 회차가 v539·v540·v541 을 만들고, 다음 회차는 v541 에서 이어간다.
    **v538 에 적은 것은 어디에도 넘어가지 않는다.** 파일이 열려 있으니 오류도 안 난다.

    이건 "하루 두 번만 반영" 규칙이 줄여 주는 문제지, 없애 주는 문제가 아니다.
    사람에게 알리는 것 말고 기계가 할 수 있는 일은 없다 — 남의 엑셀을 닫지 않는다.
    """
    try:
        from ecount_reconcile import load_config, resolve_master
        cur = resolve_master(load_config()["reconcile"]["master_xlsx"])
        folder, latest = os.path.dirname(cur), os.path.basename(cur)
        out = []
        for p in glob.glob(os.path.join(folder, "~$*.xlsx")):
            target = os.path.basename(p)[2:]          # `~$` 를 뗀 것이 열려 있는 파일
            if target == latest:
                continue                              # 최신본을 보고 있는 건 정상이다
            out.append(target)
        return sorted(out)
    except Exception:
        return []


def pid_alive(pid):
    """그 프로세스가 아직 살아 있나. **시간보다 확실한 판정이다** —
    세션이 죽으면 45분을 기다릴 것 없이 그 자리에서 잔재로 볼 수 있다.
    판정이 안 되면 None 을 돌려 시간 기준으로 넘긴다(모르면 함부로 죽었다고 하지 않는다)."""
    if not pid:
        return None
    # ★ 판정은 pid_alive.py 한 곳에서 한다 (2026-08-06 실사고 · 검증 [121]).
    #   여기 있던 옛 판정은 윈도우에서 **끝난 프로세스도 살아 있다**고 했다 —
    #   OpenProcess 는 종료된 프로세스에도 핸들을 준다. 죽은 세션의 점유가
    #   영원히 안 풀리는 쪽으로 틀렸다.
    try:
        import pid_alive
        return pid_alive.alive(pid)
    except Exception:
        return None


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
        try:
            at = float(info.get("at") or 0)
            if at > 0:
                mins = int((_t.time() - at) // 60)
        except (TypeError, ValueError):
            pass
        # ★ 주인은 `pid` 가 아니라 **`agent_pid`** 다 (2026-08-06 실사고).
        #   `pid` 는 ai_claim 을 실행한 CLI 프로세스라 명령이 끝나는 즉시 죽는다 —
        #   그것으로 판정하면 **살아 있는 옆 세션의 점유까지 '죽은 잔재'로** 표시하고,
        #   "이 명령으로 푸세요" 라고 안내한다. 실제로 그 안내대로 하면 ai_claim 이
        #   거부하므로(남의 것) 사람은 영문도 모른 채 막힌다. 판정은 한 벌이어야 한다.
        owner_pid = info.get("agent_pid") or info.get("pid")
        alive = pid_alive(owner_pid)
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
        # ★ '밀렸다'와 '밴드가 조용하다'는 다른 일이다 (2026-08-07 지시).
        #   날짜 있는 최신 글만 보면 새 글이 없는 날도 밀림으로 나온다. 그 경보를 믿고
        #   없는 번호를 긁으면 오늘처럼 쓰레기가 캐시로 들어간다. 그래서 '수집 최대 번호
        #   바로 다음이 없음으로 확인'된 근거가 **최근 것일 때만** 밀림을 내린다.
        #   근거가 오래됐으면 그 사이에 새 글이 올라왔을 수 있으므로 그대로 밀림이다.
        if row["밀림"] and name.startswith("밴드:"):
            q = quiet.get(name.split(":", 1)[1].strip()) or {}
            seen = str(q.get("확인시각") or "")[:10]
            if seen:
                try:
                    age = (datetime.strptime(day, "%Y-%m-%d")
                           - datetime.strptime(seen, "%Y-%m-%d")).days
                except ValueError:
                    age = None
                if age is not None and age <= limit:
                    row["밀림"] = False
                    row["조용함"] = "%s번까지 수집 완료 · %s 에 새 글 없음 확인" % (
                        q.get("수집최대"), q.get("확인시각"))
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
        "밴드재수집": band_recollect(),
        "업무흐름": work_flow_change(),
        "앱서버": app_server_health(),
    }


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


def app_server_health():
    """앱 서버가 **옛 코드로 돌고 있나** (2026-08-08 실사고).

    그날 반나절을 이것에 썼다. 어제 저녁에 뜬 서버가 하루치 변경을 하나도 반영하지
    못한 채 돌고 있었는데, 서버는 200 을 주고 화면도 숫자를 보여 줬다 —
    **고친 사람만 모르고 있었다.** 코드를 고치고 "왜 안 바뀌지"를 반복하는 것이
    이 사고의 모양이다. 그래서 세션 인계가 매번 이것을 본다.
    """
    try:
        sys.path.insert(0, os.path.join(ROOT, "webapp"))
        import restart_server
        s = restart_server.stale()
        if not s:
            return {"떠있음": bool(restart_server.running()), "옛코드": False}
        return {"떠있음": True, "옛코드": True, "pid": s[0], "뜬시각": s[1],
                "더새로운파일": s[2]}
    except Exception:
        return {}


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
        from pid_alive import alive
    except Exception:
        return None
    try:
        d = json.load(open(os.path.join(REPORT_DIR, ".daily_run.lock"), encoding="utf-8"))
        if not alive(d.get("pid")):
            return None
        started = datetime.fromisoformat(str(d.get("started_at")))
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
    return {"완주없음": False, "경과시간": round(age_h, 1), "중단": aborted,
            "실패단계": [f for f in failed if f], "진행중": running,
            "밀림": age_h >= DAILY_STALE_H or aborted}


def daily_step_now():
    """지금 회차가 **어느 단계**에 있나 — `.daily_run.progress.json` 을 읽는다.

    ★ 2026-08-09 지시("32시간째 미완주 왜그런거야"). 그때까지 경보는 '몇 시간째'만
      말할 수 있었다. 종합리포트는 **맨 끝에 한 번** 써지므로 완주하지 못한 회차는
      **기록을 한 줄도 안 남긴다** — 그래서 원인을 물어도 댈 말이 없었다.
      이제 daily_run 이 단계마다 자국을 남기고, 여기서 그것을 읽어 **이름을 댄다.**
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
    return {"단계": d.get("단계"), "상태": d.get("상태"), "머문분": (round(since, 1) if since is not None else None),
            "경과분": d.get("경과분"), "예산분": d.get("예산분"),
            "끝낸수": len(d.get("끝난단계") or [])}


def _step_hint():
    """경보 뒤에 붙일 한 줄 — **어느 단계에 머물러 있나**. 없으면 빈 문자열."""
    s = daily_step_now()
    if not s or not s.get("단계"):
        return ""
    txt = " · 지금 단계: **%s**(%s)" % (s["단계"], s.get("상태") or "")
    if s.get("머문분") is not None:
        txt += " %.0f분째" % s["머문분"]
    if s.get("끝낸수"):
        txt += " · 끝낸 단계 %d개" % s["끝낸수"]
    if s.get("경과분") and s.get("예산분"):
        txt += " · 회차 %.0f/%s분" % (s["경과분"], s["예산분"])
    return txt


def blockers(st, for_sol=False):
    """다음 세션이 **먼저 처리해야** 하는 것 — 안 하면 조용히 어긋난다."""
    out = []
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
    # ★ 스케줄러가 '성공'이라 말해도 완주하지 않았을 수 있다 — 잠금을 못 잡은 회차가
    #   조용히 exit 0 으로 끝나기 때문이다. 그 사이 자료현황·대조 리포트가 통째로 멈춘다.
    dr = st.get("일일대조") or {}
    run_h = dr.get("진행중")
    if dr.get("밀림"):
        hrs = dr.get("경과시간")
        why = ("마지막 회차가 **중단**으로 끝났다" if dr.get("중단")
               else "%s 완주하지 않았다" % ("한 번도" if hrs is None else "%.0f시간째" % hrs))
        bad = dr.get("실패단계") or []
        # ★ 지금 돌고 있으면 **띄우라고 하면 안 된다** — 잠금에 막혀 조용히 건너뛴다.
        #   기다리라고 말해야 한다. 조치가 정반대라 이 한 줄이 갈림길이다.
        act = ("(지금 %.0f시간째 돌고 있다 — 새로 띄우지 말고 끝나기를 기다린다)" % run_h
               if run_h is not None
               else "python daily_run.py    # 먼저 tasklist 로 앞 회차가 도는지 확인")
        why += _step_hint()
        out.append(("일일자동대조 — %s. 스케줄러는 '성공'으로 보고한다"
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
    if ap.get("옛코드"):
        out.append(("앱 서버가 **옛 코드**로 돌고 있다 (pid %s · 뜬 시각 %s) — "
                    "%s 가 서버보다 새것이다. 고쳐도 화면이 안 바뀐다"
                    % (ap.get("pid"), ap.get("뜬시각"),
                       ", ".join((ap.get("더새로운파일") or [])[:4])),
                    "python webapp/restart_server.py"))
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
        if f.get("밀림"):
            out.append(("★ %s 수집이 밀렸다 — 최신 %s (%d일 전, 한도 %d일). "
                        "지금 화면·보고 숫자는 그만큼 **적게** 나온다"
                        % (f["이름"], f["최신"], f["밀린일"], f["한도"]),
                        f["되살리는법"]))
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
            tag = " ★밀림" if f.get("밀림") else (" (조용함)" if f.get("조용함") else "")
            L.append("| %s%s | %s | %s | %d |"
                     % (f["이름"], tag, f["최신"], late, f["한도"]))
        quiet = [f for f in fr if f.get("조용함")]
        L += ["", "> 밴드·이카운트는 **사람 로그인**이 있어야 긁힌다(절대규칙 3).",
              "> 밀려 있으면 화면 숫자가 그만큼 적게 나온다 — 숫자를 의심하기 전에 여기부터 본다."]
        if quiet:
            # 날짜만 보면 밀린 것처럼 보이지만 받을 것이 없는 밴드 — 왜 안 긁어도 되는지 적는다.
            # 이 줄이 없으면 다음 세션이 또 없는 번호를 긁는다(2026-08-07 사고).
            L += ["> **조용함**: 최신 글이 오래됐지만 그 위로 새 글이 없음을 확인한 것이다 —"
                  " 긁을 것이 없다. 없는 번호를 긁으면 쓰레기가 캐시에 들어간다."]
            L += ["> · %s — %s" % (f["이름"], f["조용함"]) for f in quiet]
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


def write_snapshot(st, for_sol=False):
    """reports/세션인계.md|json 갱신 — 워치독(--snapshot)과 --adopt 가 같은 것을 남긴다."""
    os.makedirs(REPORT_DIR, exist_ok=True)
    open(os.path.join(REPORT_DIR, "세션인계.md"), "w", encoding="utf-8").write(
        to_md(st, for_sol=for_sol))
    json.dump(st, open(os.path.join(REPORT_DIR, "세션인계.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1, default=str)
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
