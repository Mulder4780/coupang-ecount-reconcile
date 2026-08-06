# -*- coding: utf-8 -*-
"""worktree_state.py — 워크트리에서 일해도 **상태는 하나**를 보게 한다

사용자 지시(2026-08-06): "이 워크트리도 다른 계정으로 로그인할 때 관리 가능하게
추가하는 알고리즘 정리하고 보고해"

무엇이 문제였나 (2026-08-06 실측)
  Claude Code 가 `.claude/worktrees/<이름>` 에 git 워크트리를 만들어 일을 시킨다.
  워크트리는 **추적 파일만** 체크아웃한다. 그런데 이 프로젝트의 상태는 거의 전부
  git 밖에 있다(.gitignore: `reports/`·`updates/`·`config/*.json`·`db/*.db`).
  그래서 워크트리 안에서는 모듈들이 제 폴더 기준으로 경로를 잡는 순간
  **본체와 다른 상태**를 보게 된다. 실제로 확인한 것:

    · `ai_claim` 의 점유 파일이 갈린다 — 워크트리 `reports/ai_claims.json` 과
      본체 `ecount/reports/ai_claims.json` 이 **다른 파일**이다. 두 세션이
      동시에 `ledger` 를 잡아도 서로 안 보인다. 관리대장 동시 쓰기 금지가
      조용히 무너진다(CLAUDE.md 동시 작업 규칙의 근간).
    · `ledger_db` 의 `db/ledger_queue.db` 가 없다 — 워크트리에서 `enqueue()` 한
      입력은 **11:00·15:00 반영이 영원히 못 본다**. 큐에 넣었으니 됐다고 믿는
      동안 값이 사라진다.
    · `config/ecount_config.json` 이 없다 — `tests/synthetic_check.py` 가
      t1_erp_check 에서 죽는다. 즉 **"ALL GREEN 확인 후 실작업" 관문 자체를
      통과할 수 없다**(실측: FileNotFoundError, 워크트리 첫 실행).
    · `session_handoff.rule_copies()` 가 `dirname(BASE)` 를 루트로 봐서
      `.claude/worktrees/CLAUDE.md` 를 찾는다 → 없으니 "정본과 다르다"는
      **거짓 경보**가 매번 '먼저 처리할 것' 맨 위에 뜬다(실측: 해시는 동일했다).

무엇을 고르는가 — 대상마다 방법이 다르다. 폴더가 '전부 무시'인지가 갈림길이다.
  · `updates/`      → **정션**(디렉터리 링크). 통째로 무시 대상이라 안전하다.
  · `config/*.json` → **하드링크**(파일 단위). 폴더에 추적 파일
                      (`*.example.json`)이 섞여 있어 통째로 잇지 못한다.
  · `reports/ai_claims.json` → **링크 금지, 코드가 본체 경로를 쓴다.**
                      `os.replace()` 로 갈아치우는 파일이라 링크가 그 순간 끊긴다.
  · `db/*.db`       → **링크 금지, 코드가 본체 경로를 쓴다.**
                      SQLite 는 `-wal`·`-journal` 사이드카를 제 경로 옆에 만든다.
                      링크로 두 경로를 만들면 사이드카가 갈려 DB 가 깨진다.

사용
  python worktree_state.py            # 지금 어떤 상태인가(읽기만)
  python worktree_state.py --apply    # 이을 수 있는 것을 잇는다
  from worktree_state import shared   # shared("db") → 본체의 db 경로
"""
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.abspath(__file__))

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# 정션으로 통째로 잇는 폴더 — `.gitignore` 에서 **폴더 전체**가 무시되는 것만 온다.
# 추적 파일이 하나라도 섞이면 워크트리 git 이 그 파일을 '삭제됨'으로 본다.
#
# ★ `reports/` 는 일부러 뺐다. 통째로 무시되는 것처럼 보이지만 추적 파일이 하나
#   섞여 있고(`reports/클로드_미완료_정리_20260801.md`), 실측(2026-08-06) 결과
#   본체와 워크트리의 그 파일 내용이 **서로 달랐다**. 정션으로 이으면 워크트리 git 이
#   남의 체크아웃 파일을 제 것으로 보고 '수정됨'으로 잡거나, 브랜치를 바꿀 때
#   본체 파일에 덮어써 버린다. 그래서 `reports/` 안에서 **꼭 공유해야 하는 것만**
#   코드가 본체 경로로 집는다(CODE_SHARED).
LINK_DIRS = ("updates", "band/cache", "band/ocr_cache")

# 하드링크로 개별로 잇는 비밀 설정. 폴더(`config/`)에 추적되는 예시 파일이 있어
# 통째로 못 잇는다. 손으로만 고치는 파일이라 하드링크가 끊길 일이 드물다.
LINK_FILES = (
    "config/ecount_config.json",
    "config/webapp.json",
    "config/gcal.json",
    "config/erp_allowed_ips.json",
    "config/cloud.json",
    "config/cal_share.local.json",
    "config/cloud_queue.local.json",
    "config/manual_events.local.json",
    "config/staff_contacts.local.json",
)

# 링크하지 않고 **코드가 본체 경로를 직접 쓰는** 것들. 여기 적힌 이유가 위 docstring 에 있다.
CODE_SHARED = ("reports/ai_claims.json", "db/ledger_queue.db")

_MAIN_ROOT = None


def _git(*args):
    try:
        r = subprocess.run(["git", *args], cwd=BASE, capture_output=True,
                           text=True, encoding="utf-8", errors="replace", timeout=30)
        return (r.stdout or "").strip() if r.returncode == 0 else ""
    except Exception:
        return ""


def main_root():
    """공용 상태의 주인 = **본체 체크아웃**의 절대경로.

    본체에서 부르면 자기 자신이 나온다(그래서 본체 동작은 하나도 안 바뀐다).
    워크트리에서 부르면 본체 경로가 나온다.

    판정은 `git rev-parse --git-common-dir` 로 한다. 워크트리의 `--git-dir` 은
    `.git/worktrees/<이름>` 이지만 `--git-common-dir` 은 언제나 **본체의 `.git`** 이다.
    """
    global _MAIN_ROOT
    if _MAIN_ROOT:
        return _MAIN_ROOT
    forced = os.environ.get("COUPANG_MAIN_ROOT")
    if forced and os.path.isdir(forced):
        _MAIN_ROOT = os.path.abspath(forced)
        return _MAIN_ROOT
    common = _git("rev-parse", "--git-common-dir")
    root = BASE
    if common:
        if not os.path.isabs(common):
            common = os.path.join(BASE, common)
        cand = os.path.dirname(os.path.abspath(common))
        # 본체라면 cand == BASE 다. 아니면 워크트리이고, 그 경로가 실재해야 믿는다.
        if os.path.isdir(cand):
            root = cand
    _MAIN_ROOT = os.path.abspath(root)
    return _MAIN_ROOT


def is_worktree():
    """지금 워크트리에서 돌고 있나."""
    return os.path.normcase(main_root()) != os.path.normcase(os.path.abspath(BASE))


def shared(*parts):
    """공용 상태 경로. 본체에서는 기존 경로와 **글자 그대로 같다**."""
    return os.path.join(main_root(), *parts)


def _same_file(a, b):
    """같은 실체(하드링크로 이어졌나)."""
    try:
        sa, sb = os.stat(a), os.stat(b)
    except OSError:
        return False
    if sa.st_ino and sa.st_ino == sb.st_ino and sa.st_dev == sb.st_dev:
        return True
    # Windows 에서 st_ino 가 0 으로 오는 경우가 있어 크기·수정시각으로 보조 판정한다.
    return sa.st_size == sb.st_size and int(sa.st_mtime) == int(sb.st_mtime)


def _is_link_dir(path, target):
    """이 경로가 **본체의 그 폴더와 같은 실체**인가.

    `os.path.islink()` 로 보면 안 된다 — 윈도우 **정션은 심볼릭 링크가 아니라서**
    `islink()` 가 False 를 돌려준다(실측 2026-08-06: 방금 만든 정션이 '따로있음'
    으로 잡혔다). 링크인지를 묻지 말고 **가리키는 곳이 같은지**를 묻는 게 맞다.
    """
    try:
        if not os.path.isdir(path):
            return False
        return (os.path.normcase(os.path.realpath(path))
                == os.path.normcase(os.path.realpath(target)))
    except Exception:
        return False


def _make_junction(link, target):
    """디렉터리 정션. **관리자 권한이 필요 없다**(심볼릭 링크와 달리)."""
    if os.name != "nt":
        os.symlink(target, link, target_is_directory=True)
        return True
    try:
        import _winapi
        _winapi.CreateJunction(target, link)      # (가리킬 곳, 만들 링크)
        return True
    except Exception:
        r = subprocess.run(["cmd", "/c", "mklink", "/J", link, target],
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=30)
        return r.returncode == 0


def status():
    """무엇이 이어졌고 무엇이 안 이어졌나. 아무것도 고치지 않는다."""
    out = {"워크트리": is_worktree(), "여기": os.path.abspath(BASE),
           "본체": main_root(), "항목": []}
    if not out["워크트리"]:
        return out
    for rel in LINK_DIRS:
        here = os.path.join(BASE, *rel.split("/"))
        there = shared(*rel.split("/"))
        if not os.path.isdir(there):
            state = "본체에없음"
        elif _is_link_dir(here, there):
            state = "이어짐"
        elif os.path.isdir(here):
            state = "따로있음"          # 실제 폴더가 이미 있다 — 덮지 않는다
        else:
            state = "없음"
        out["항목"].append({"대상": rel, "방법": "정션", "상태": state})
    for rel in LINK_FILES:
        here = os.path.join(BASE, *rel.split("/"))
        there = shared(*rel.split("/"))
        if not os.path.isfile(there):
            state = "본체에없음"        # 본체에도 없는 설정이면 할 일이 없다
        elif not os.path.exists(here):
            state = "없음"
        elif _same_file(here, there):
            state = "이어짐"
        else:
            state = "따로있음"
        out["항목"].append({"대상": rel, "방법": "하드링크", "상태": state})
    for rel in CODE_SHARED:
        out["항목"].append({"대상": rel, "방법": "코드해석",
                            "상태": "이어짐" if os.path.exists(shared(*rel.split("/")))
                                    else "본체에없음"})
    return out


def apply(dry=False):
    """이을 수 있는 것만 잇는다. **이미 따로 있는 것은 절대 덮지 않는다.**

    덮지 않는 이유: 워크트리에 사람이 일부러 다른 설정을 둔 경우를 지울 수 없다.
    그런 것은 결과에 '따로있음' 으로 남겨 사람이 판단하게 한다.
    """
    done, skip = [], []
    if not is_worktree():
        return {"워크트리": False, "이음": done, "건너뜀": skip}
    for item in status()["항목"]:
        rel, how, st = item["대상"], item["방법"], item["상태"]
        if how == "코드해석" or st in ("이어짐", "본체에없음"):
            continue
        if st == "따로있음":
            skip.append((rel, "워크트리에 따로 있어 덮지 않음"))
            continue
        here = os.path.join(BASE, *rel.split("/"))
        there = shared(*rel.split("/"))
        if dry:
            done.append((rel, how + "(예정)"))
            continue
        try:
            os.makedirs(os.path.dirname(here), exist_ok=True)
            if how == "정션":
                ok = _make_junction(here, there)
                if not ok:
                    skip.append((rel, "정션 생성 실패"))
                    continue
            else:
                os.link(there, here)
            done.append((rel, how))
        except Exception as e:
            skip.append((rel, "%s: %s" % (type(e).__name__, str(e)[:60])))
    return {"워크트리": True, "이음": done, "건너뜀": skip}


def _print():
    st = status()
    if not st["워크트리"]:
        print("본체 체크아웃입니다 — 공용 상태가 곧 제 폴더입니다.")
        print("  경로: %s" % st["여기"])
        return 0
    print("워크트리에서 실행 중입니다.")
    print("  여기: %s" % st["여기"])
    print("  본체: %s" % st["본체"])
    print("")
    bad = 0
    for i in st["항목"]:
        mark = {"이어짐": "  OK  ", "없음": "  ★   ", "따로있음": "  !   ",
                "본체에없음": "  -   "}.get(i["상태"], "  ?   ")
        if i["상태"] == "없음":
            bad += 1
        print("%s%-34s %-8s %s" % (mark, i["대상"], i["방법"], i["상태"]))
    if bad:
        print("\n★ %d개가 끊겨 있습니다 → python worktree_state.py --apply" % bad)
    else:
        print("\n모두 이어져 있습니다.")
    return 0


if __name__ == "__main__":
    args = sys.argv[1:]
    if "--apply" in args or "--dry" in args:
        res = apply(dry="--dry" in args)
        if not res["워크트리"]:
            print("본체 체크아웃이라 이을 것이 없습니다.")
        else:
            for rel, how in res["이음"]:
                print("이음   %-34s %s" % (rel, how))
            for rel, why in res["건너뜀"]:
                print("건너뜀 %-34s %s" % (rel, why))
            if not res["이음"] and not res["건너뜀"]:
                print("이미 모두 이어져 있습니다.")
        sys.exit(0)
    sys.exit(_print())
