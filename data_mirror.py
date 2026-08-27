# -*- coding: utf-8 -*-
"""원본 자료를 새 정본 자리로 **복사**한다 — 회차로, 진도를 남기며 (2026-08-27 지시).

형님 지시: **"DB 및 앱 관련 데이터는 `…COUPANG_INTEGRATED_WORK_AGENT` 여기서 계속
관리하고 (용량 커지지 않게) `Z:\\25. AI_RND_Data\\2000. ULH APPS DATA\\2. CSOS DATA`
이 폴더에 원본데이터 폴더 번호 순서대로 계속 저장하고 기존 자료도 이쪽으로 전부
복사해서 가져와서 관리해 / 알고리즘에 구현해"**

## 왜 한 번에 안 하나
회선이 느리다 — 실측 2026-08-27 로 폴더 확인 한 번에 **1.3초~28.9초**(붐빌 때)이고
`1. ERP 내보내기` 한 폴더를 **훑기만** 하는 데 102.2초다. 한 번에 다 하려다 끊기면
**진도가 0** 이 된다 — 이 저장소가 세 번 밟은 자리다([388] 밴드 수확 · [406] 보관
색인 · [427] 게시글 보관). **상한은 개수가 아니라 시간으로 정한다.**

## 진도를 어떻게 남기나 — 파일 목록을 안 쌓는다
"어느 파일을 했나"를 쌓으면 그 목록 자체가 수만 줄이 된다. 대신 **목적지에 있나**로
판정한다([381] `_dst_entries` 와 같은 생각) — 그러면 진도 파일이 없어도 매 회차가
"아직 없거나 다른 것"만 복사해 **저절로 이어진다**. 남기는 진도는
**"어느 폴더까지 갔나"** 하나뿐이라 작고, 잃어도 처음부터 다시 훑을 뿐 **일이
두 번 되지 않는다**(같은 파일은 `_same` 이 건너뛴다).

## 지키는 것
* **원본을 한 글자도 안 지운다** — 형님이 "복사" 라 하셨다. 지우는 것은 되돌릴 수 없다.
* **내용이 다르면 안 덮는다** — 이름을 바꿔 둘 다 남긴다([381] 과 같은 규칙).
  어느 쪽이 맞는지는 사람이 본다.
* **같은 파일 판정을 새로 만들지 않는다**([162]) — `collect_sources._same`
  (크기 + 수정시각 초). 느린 드라이브에서 해시는 과하다.
* **딸려 온 stat 을 버리지 않는다**([198]) — `source_index.walk_stat`.
  `os.stat(경로)` 는 Z: 에서 파일당 왕복 한 번(135~155ms)이고 목록에 딸려 온 stat 은
  0.04ms 다. 3,000배다.
* ⚠ **`skip_dirs=()` 를 명시한다**([198] 의 ⚠) — 공용 워커의 기본값은 *색인의*
  목록(`_보관`·`_바로가기`)이라, 말없이 물려받으면 그 폴더가 **조용히 빠지면서
  '완료'라고 적힌다**([165]). 여기서는 **전부** 옮겨야 한다.
* **목적지는 폴더마다 한 번만 scandir**([381]) — 파일당 왕복 0.
* **못 읽은 것을 조용히 넘기지 않는다**([169]) — 숫자로 말한다.

## 아직 안 한 것 ([169] — 안 한 것을 한 것처럼 적지 않는다)
**주소는 아직 안 바꿨다.** `source_dirs.ORIGIN_ROOT` 는 그대로 옛 자리를 가리킨다.
실측으로 원본 폴더 이름을 **194곳**이 직접 쓰므로 주소 한 줄이면 전부 따라오는데,
**복사 전에 바꾸면 194곳이 빈 폴더를 보고 "0건"이라고 조용히 확언한다**([169]).
그러니 순서는 ① 복사 회차(여기) ② 다 채워진 것 확인 ③ 주소 한 줄 ④ 옛 자리는
읽기로 남긴다(형님이 "복사" 라 하셨다).

사람이 보는 명령:
    python data_mirror.py                 # 무엇이 얼마나 남았나 (아무것도 안 옮긴다)
    python data_mirror.py --apply         # 실제로 복사 (예산 안에서)
    python data_mirror.py --apply --budget 1800
"""
import argparse
import json
import os
import shutil
import sys
import time

BASE = os.path.dirname(os.path.abspath(__file__))
if BASE not in sys.path:
    sys.path.insert(0, BASE)

import source_dirs as SD                                    # noqa: E402
from collect_sources import _same                           # noqa: E402
from source_index import walk_stat                          # noqa: E402

REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(BASE, "reports")
STATE = os.path.join(REPORT_DIR, "원본이전_진도.json")

#: 예산 — 자율복구·회차가 이 열쇠로 준다([427]). 없으면 **제한 없음**이다
#: (`[169]` — 못 읽은 것을 0초로 치면 아무것도 안 옮기면서 '완료'라 적는다).
BUDGET_ENV = "DATA_MIRROR_BUDGET_SEC"
SAVE_EVERY_S = 30.0      # 진도 저장 상한 — **개수가 아니라 시간**이다([388])


def _budget():
    try:
        v = float(os.environ.get(BUDGET_ENV) or 0)
    except (TypeError, ValueError):
        return None
    return v if v > 0 else None


def _load():
    try:
        with open(STATE, encoding="utf-8") as fh:
            d = json.load(fh)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def _save(d):
    """진도를 남긴다 — **못 남겨도 복사한 것을 되돌리지는 않는다**."""
    try:
        os.makedirs(os.path.dirname(STATE), exist_ok=True)
        tmp = STATE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(d, fh, ensure_ascii=False, indent=1)
        os.replace(tmp, STATE)
        return True
    except Exception:
        return False


def _dst_entries(d):
    """목적지 폴더를 **한 번만** 훑어 `{이름: stat}` 으로 든다([381])."""
    out = {}
    try:
        for e in os.scandir(d):
            if e.is_file():
                try:
                    out[e.name] = e.stat()
                except OSError:
                    pass
    except OSError:
        pass
    return out


def tops():
    """옛 원본 한 겹 밑 폴더 목록 — **번호 순서 그대로**.

    ★ `source_dirs` 의 상수 목록으로 고르지 않는다([169]) — 거기 없는 폴더
      (`10. 기준·참고 자료` 등)가 실재하고, 형님은 **전부** 옮기라 하셨다.
    """
    root = SD.ORIGIN_ROOT
    out = []
    try:
        for e in os.scandir(root):
            if e.is_dir():
                out.append(e.name)
    except OSError:
        return []
    out.sort()
    return out


def _one(rel, apply, deadline, seen):
    """폴더 하나를 옮긴다. 반환: (복사, 동일, 이름바꿈, 실패, 바이트, 예산끝)"""
    src_root = os.path.join(SD.ORIGIN_ROOT, rel)
    dst_root = os.path.join(SD.CSOS_DATA_ROOT, rel)
    copied = same = renamed = failed = 0
    nbytes = 0
    out_of_time = False
    ents_cache = {}

    for base, name, st in walk_stat(src_root, skip_dirs=()):
        if deadline and time.time() >= deadline:
            out_of_time = True
            break
        sub = os.path.relpath(base, src_root)
        dst_dir = dst_root if sub == "." else os.path.join(dst_root, sub)
        ents = ents_cache.get(dst_dir)
        if ents is None:
            ents = ents_cache[dst_dir] = _dst_entries(dst_dir)
        src = os.path.join(base, name)
        dst = os.path.join(dst_dir, name)
        if name in ents:
            if _same(src, dst, st, ents[name]):
                same += 1
                continue
            # 내용이 다르면 **안 덮는다** — 어느 쪽이 맞는지 사람이 본다([381])
            stamp = time.strftime("%y%m%d", time.localtime(st.st_mtime))
            root_, ext = os.path.splitext(name)
            name2 = "%s_%s%s" % (root_, stamp, ext)
            if name2 in ents and _same(src, os.path.join(dst_dir, name2),
                                       st, ents[name2]):
                same += 1
                continue
            dst = os.path.join(dst_dir, name2)
            renamed += 1
        else:
            copied += 1
        if apply:
            try:
                os.makedirs(dst_dir, exist_ok=True)
                shutil.copy2(src, dst)          # copy2 = 수정시각 보존 → 다음엔 '동일'
                nbytes += st.st_size
                try:
                    ents[os.path.basename(dst)] = os.stat(dst)
                except OSError:
                    pass
            except OSError:
                failed += 1
                if name in ents:
                    renamed -= 1
                else:
                    copied -= 1
        else:
            nbytes += st.st_size
        seen[0] += 1
    return copied, same, renamed, failed, nbytes, out_of_time


def run(apply=False, budget_s=None, only=None, on_folder=None):
    """폴더 하나를 끝낼 때마다 `on_folder(dict)` 를 부른다 — **죽어도 진도가 보인다**([180]).

    ★ `run` 이 직접 `print` 하지 않는 이유: 회차는 `pythonw` 로 돌아 `sys.stdout` 이
      **None** 이다([235]). 찍는 것은 `main()` 몫이다.
    """
    if budget_s is None:
        budget_s = _budget()
    deadline = (time.time() + budget_s) if budget_s else None
    out = {"복사": 0, "동일": 0, "이름바꿈": 0, "실패": 0, "바이트": 0,
           "폴더": [], "남은폴더": [], "예산끝": False, "왜못함": ""}

    if not os.path.isdir(SD.ORIGIN_ROOT):
        out["왜못함"] = "옛 원본 자리에 못 닿았다(네트워크 드라이브 연결부터 본다)"
        return out
    if apply:
        try:
            os.makedirs(SD.CSOS_DATA_ROOT, exist_ok=True)
        except OSError as e:
            out["왜못함"] = "새 자리를 못 만들었다(%s)" % (e.strerror or type(e).__name__)
            return out

    names = tops()
    if only:
        names = [n for n in names if n == only]
    st = _load()
    done = st.get("끝난폴더") or []
    # 지난 회차가 끝낸 폴더는 **뒤로 돌린다** — 아직 안 한 것을 먼저 한다.
    # (지운 것이 아니다 — 한 바퀴 돌면 다시 확인해 새로 생긴 파일을 나른다.)
    names = [n for n in names if n not in done] + [n for n in names if n in done]

    seen = [0]
    last_save = time.time()
    for rel in names:
        if deadline and time.time() >= deadline:
            out["예산끝"] = True
            out["남은폴더"].append(rel)
            continue
        if out["예산끝"]:
            out["남은폴더"].append(rel)
            continue
        c, s, r, f, b, over = _one(rel, apply, deadline, seen)
        out["복사"] += c
        out["동일"] += s
        out["이름바꿈"] += r
        out["실패"] += f
        out["바이트"] += b
        _f = {"이름": rel, "복사": c, "동일": s,
              "이름바꿈": r, "실패": f, "끝": not over, "바이트": b}
        out["폴더"].append(_f)
        if on_folder:
            try:
                on_folder(_f)
            except Exception:
                pass
        if over:
            out["예산끝"] = True
            out["남은폴더"].append(rel)
        elif apply and rel not in done:
            done.append(rel)
        if apply and time.time() - last_save >= SAVE_EVERY_S:
            _save({"끝난폴더": done, "시각": time.strftime("%Y-%m-%dT%H:%M:%S"),
                   "본파일": seen[0]})
            last_save = time.time()

    if apply:
        _save({"끝난폴더": done, "시각": time.strftime("%Y-%m-%dT%H:%M:%S"),
               "본파일": seen[0]})
    return out


def line(res):
    """회차가 로그에 남기는 한 줄 — **조용히 넘어가지 않는다**([169])."""
    if res.get("왜못함"):
        return "원본 이전 못 함: " + res["왜못함"]
    mb = res["바이트"] / 1024.0 ** 2
    out = ("원본 이전 복사 %d개 %.0fMB · 이미 있음 %d"
           % (res["복사"] + res["이름바꿈"], mb, res["동일"]))
    if res["실패"]:
        out += " · 못 옮김 %d" % res["실패"]
    if res["예산끝"]:
        out += " · 예산 끝(남은 폴더 %d개 — 다음 회차가 잇는다)" % len(res["남은폴더"])
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description="원본 자료를 새 정본 자리로 복사")
    ap.add_argument("--apply", action="store_true", help="실제로 복사한다")
    ap.add_argument("--budget", type=float, help="예산(초) — 넘으면 멈추고 진도를 남긴다")
    ap.add_argument("--only", help="폴더 하나만")
    a = ap.parse_args(argv)
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    print("옛 자리:", SD.ORIGIN_ROOT)
    print("새 자리:", SD.CSOS_DATA_ROOT)
    def _say(f):
        print("  %-38s 복사 %5d · 이미 %5d · 이름바꿈 %3d · 실패 %3d · %8.1fMB %s"
              % (f["이름"][:38], f["복사"], f["동일"], f["이름바꿈"], f["실패"],
                 f.get("바이트", 0) / 1024.0 ** 2,
                 "" if f["끝"] else "(예산 끝)"), flush=True)

    res = run(apply=a.apply, budget_s=a.budget, only=a.only, on_folder=_say)
    print(line(res))
    if not a.apply:
        print("※ 미리보기입니다 — 실제로 옮기려면 --apply")
    return 0


if __name__ == "__main__":
    sys.exit(main())
