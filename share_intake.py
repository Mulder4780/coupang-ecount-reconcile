# -*- coding: utf-8 -*-
"""16. Share 공유폴더를 **항상** 끌어온다 (2026-08-03 지시, 상시).

사용자 지시: "Z:\\16. Share\\유현민\\오종현 항상 이 폴더 데이터 긁어오는 알고리즘 추가."

* 원본은 지우지 않는다 — 오종현 매니저의 공유폴더는 그의 저장소다. 새/변경 파일만
  투입함(``100. 업로드용 자료/공유폴더_동기화/<보낸이>``)으로 **복사**하고,
  분류·정본 이동은 기존 ``upload_intake`` 가 이어서 한다.
* ``26년도 PO 모음``·``26년도 쿠팡 입금내역`` 은 이미 source_dirs 의 정본 원천이라
  (PO_DIRS·RECEIPT_DIRS, 제자리 직접 읽기) 여기서 다시 복사하지 않는다 — 두 벌이
  되면 중복 제거만 늘어난다.
* 상태(``db/share_pull.json``: 경로→크기·mtime)를 기억해 같은 파일을 두 번 옮기지
  않는다. 방금 저장 중인 파일(30초 미만)은 다음 회차로 미룬다.
* 실행 경로: ``upload_intake --apply`` 첫 단계에서 자동 호출 — 워치독 30분·09:35
  원본정리·09:50 일일대조가 모두 이 흡수를 거친다. 단독 실행: ``python share_intake.py``.
"""
import json
import os
import shutil
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

import source_index                               # noqa: E402  (sys.path 뒤라야 한다)

STATE = os.path.join(ROOT, "db", "share_pull.json")
MIN_STABLE_SECONDS = 30
SKIP_NAMES = {"thumbs.db", "desktop.ini"}
SKIP_DIRS = {"old"}

# (원본 폴더, 보낸이 표시명, 제외 하위폴더 — 이미 정본 원천으로 직접 읽는 것들)
def pull_targets():
    return [(
        r"Z:\16. Share\유현민\오종현", "오종현",
        {"26년도 po 모음", "26년도 쿠팡 입금내역"},
    )]


def _state_load():
    try:
        return json.load(open(STATE, encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _state_save(state):
    os.makedirs(os.path.dirname(STATE), exist_ok=True)
    tmp = STATE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as handle:
        json.dump(state, handle, ensure_ascii=False)
    os.replace(tmp, STATE)


def pull(targets=None, upload_dir=None, state_path=None):
    """새/변경 파일을 투입함으로 복사. [{파일,보낸이,목적지}] 를 돌려준다."""
    global STATE
    if state_path:
        STATE = state_path
    if upload_dir is None:
        from source_dirs import UPLOAD_DIR
        upload_dir = UPLOAD_DIR
    state = _state_load()
    copied, now = [], time.time()
    for src_root, owner, excludes in (targets or pull_targets()):
        if not os.path.isdir(src_root):
            continue
        # ★ 목록을 받을 때 딸려 온 stat 을 쓴다 — 파일마다 `os.stat` 을 다시 부르면
        #   Z:(SMB)에서 파일당 왕복 한 번이다(검증 [198]). 같은 폴더 실측:
        #   예전 `os.walk`+`os.stat` **99.7초 / 602개** → `walk_stat` **4.1초**(24배).
        #   이 함수는 5분마다(automation_pipeline) 도는 자리라 그대로 회차 비용이 된다.
        #   거르기는 `skip_dirs` 에 맡기지 않고 **경로 요소**로 한다 — `skip_dirs` 는
        #   대소문자를 정확히 맞춰야 해서 `Old`·`OLD` 가 조용히 새어 들어온다.
        for base, name, st in source_index.walk_stat(src_root, skip_dirs=()):
            rel = os.path.relpath(os.path.join(base, name), src_root)
            parts = [p.lower() for p in rel.split(os.sep)[:-1]]
            if any(p in SKIP_DIRS for p in parts):
                continue
            if parts and parts[0] in excludes:
                continue
            if name.lower() in SKIP_NAMES or name.startswith("~$"):
                continue
            src = os.path.join(base, name)
            if now - st.st_mtime < MIN_STABLE_SECONDS:
                continue                          # 저장 중일 수 있다 — 다음 회차에
            key = owner + "/" + rel.replace("\\", "/")
            sig = [st.st_size, int(st.st_mtime)]
            if state.get(key) == sig:
                continue
            dst = os.path.join(upload_dir, "공유폴더_동기화", owner, rel)
            try:
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                tmp = dst + ".part"
                shutil.copy2(src, tmp)
                os.replace(tmp, dst)
            except OSError:
                continue                          # 네트워크 순단 — 상태 미기록, 재시도됨
            state[key] = sig
            copied.append({"파일": rel, "보낸이": owner, "목적지": dst})
    _state_save(state)
    # ★ 공유폴더는 **올려도 조용한** 자리였다 — 파일이 투입함으로 들어가고 끝이라
    #   형님은 무엇이 새로 왔는지 물어봐야 알았다(2026-08-14 지시).
    #   알리는 자리는 부르는 쪽이 아니라 여기 하나다 — 이 함수는 워치독·09:35·09:50·
    #   5분 파이프라인이 다 부른다. 부르는 쪽마다 붙이면 사본이 넷이 된다([162]).
    #   같은 갈래가 5분 안에 여러 번이면 notify 가 한 줄로 합친다([170]).
    if copied:
        try:
            import notify
            보낸이 = sorted({row["보낸이"] for row in copied})
            notify.push(
                "공유폴더 자료",
                f"공유폴더에서 새 자료 {len(copied)}건을 받았습니다 — {', '.join(보낸이)}",
                "투입함에 넣었습니다 · 분류는 원본 정리 회차가 합니다.",
                evidence="share_intake.pull", 상태=f"{len(copied)}건")
        except Exception:
            pass                                  # 알리려다 흡수를 막지 않는다
    return copied


def main():
    copied = pull()
    print(f"공유폴더 끌어오기: {len(copied)}건")
    for row in copied[:10]:
        print("  ", row["보낸이"], "|", row["파일"][:70])
    if len(copied) > 10:
        print(f"   … 외 {len(copied) - 10}건")


if __name__ == "__main__":
    main()
