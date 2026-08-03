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
        for base, dirs, files in os.walk(src_root):
            rel_base = os.path.relpath(base, src_root)
            top = rel_base.split(os.sep)[0].lower() if rel_base != "." else ""
            dirs[:] = [d for d in dirs if d.lower() not in SKIP_DIRS
                       and not (rel_base == "." and d.lower() in excludes)]
            if top in excludes:
                continue
            for name in files:
                if name.lower() in SKIP_NAMES or name.startswith("~$"):
                    continue
                src = os.path.join(base, name)
                try:
                    st = os.stat(src)
                except OSError:
                    continue
                if now - st.st_mtime < MIN_STABLE_SECONDS:
                    continue                      # 저장 중일 수 있다 — 다음 회차에
                key = owner + "/" + os.path.relpath(src, src_root).replace("\\", "/")
                sig = [st.st_size, int(st.st_mtime)]
                if state.get(key) == sig:
                    continue
                rel = os.path.relpath(src, src_root)
                dst = os.path.join(upload_dir, "공유폴더_동기화", owner, rel)
                try:
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    tmp = dst + ".part"
                    shutil.copy2(src, tmp)
                    os.replace(tmp, dst)
                except OSError:
                    continue                      # 네트워크 순단 — 상태 미기록, 재시도됨
                state[key] = sig
                copied.append({"파일": rel, "보낸이": owner, "목적지": dst})
    _state_save(state)
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
