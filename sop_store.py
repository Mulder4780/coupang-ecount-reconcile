# -*- coding: utf-8 -*-
"""표준 업무 절차서 — 담는 자리 하나 (2026-09-10 형님 지시)

형님 지시: "절차서는 내가 지정한 폴더에 별도로 word나 ppt로 만들 수 있는 구조
준비해(앱에는 메뉴가 있어야함) / 그리고 절차서의 경우 나중에 유니웍스에 자료
넘겨서 유니웍스에서 통합관리할거야."

★ 우리가 최종 주인이 아니다.
  절차서의 마지막 자리는 **유니웍스**(노승용 매니저 앱 · '업무절차서' 메뉴가
  이미 있다 — 2026-09-10 형님 캡처).  그러므로 여기 담는 모양은
  **평평하고 넘기기 쉬워야** 한다 — 우리 앱 화면에 맞춘 모양으로 담으면
  넘길 때 통째로 다시 짜야 한다.
  · 칸 이름은 사람이 읽는 한글 그대로 둔다(유니웍스도 한글 앱이다).
  · 중첩은 `단계` 하나뿐이다.  더 깊게 만들지 않는다.
  · `출처`·`작성자`·`수정일` 을 반드시 담는다 — 넘긴 뒤 "이건 누가 언제
    적은 거냐"를 물을 사람이 반드시 있다.

★ 정본 규칙(2026-08-10)과의 자리.
  절차서는 **업무값이 아니라 문서**다(금액·상태가 아니다).  그래서 앱 DB
  대신 JSON 파일 하나에 담는다 — 통째로 내보내고 통째로 넘기는 것이 이
  자료의 쓰임이기 때문이다.  대신 **한 곳**이다([162]) — 화면·내보내기·
  유니웍스 넘김이 전부 이 파일을 읽는다.

★ Z: 를 한 글자도 안 만진다([168]).  웹 요청이 이것을 읽으므로 공유폴더를
  훑으면 그 화면이 회차와 다투며 몇 분씩 선다.
"""

import hashlib
import io
import json
import os
import re
import shutil
import sys

if hasattr(sys.stdout, "reconfigure"):          # 무인 회차는 sys.stdout 이 None 이다([235])
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

BASE = os.path.dirname(os.path.abspath(__file__))
REPORTS = os.path.join(BASE, "reports")

STORE = os.path.join(REPORTS, "표준업무절차서.json")
CONFIG = os.path.join(REPORTS, "절차서_설정.json")

# ── 낱말은 한 곳에서 온다([162]) ─────────────────────────────────────────
# 부위: 2026-09-09 유수비 대표 통화 그대로.  화면·워드·PPT·유니웍스 넘김이
# 전부 이 표를 읽는다 — 화면에 손으로 적으면 표를 고친 날 한쪽만 바뀐다([165]).
부위목록 = ["기계부", "전장부", "컨트롤박스부", "마스터부", "프레임부", "기타"]
작업종류목록 = ["돌발AS", "정기점검", "신규설치", "기타"]

# 절차서 한 건이 가지는 칸.  유니웍스로 넘길 때 이 목록이 곧 계약이다.
칸 = ("id", "제목", "부위", "작업종류", "공구", "단계", "마무리",
      "주의사항", "출처", "밴드글번호", "작성자", "수정일",
      "판", "원천지문", "자동생성", "이전판")

# 판을 올릴지 볼 때 **빼는** 칸.  ★ 여기에 시계·판 자신이 남아 있으면
#   내용이 하나도 안 바뀌어도 매 회차 "바뀜"이 되어 아무도 안 본다([170]).
_판무관 = ("판", "수정일", "이전판", "원천지문")
_판보관 = 5                                   # 이전 판을 몇 개까지 남기나

_ID_RE = re.compile(r"^SOP-(\d{3,})$")


def _now():
    """지금 시각(이 PC 기준).  검증이 목으로 갈아 끼울 수 있게 한 곳에 둔다."""
    from datetime import datetime
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def _empty():
    return {"판": 1, "갱신": "", "절차서": []}


def load(path=None):
    """담긴 절차서를 읽는다.

    ★ 못 읽으면 **빈 것**이 아니라 예외를 올리지 않고 빈 판을 준다 —
      다만 부르는 쪽이 '0건'과 '못 읽음'을 가를 수 있게 `_못읽음` 을 붙인다([169]).
      조용히 0건으로 보이면 사람이 "절차서가 하나도 없네"로 읽는다.
    """
    p = path or STORE
    if not os.path.exists(p):
        d = _empty()
        d["_없음"] = True                      # 아직 한 건도 안 담았다(고장이 아니다)
        return d
    try:
        with io.open(p, encoding="utf-8") as f:
            d = json.load(f)
    except Exception as e:
        d = _empty()
        d["_못읽음"] = "%s: %s" % (type(e).__name__, e)
        return d
    if not isinstance(d, dict) or not isinstance(d.get("절차서"), list):
        d = _empty()
        d["_못읽음"] = "모양이 절차서 판이 아니다"
        return d
    return d


def save(d, path=None):
    """원자적으로 갈아끼운다([171]).  반쪽 파일을 안 남긴다."""
    p = path or STORE
    os.makedirs(os.path.dirname(p), exist_ok=True)
    d = dict(d)
    for k in ("_없음", "_못읽음"):
        d.pop(k, None)
    d["갱신"] = _now()
    tmp = p + ".tmp"
    with io.open(tmp, "w", encoding="utf-8", newline="") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)
    os.replace(tmp, p)
    return p


def next_id(d):
    """다음 번호.  ★ `len()` 으로 세지 않는다 — 하나 지우면 번호가 겹친다([263])."""
    top = 0
    for r in d.get("절차서", []):
        m = _ID_RE.match(str(r.get("id") or ""))
        if m:
            top = max(top, int(m.group(1)))
    return "SOP-%03d" % (top + 1)


def normalize(row):
    """한 건을 계약된 모양으로 맞춘다.

    ★ 값을 지어내지 않는다([169]) — 없는 칸은 빈 값이지 '기타'가 아니다.
      부위를 모르는데 '기타'로 채우면 그 절차서는 **찾을 수 없는 자리**에 묻힌다.
    """
    r = {k: row.get(k) for k in 칸}
    r["제목"] = (r.get("제목") or "").strip()
    r["부위"] = (r.get("부위") or "").strip()
    r["작업종류"] = (r.get("작업종류") or "").strip()
    r["마무리"] = (r.get("마무리") or "").strip()
    r["주의사항"] = (r.get("주의사항") or "").strip()
    r["출처"] = (r.get("출처") or "").strip()
    r["작성자"] = (r.get("작성자") or "").strip()
    r["수정일"] = (r.get("수정일") or "").strip() or _now()

    공구 = r.get("공구") or []
    if isinstance(공구, str):
        공구 = [x.strip() for x in re.split(r"[,\n]", 공구) if x.strip()]
    r["공구"] = [str(x).strip() for x in 공구 if str(x).strip()]

    단계 = r.get("단계") or []
    out = []
    for i, s in enumerate(단계, 1):
        if isinstance(s, str):
            s = {"내용": s}
        if not isinstance(s, dict):
            continue
        내용 = (s.get("내용") or "").strip()
        if not 내용:
            continue                            # 빈 줄은 안 담는다(단계 번호가 헛돈다)
        out.append({"순서": i, "내용": 내용, "주의": (s.get("주의") or "").strip()})
    r["단계"] = out

    # 판(버전).  ★ 없으면 1 이다 — 0 으로 두면 "아직 안 만들어졌다"와 구별이 안 된다([169]).
    try:
        r["판"] = max(1, int(r.get("판") or 1))
    except Exception:
        r["판"] = 1
    r["원천지문"] = (r.get("원천지문") or "").strip()
    r["자동생성"] = bool(r.get("자동생성"))
    이전 = r.get("이전판") or []
    r["이전판"] = [x for x in 이전 if isinstance(x, dict)][-_판보관:]
    return r


def put(row, path=None):
    """한 건을 담거나 고친다.  id 가 없으면 새로 준다."""
    d = load(path)
    if d.get("_못읽음"):
        raise RuntimeError("절차서 판을 못 읽어 쓰지 않는다: " + d["_못읽음"])
    r = normalize(row)
    if not r["제목"]:
        raise ValueError("제목이 없다 — 제목 없는 절차서는 찾을 수 없다")
    if not r["id"]:
        r["id"] = next_id(d)
    rows = d.setdefault("절차서", [])
    for i, old in enumerate(rows):
        if old.get("id") == r["id"]:
            rows[i] = r
            break
    else:
        rows.append(r)
    save(d, path)
    return r["id"]


def _본문지문(r):
    """판을 올릴지 가르는 지문 — 내용만 본다(시계·판 자신은 뺀다)."""
    핵 = {k: r.get(k) for k in 칸 if k not in _판무관}
    글 = json.dumps(핵, ensure_ascii=False, sort_keys=True)
    return hashlib.sha1(글.encode("utf-8")).hexdigest()[:16]


def upsert_auto(row, path=None, force=False):
    """DB 에서 캔 절차서를 담는다 — **내용이 달라졌을 때만 판을 올린다**.

    2026-09-10 형님 지시: "자료가 DB에 추가되면 계속 수정해서 버전 업하는
    알고리즘 구현해."

    돌려주는 것: (갈래, id, 말)  갈래는 넷이다 —
      "새로"   처음 담았다
      "판올림" 내용이 달라져 판을 1 올렸다(옛 판은 `이전판` 에 남는다)
      "그대로" 내용이 한 톨도 안 바뀌었다 — **아무것도 안 쓴다**
      "사람것" 사람이 손댄 절차서다 — **안 덮고 그 사실만 말한다**

    ★ 사람이 적은 것을 기계가 덮지 않는다([326] — 근거의 세기가 다르다).
      현장에서 차동호 팀장이 고쳐 둔 순서를, DB 를 훑은 값이 조용히
      되돌리면 그 절차서는 다음 점검부터 틀린 순서를 말한다.
      정말 덮어야 하면 사람이 `force=True` 로 명령한다.

    ★ 안 바뀌었으면 **수정일도 안 건드린다**([170]).  건드리면 매 회차
      "갱신됨"으로 보여 진짜 갱신이 묻힌다.
    """
    d = load(path)
    if d.get("_못읽음"):
        raise RuntimeError("절차서 판을 못 읽어 쓰지 않는다: " + d["_못읽음"])

    새 = normalize(row)
    새["자동생성"] = True
    if not 새["제목"]:
        raise ValueError("제목이 없다 — 제목 없는 절차서는 찾을 수 없다")
    if not 새["id"]:
        raise ValueError("id 가 없다 — 자동 갱신은 같은 id 를 다시 찾아야 한다")

    rows = d.setdefault("절차서", [])
    옛 = None
    자리 = -1
    for i, x in enumerate(rows):
        if x.get("id") == 새["id"]:
            옛, 자리 = normalize(x), i
            break

    if 옛 is None:
        새["판"] = 1
        새["원천지문"] = _본문지문(새)
        rows.append(새)
        save(d, path)
        return ("새로", 새["id"], "판 1")

    if (not 옛.get("자동생성")) and not force:
        return ("사람것", 새["id"],
                "사람이 손댄 절차서라 안 덮었다(판 %d) — 덮으려면 force" % 옛["판"])

    if _본문지문(옛) == _본문지문(새):
        return ("그대로", 새["id"], "판 %d · 내용이 안 바뀌었다" % 옛["판"])

    # ★ 옛 판을 버리지 않는다([169]) — "그때 뭐라고 적혀 있었나"를 잃지 않는다.
    자취 = list(옛.get("이전판") or [])
    자취.append({"판": 옛["판"], "수정일": 옛.get("수정일") or "",
                 "지문": _본문지문(옛), "단계수": len(옛.get("단계") or []),
                 "제목": 옛.get("제목") or ""})
    새["판"] = 옛["판"] + 1
    새["이전판"] = 자취[-_판보관:]
    새["원천지문"] = _본문지문(새)
    새["수정일"] = _now()
    rows[자리] = 새
    save(d, path)
    return ("판올림", 새["id"], "판 %d -> %d" % (옛["판"], 새["판"]))

def remove(sop_id, path=None):
    d = load(path)
    if d.get("_못읽음"):
        raise RuntimeError("절차서 판을 못 읽어 쓰지 않는다: " + d["_못읽음"])
    rows = d.get("절차서", [])
    n = len(rows)
    d["절차서"] = [r for r in rows if r.get("id") != sop_id]
    if len(d["절차서"]) == n:
        return False
    save(d, path)
    return True


# ── 내보내기 폴더 — 형님이 지정하신다 ─────────────────────────────────
def _default_out_dir():
    """기본값.  ★ Z: 를 기본값으로 두지 않는다 — 회차가 그것을 훑고 있으면
    내보내기 한 번이 몇 분씩 선다([168])."""
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    return os.path.join(docs, "쿠팡_표준절차서")


def out_dir():
    """지금 지정된 폴더.  못 읽으면 **기본값**이지 빈 값이 아니다([169]) —
    빈 값을 주면 부르는 쪽이 현재 폴더에 파일을 쏟는다."""
    try:
        with io.open(CONFIG, encoding="utf-8") as f:
            v = (json.load(f).get("내보내기폴더") or "").strip()
        if v:
            return v
    except Exception:
        pass
    return _default_out_dir()


def set_out_dir(path):
    """폴더를 지정한다.  ★ 만들지는 않는다 — 오타 한 번이 엉뚱한 자리에
    폴더를 만든다.  실제로 쓸 때(export) 만든다."""
    path = (path or "").strip()
    if not path:
        raise ValueError("폴더 경로가 비었다")
    os.makedirs(os.path.dirname(CONFIG) or ".", exist_ok=True)
    tmp = CONFIG + ".tmp"
    with io.open(tmp, "w", encoding="utf-8", newline="") as f:
        json.dump({"내보내기폴더": path, "정한때": _now()}, f,
                  ensure_ascii=False, indent=2)
    os.replace(tmp, CONFIG)
    return path


def pick(d, ids=None, 부위=None, 작업종류=None):
    """내보낼 것을 고른다.  아무 조건이 없으면 전부."""
    rows = list(d.get("절차서", []))
    if ids:
        want = set(ids)
        rows = [r for r in rows if r.get("id") in want]
    if 부위:
        rows = [r for r in rows if (r.get("부위") or "") == 부위]
    if 작업종류:
        rows = [r for r in rows if (r.get("작업종류") or "") == 작업종류]
    # 부위 차례대로, 그 안에서 제목 가나다.  ★ 표에 없는 부위는 맨 뒤로 —
    # 조용히 빼면 그 절차서가 목록에서 사라진다([169]).
    def key(r):
        b = r.get("부위") or ""
        i = 부위목록.index(b) if b in 부위목록 else len(부위목록)
        return (i, r.get("제목") or "", r.get("id") or "")
    return sorted(rows, key=key)


def summary(path=None):
    """앱 화면이 읽는 요약.  ★ '없음'과 '못 읽음'을 가른다([169])."""
    d = load(path)
    rows = d.get("절차서", [])
    per = {}
    for r in rows:
        b = (r.get("부위") or "미지정")
        per[b] = per.get(b, 0) + 1
    return {
        "건수": len(rows),
        "부위별": per,
        "갱신": d.get("갱신", ""),
        "내보내기폴더": out_dir(),
        "부위목록": list(부위목록),
        "작업종류목록": list(작업종류목록),
        "없음": bool(d.get("_없음")),
        "못읽음": d.get("_못읽음"),
    }


def _main(argv):
    import argparse
    ap = argparse.ArgumentParser(description="표준 업무 절차서 — 담는 자리")
    ap.add_argument("--set-dir", help="내보내기 폴더를 지정한다")
    ap.add_argument("--list", action="store_true", help="담긴 절차서를 본다")
    ap.add_argument("--import-json", help="JSON 파일에서 여러 건을 담는다")
    a = ap.parse_args(argv)

    if a.set_dir:
        print("내보내기 폴더:", set_out_dir(a.set_dir))
        return 0

    if a.import_json:
        with io.open(a.import_json, encoding="utf-8") as f:
            src = json.load(f)
        rows = src if isinstance(src, list) else src.get("절차서", [])
        n = 0
        for r in rows:
            put(r)
            n += 1
        print("담음 %d건 → %s" % (n, STORE))
        return 0

    s = summary()
    if s["못읽음"]:
        print("★ 절차서 판을 못 읽었다:", s["못읽음"])
        print("  (0건이라는 뜻이 아니다 — 파일을 확인한다:", STORE, ")")
        return 1
    if s["없음"] or s["건수"] == 0:
        print("담긴 절차서 0건 — 아직 한 건도 안 들어왔다.")
        print("  차동호 팀장이 적어 주신 절차서를 받으면 앱 [표준절차서] 화면이나")
        print("  python sop_store.py --import-json <파일> 로 담는다.")
    else:
        print("절차서 %d건 · 갱신 %s" % (s["건수"], s["갱신"]))
        for b, n in sorted(s["부위별"].items()):
            print("  %-14s %d건" % (b, n))
        if a.list:
            for r in pick(load()):
                print("  [%s] %s (%s · 단계 %d)" % (
                    r.get("id"), r.get("제목"), r.get("부위") or "부위 미지정",
                    len(r.get("단계") or [])))
    print("내보내기 폴더:", s["내보내기폴더"])
    return 0


if __name__ == "__main__":
    sys.exit(_main(sys.argv[1:]))
