# -*- coding: utf-8 -*-
"""사람이 바뀌어도 기록은 이어진다 — 이름·전화번호 인계 사전 (2026-08-08 지시).

사용자 지시: **"김혜진 매니저 퇴사한지 1달이 넘었어. 카톡에 이렇게 올라와도
류지영으로 인식하는 알고리즘 추가. 번호는 같아"**

밴드·카톡 글에는 접수 안내 문구가 **양식으로 박혀** 있다:

    ☎ 쿠팡 A/S 접수부서 및 담당
    담당자님, 향후 AS접수시 아래 쿠팡 AS 전담부서~
    김혜진 매니저님께 접수 부탁드리겠습니다.
    [담당이름] 김혜진 매니저
    [휴대전화] 010-6645-4535

사람이 그만두어도 **양식은 그대로 돈다.** 2026-08-08 실측에서 밴드 글 871건이
아직 이 이름을 담고 있었고, 그중 **35건은 `A/S 담당` 칸 자체가 그 이름**이었다.
그대로 두면 없는 사람이 담당기사로 집계되고, "누구한테 접수했나"를 되짚을 때
현재 담당자에게 닿지 못한다.

★ **원문은 고치지 않는다.** 캐시·원본은 쓰인 그대로 두고, **읽을 때** 현재 담당으로
  옮긴다. 원문을 고치면 "그때 정말 누구 앞으로 왔나"를 잃는다 — 인계 이전 글까지
  소급해 바꿔 버리기 때문이다. 그래서 `since`(바뀐 때)를 함께 둔다.

★ 전화번호가 같다는 것이 근거다. 이름은 양식에 박혀 안 바뀌지만 번호는 실제 인계
  대상이라, 번호가 같으면 그 자리를 물려받은 것이 확실하다.
"""
import re

# 인계 사전. 새 인계가 생기면 여기에 한 줄 더한다 — 코드는 안 고친다.
HANDOVERS = (
    {
        "before": "김혜진",
        "after": "류지영",
        "role": "쿠팡 A/S 접수 담당",
        "since": "2026-07-01",          # 퇴사 시점(사용자 확인: 2026-08-08 기준 한 달 넘음)
        "phones": ("010-6645-4535",),
        "why": "김혜진 매니저 퇴사. 접수 번호는 그대로 인계돼 같은 자리를 이어받았다.",
    },
)

_DIGITS = re.compile(r"\D+")
# 문서에 [담당이름] 김혜진 매니저 / A/S 담당 : 김혜진 처럼 직책이 붙어 온다
_TITLE_TAIL = re.compile(r"\s*(매니저|과장|차장|부장|대리|팀장|소장|실장|반장|기사|기장)\s*(님)?$")


def _phone_key(s):
    """전화번호를 숫자만 남겨 비교한다 — 010-6645-4535 · 01066454535 · +82 10 … 을 같게."""
    d = _DIGITS.sub("", str(s or ""))
    if d.startswith("82"):
        d = "0" + d[2:]
    return d


_PHONE_INDEX = {}
for _h in HANDOVERS:
    for _p in _h.get("phones") or ():
        _PHONE_INDEX[_phone_key(_p)] = _h
_NAME_INDEX = {h["before"]: h for h in HANDOVERS}


def strip_title(name):
    """'김혜진 매니저' · '김혜진 매니저님' → '김혜진'."""
    return _TITLE_TAIL.sub("", str(name or "").strip()).strip()


def handover_of(name="", phone=""):
    """이 이름/번호가 인계된 자리인가 — 맞으면 인계 항목, 아니면 None.

    번호를 먼저 본다. 이름은 동명이인이 있을 수 있지만 번호는 그 자리 자체다.
    """
    h = _PHONE_INDEX.get(_phone_key(phone)) if phone else None
    return h or _NAME_INDEX.get(strip_title(name))


def resolve_person(name="", phone="", when=""):
    """지금 이 일을 맡고 있는 사람 이름을 돌려준다(모르면 받은 이름 그대로).

    `when` 은 그 글이 쓰인 날(YYYY-MM-DD). 인계 **전** 글이면 옮기지 않는다 —
    그때는 정말 그 사람이 맡고 있었다. 날짜를 모르면 옮긴다(양식 문구가 그대로
    도는 것이 실제 사례이므로, 모를 때는 현재 담당에 닿는 쪽이 쓸모 있다).
    """
    h = handover_of(name, phone)
    if not h:
        return strip_title(name)
    day = str(when or "")[:10]
    if day and day < h["since"]:
        return strip_title(name)
    return h["after"]


def resolve_text(text, when=""):
    """글 본문에서 인계된 이름을 찾아 **현재 담당자**를 돌려준다(없으면 "").

    번호가 본문에 있으면 그것을 근거로 삼는다 — 이름 철자가 틀려도 잡힌다.
    """
    s = str(text or "")
    if not s:
        return ""
    for h in HANDOVERS:
        for p in h.get("phones") or ():
            if _phone_key(p) and _phone_key(p) in _phone_key(s):
                return resolve_person(h["before"], p, when)
    for h in HANDOVERS:
        if h["before"] in s:
            return resolve_person(h["before"], "", when)
    return ""


def note_of(name="", phone=""):
    """왜 이름이 바뀌었는지 한 줄 — 화면·리포트에 근거로 남긴다."""
    h = handover_of(name, phone)
    if not h:
        return ""
    return (f"{h['before']} → {h['after']} ({h['role']}, {h['since']} 인계) · {h['why']}")


def main(argv=None):
    import sys
    args = list(sys.argv[1:] if argv is None else argv)
    if args:
        for a in args:
            print(f"{a} → {resolve_person(a, a)}   {note_of(a, a)}")
        return 0
    print("이름·번호 인계 사전")
    for h in HANDOVERS:
        print(f"  {h['before']} → {h['after']}  ({h['role']}, {h['since']}부터)")
        print(f"     번호 {', '.join(h.get('phones') or ()) or '-'} · {h['why']}")
    print("\n확인:  python people_alias.py 김혜진 010-6645-4535")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
