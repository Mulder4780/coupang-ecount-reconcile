# -*- coding: utf-8 -*-
"""
camp_edit.py — **전국 쿠팡캠프 담당자를 앱에서 고친다**(추가·수정·저장)

사용자 지시(2026-08-18): **"이 화면에서 추가 입력, 수정, 변경, 저장 등의 업무처리
가능하게 고도화 해"** · **"전국 쿠팡 캠프 담당자 추가 및 변경시 자동으로 이
카테고리에 반영해서 업데이트 하는 알고리즘 추가해"**

왜 새 저장소가 필요한가 (실측)
  · 화면이 읽는 `reports/캠프_담당자.json` 은 **회차 산출물**이다 —
    09:50 `camp_contacts.py --write` 가 밴드에서 **다시 뽑아 통째로 덮는다**.
    거기에 사람이 고친 값을 써 넣으면 **다음 날 소리 없이 사라진다**([169] 모양:
    오류도 안 나고 값이 비지도 않는다. 어제 고친 번호가 옛 번호로 돌아갈 뿐이다).
  · 그러므로 정본은 2026-08-10 확정 규칙 그대로 **앱 뒤 SQLite** 다.
    `app_store.set_setting` 을 쓴다 — 한 트랜잭션에 값 + 감사로그(`change_event`),
    **멱등키**, **낙관잠금(record_version)** 이 이미 다 들어 있다. 여기서
    새 표·새 판정을 만들지 않는다([162]).

지키는 것
  · **사람이 적은 값이 이긴다.** 밴드 자동값은 사람 값이 **없는 칸에만** 채운다.
  · **'안 고침' 과 '비우기로 정함' 을 가른다**([169]). `사람값` 에 **키가 있으면**
    사람이 정한 것이다 — 빈 문자열이어도 그렇다. 키가 **없으면** 자동값을 쓴다.
    이 구별이 없으면 "자동이 틀렸으니 비워 둔다"는 결정을 표현할 길이 없다.
  · **화면이 출처를 밝힌다.** 칸마다 `자동(밴드)` 인지 `사람이 고침` 인지 적는다.
    안 적으면 왜 이 번호가 밴드와 다른지 아무도 설명하지 못한다.
  · ★ **사람이 고친 뒤 원본이 또 바뀌면 그 사실을 말한다.** 고칠 때 본 자동값의
    지문을 같이 저장해 두고, 나중에 자동값이 그것과 달라지면 화면이
    `밴드에 더 새 값이 있습니다` 라고 적는다. **자동으로 되돌리지 않는다** —
    사람이 일부러 고친 것을 기계가 덮으면 그 결정이 소리 없이 사라진다([172]).
  · **원본은 한 글자도 안 고친다** — 밴드 캐시·`캠프_담당자.json` 은 그대로 둔다.
    읽을 때만 겹쳐 놓는다(`people_alias` 와 같은 자리).

  python camp_edit.py            (사람이 고친 캠프 목록 요약)
"""
import json
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
try:  # 무인 회차는 pythonw 라 sys.stdout 이 None 이다 — [235]
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import camp_contacts as CC  # noqa: E402

KEY_PREFIX = "camp_contact:"

# 사람이 고칠 수 있는 역할. **낱말은 camp_contacts 와 같아야 한다** — 갈리면
# 저장은 되는데 화면에 영영 안 나타난다(오류는 안 난다 · [165]).
ROLES = ("현장책임", "안전관리", "담당자")
PERSON_KEYS = ("이름", "전화", "메일")
# 캠프 자체 칸. `정기점검` 은 참/거짓이라 따로 다룬다.
PLAIN_KEYS = ("캠프명", "캠프주소", "거래처코드", "비고")

FIELDS = tuple(list(PLAIN_KEYS) + ["정기점검"]
               + ["%s.%s" % (r, k) for r in ROLES for k in PERSON_KEYS])

MAXLEN = 200        # 한 칸 길이 상한 — 사람 손 입력이라 사고로 긴 값이 들어온다


def _store():
    from app_store import default_store
    return default_store()


def norm(name):
    """캠프 열쇠. **`camp_contacts._norm` 을 빌린다** — 여기서 다시 쓰면 표기가
    조금 다른 날 같은 캠프가 두 줄이 된다([162])."""
    return CC._norm(name)


def _clean(value):
    return str("" if value is None else value).strip()[:MAXLEN]


# ─────────────────────────────────────────────────────────────────────────────
# 읽기
# ─────────────────────────────────────────────────────────────────────────────
def load_edits():
    """앱 DB 에 저장된 **사람이 고친 값** 전부. {열쇠: {...}}

    ★ 못 읽으면 **빈 것이 아니라 예외**다 — 부르는 쪽이 '고친 값 없음'으로
      조용히 넘어가면 사람이 어제 고친 번호가 화면에서 사라진다([169]).
    """
    out = {}
    with _store().reader() as conn:
        for row in conn.execute(
            "SELECT key,value_json,record_version,updated_at,updated_by "
            "FROM app_setting WHERE key LIKE ? ORDER BY key",
            (KEY_PREFIX + "%",),
        ):
            key = str(row["key"])[len(KEY_PREFIX):]
            try:
                val = json.loads(row["value_json"]) or {}
            except Exception:
                continue
            if not isinstance(val, dict):
                continue
            val["_version"] = int(row["record_version"])
            val["_고친때"] = row["updated_at"]
            val["_고친이"] = row["updated_by"]
            out[key] = val
    return out


def get_edit(key):
    rec = _store().get_setting(KEY_PREFIX + key)
    val = rec.get("value")
    return ({} if not isinstance(val, dict) else val), int(rec.get("record_version") or 0)


# ─────────────────────────────────────────────────────────────────────────────
# 겹쳐 놓기 — 읽는 길 하나
# ─────────────────────────────────────────────────────────────────────────────
def _blank_row(name):
    """사람이 **새로 만든** 캠프. 밴드에 글이 아직 없으므로 건수는 전부 0 이고
    `근거` 는 비운다 — 화면이 '아직 밴드 글 없음'으로 그대로 읽는다."""
    return {
        "캠프명": name, "이름확인필요": len(name) > CC.NAME_MAX, "다른표기": [],
        "캠프주소": "", "거래처코드": "", "정기점검": False,
        "정기점검건수": 0, "돌발AS건수": 0, "총건수": 0, "최근작업일": "",
        "현장책임": {}, "안전관리": {}, "담당자": {}, "근거": {},
        "사람이만듦": True,
    }


def apply_edit(row, edit):
    """한 행에 사람 값을 덮는다. **키가 있는 칸만** 덮는다([169])."""
    vals = edit.get("사람값")
    if not isinstance(vals, dict):
        vals = {}
    origin = {}
    for field, v in vals.items():
        if field not in FIELDS:
            continue                      # 모르는 칸은 조용히 버린다(저장 때 이미 막았다)
        origin[field] = "사람"
        if field == "정기점검":
            row["정기점검"] = bool(v)
        elif "." in field:
            slot, sub = field.split(".", 1)
            person = dict(row.get(slot) or {})
            person[sub] = v
            row[slot] = person
        else:
            row[field] = v
    if "캠프명" in vals:
        row["이름확인필요"] = len(str(row.get("캠프명") or "")) > CC.NAME_MAX
    row["사람고침"] = origin
    row["사람고친때"] = edit.get("_고친때") or ""
    row["사람고친이"] = edit.get("_고친이") or ""
    # ★ 화면이 다음 저장 때 그대로 돌려보낼 **판**이다. 안 실어 보내면 화면은
    #   판을 모른 채 저장하게 되고, 그때 0(=새로 만들기)을 보내면 남의 값을
    #   덮는다([296] 이 그 모양이다). 모르면 화면이 0 대신 -1 을 보내 충돌로 끝난다.
    row["사람고친판"] = int(edit.get("_version") or 0)
    return row


def overlay(data, edits=None):
    """`camp_contacts` 산출물 위에 사람 값을 겹친다. **화면·엑셀·캡처가 이것 하나를 본다.**

    돌려주는 것은 같은 모양의 딕셔너리이며 머리글 숫자를 **다시 센다** —
    안 세면 표에는 전화가 있는데 머리글은 '모름'이라 말한다.
    """
    if edits is None:
        edits = load_edits()
    rows = []
    seen = set()
    for src in (data.get("rows") or []):
        key = norm(src.get("캠프명"))
        seen.add(key)
        edit = edits.get(key)
        if not edit:
            # 자동지문은 **모든 행**에 싣는다 — 처음 고치는 사람도 "내가 볼 때
            # 자동값이 무엇이었나"를 같이 저장해야 나중에 '밴드가 또 바뀜'을 묻는다.
            src = dict(src, 자동지문=CC.contact_sig(src))
            rows.append(src)
            continue
        # ★ **원본 행을 고치지 않는다.** 실측으로 여기서 걸렸다: 예전엔 제자리에서
        #   덮어써서, 같은 자료로 한 번 더 겹치면 `contact_sig` 가 **자동값이 아니라
        #   이미 사람 값이 덮인 행**을 재고 "밴드가 바뀌었다"는 없는 경보를 냈다.
        #   부르는 쪽이 그 자료를 또 쓰면 자동값 자체를 잃는다([172]).
        was = str(edit.get("자동지문") or "")
        now = CC.contact_sig(src)          # 겹치기 **전**이라야 자동값이다
        row = apply_edit(dict(src), edit)  # 역할 딕셔너리는 apply_edit 가 따로 복사한다
        row["자동지문"] = now
        if was and was != now:
            row["자동이바뀜"] = now
        rows.append(row)
    for key, edit in edits.items():
        if key in seen:
            continue
        name = ((edit.get("사람값") or {}).get("캠프명") or edit.get("캠프명") or key)
        rows.append(apply_edit(_blank_row(str(name)), edit))
    CC.sort_rows(rows)
    out = dict(data)
    out["rows"] = rows
    out.update(CC.summarize(rows))
    out["사람이고친캠프"] = sum(1 for r in rows if r.get("사람고침"))
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 쓰기 — 정본은 앱 SQLite 하나다(2026-08-10 확정 규칙)
# ─────────────────────────────────────────────────────────────────────────────
class 입력오류(ValueError):
    pass


def save(camp, values, *, actor, expected_version, idempotency_key=None,
         auto_sig="", clear=()):
    """한 캠프의 사람 값을 저장한다.

    · `values`  : {칸이름: 값} — `FIELDS` 밖 이름은 **거부**한다(조용히 버리지 않는다).
    · `clear`   : 사람 값을 **지워 자동값으로 되돌릴** 칸 이름들.
    · `expected_version` : 화면이 받아 간 그 판. 새로 만들 때는 0.
                  다른 기기가 먼저 고쳤으면 `VersionConflict` 가 올라간다 —
                  **덮지 않는다.**
    · `auto_sig`: 지금 화면이 보고 있던 **자동값 지문**. 나중에 밴드가 또 바뀌면
                  화면이 그 사실을 말할 수 있게 같이 저장한다.
    """
    key = norm(camp)
    if not key:
        raise 입력오류("캠프명이 없습니다")
    if not isinstance(values, dict):
        raise 입력오류("고칠 값이 없습니다")
    bad = [k for k in values if k not in FIELDS]
    if bad:
        raise 입력오류("모르는 항목입니다: " + ", ".join(sorted(bad)[:5]))
    bad = [k for k in clear if k not in FIELDS]
    if bad:
        raise 입력오류("모르는 항목입니다: " + ", ".join(sorted(bad)[:5]))

    before, ver = get_edit(key)
    if int(expected_version or 0) != ver:
        # 낙관잠금을 여기서 한 번, set_setting 이 한 번 더 본다 — 사이에 끼어들면
        # 거기서 잡힌다. 여기 검사는 **더 친절한 문구**를 주기 위한 것이다.
        from app_store import VersionConflict
        raise VersionConflict(key, int(expected_version or 0), ver)

    merged = dict(before.get("사람값") or {})
    for field, v in values.items():
        merged[field] = bool(v) if field == "정기점검" else _clean(v)
    for field in clear:
        merged.pop(field, None)          # 키를 **없애야** 자동값으로 돌아간다

    payload = {
        "사람값": merged,
        # 지문은 **고친 그 순간의 자동값**이다. 안 주면 옛 것을 그대로 둔다 —
        # 지우면 '밴드가 그 뒤 바뀌었나'를 영영 못 묻는다.
        "자동지문": _clean(auto_sig) if auto_sig else (before.get("자동지문") or ""),
        "캠프명": _clean(values.get("캠프명") or before.get("캠프명") or camp),
    }
    res = _store().set_setting(
        KEY_PREFIX + key, payload,
        expected_version=(None if ver == 0 else ver),
        actor=actor or "app", source="app-camp-edit",
        evidence="전국쿠팡캠프 화면 저장",
        idempotency_key=idempotency_key,
    )
    setting = res.get("setting") or {}
    return {"ok": True, "열쇠": key, "판": int(setting.get("record_version") or 1),
            "고친때": setting.get("updated_at"), "칸수": len(merged)}


def main():
    edits = load_edits()
    print("사람이 고친 캠프 %d개" % len(edits))
    for key, e in list(edits.items())[:10]:
        vals = e.get("사람값") or {}
        print("  %-20s 판%-3s %s"
              % (str(e.get("캠프명") or key)[:20], e.get("_version"),
                 ", ".join(sorted(vals)[:4])))


if __name__ == "__main__":
    main()
