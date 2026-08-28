# -*- coding: utf-8 -*-
"""정기점검 **예정**(스케줄 원본에서 온 것)을 목록에서 숨긴다 — 지우지 않는다.

★ **왜 '삭제'가 아니라 '숨김'인가** (2026-08-28 류지영 매니저 요청:
  *"이건은 아직 앱에 등록되지않은 예정입니다 … 이것도 삭제가능하게 해주세여"*).
  이 예정의 원천은 **류지영 정기점검 스케줄 원본**이고 관리대장 `27_정기점검원본일정`
  시트를 거쳐 화면에 온다. 앱 DB(`work_item`)에는 **행이 아예 없다** — 그래서
  `soft_delete_work` 가 못 돈다(그 함수는 있는 행의 `deleted_at` 을 세운다).
  게다가 여기서 원본을 지워도 **다음 회차가 도로 만든다** — 지우는 시늉만 하고
  다음 날 되살아나면 사람은 앱을 못 믿는다([355] 가 캠프에서 배운 그대로다).

★ **그래서 숨김 표시를 앱 DB 에 남기고 읽는 쪽이 거른다.** 원본은 한 글자도
  안 건드린다(2026-08-10 정본 규칙 · 역수입 금지).

★ **사유 없이는 못 숨긴다.** 몇 달 뒤 "이 점검 왜 없지"를 물을 사람이 반드시
  있고, 그때 답할 수 있는 것은 그 한 줄뿐이다([355] 와 같은 문).

★ **되살릴 수 있다.** 삭제만 있고 되살리기가 없으면 잘못 지운 순간 사람이 할 수
  있는 일이 없어진다(`app_store.restore_work` 독스트링이 적어 둔 그대로다).

⚠ **`app_setting` 표를 빌린다**([162]) — 낙관잠금·감사로그·멱등키가 공짜로 붙는다.
  새 표를 만들면 그 셋을 여기서 다시 짜야 하고, 짠 것은 언젠가 갈린다.

실측 2026-08-28(v622): `get_works()['pm']` 469건 중 앱 DB 에 행이 없는 것 **73건** —
그중 `SCH-` **71건**이 이 갈래다. 나머지 2건은 ERP 에서 온 **완료** 실적이라
대상이 아니다(ID 도 없어 열쇠를 못 만든다 · 손대지 않았다 · [172]).
`as` 는 **610건 전부** 앱 DB 에 있어 예전부터 삭제가 된다.
"""
from __future__ import annotations

import json
import re
import sys
from datetime import datetime

KEY_PREFIX = "pm_plan_hide:"

# 열쇠 모양 — `pm_schedule_sync._plan_id()` 가 만든다: SCH-<연도>Q<분기>-<sha1 앞10 대문자>
# ⚠ 여기서 새 모양을 지어내지 않는다([165]) — 어긋나면 한 건도 안 걸리면서 오류도 안 난다.
ID_RE = re.compile(r"^SCH-\d{4}Q[1-4]-[0-9A-F]{10}$")

REASON_MIN = 2
REASON_MAX = 200


class 입력오류(ValueError):
    pass


def _store():
    from app_store import default_store
    return default_store()


def _clean(v):
    return str("" if v is None else v).strip()[:REASON_MAX]


def valid_id(plan_id) -> bool:
    """이 열쇠가 **스케줄 원본에서 온 예정**의 것인가.

    ★ 모양을 검사하는 이유는 취향이 아니다. 여기에 `PM-`·`UJ...` 가 들어오면
      그 건은 앱 DB 에 행이 있어 **`soft_delete_work` 로 지워야 하는 것**인데,
      숨김으로 처리하면 감사로그도 되살리기도 그쪽 규칙을 못 탄다.
      곧 **두 갈래가 섞이면 어느 쪽 규칙이 도는지 아무도 모른다**.
    """
    return bool(ID_RE.match(str(plan_id or "").strip().upper()))


# ─────────────────────────────────────────────────────────────────────────────
# 읽기
# ─────────────────────────────────────────────────────────────────────────────
def hidden_map():
    """숨긴 예정 전부. `({일정ID: {때,누가,왜}}, 못읽은이유)`.

    ★ **못 읽어도 화면을 죽이지 않는다** — 이것 하나 때문에 정기점검 목록이
      통째로 안 뜨면 고치려던 것보다 나쁘다([172]).
    ★ **그러나 조용히 넘어가지도 않는다**([169]). 못 읽었으면 그 이유를 같이
      돌려주고 부르는 쪽이 화면에 적는다 — 안 적으면 숨긴 예정이 **말없이
      되살아나** 사람이 두 번 지운다.
    """
    out = {}
    try:
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
                if not isinstance(val, dict) or not val.get("숨김"):
                    continue
                rec = dict(val["숨김"])
                rec["_version"] = int(row["record_version"])
                rec["캠프명"] = str(val.get("캠프명") or "")
                rec["점검예정일"] = str(val.get("점검예정일") or "")
                out[key] = rec
    except Exception as exc:
        return {}, "%s: %s" % (type(exc).__name__, exc)
    return out, ""


def hidden_ids():
    """숨긴 열쇠 집합만 — 거르는 자리에서 쓴다. 못 읽으면 **빈 집합**이다."""
    m, _why = hidden_map()
    return set(m)


def get_one(plan_id):
    """`(그 기록, 판)`. 없으면 `({}, 0)`."""
    key = str(plan_id or "").strip().upper()
    rec = _store().get_setting(KEY_PREFIX + key)
    val = rec.get("value")
    return ({} if not isinstance(val, dict) else val), int(rec.get("record_version") or 0)


# ─────────────────────────────────────────────────────────────────────────────
# 쓰기
# ─────────────────────────────────────────────────────────────────────────────
def set_hidden(plan_id, hidden, *, actor, reason="", expected_version=None,
               idempotency_key=None, camp="", plan_date=""):
    """예정 하나를 숨기거나(`hidden=True`) 되살린다(`False`).

    · `reason` : 숨길 때 **필수**(2자 이상). 되살릴 때는 안 받는다.
    · `expected_version` : 화면이 받아 간 판. 처음이면 0/None.
      다른 기기가 먼저 고쳤으면 `VersionConflict` 가 올라간다 — **덮지 않는다.**
    · `camp`·`plan_date` : 숨긴 목록을 사람이 읽을 수 있게 같이 적어 둔다.
      ⚠ 이것은 **표시용 사본**이다 — 판정에 쓰지 않는다(원본이 바뀌면 낡는다).
    """
    key = str(plan_id or "").strip().upper()
    if not valid_id(key):
        raise 입력오류(
            "스케줄 원본에서 온 예정이 아닙니다(%s) — 앱에 등록된 기록은 "
            "[삭제] 로 지웁니다" % (key[:40] or "빈 값"))
    if hidden:
        why = _clean(reason)
        if len(why) < REASON_MIN:
            raise 입력오류("지우는 이유를 적어 주세요 (예: 이 캠프는 점검 대상이 아님)")
    else:
        why = ""

    before, ver = get_one(key)
    if int(expected_version or 0) != ver:
        from app_store import VersionConflict
        raise VersionConflict(key, int(expected_version or 0), ver)

    payload = {
        "캠프명": _clean(camp) or _clean(before.get("캠프명")),
        "점검예정일": _clean(plan_date) or _clean(before.get("점검예정일")),
    }
    if hidden:
        payload["숨김"] = {
            "때": datetime.now().isoformat(timespec="seconds"),
            "누가": _clean(actor or "app"),
            "왜": why,
        }
    # 되살릴 때는 `숨김` 키를 **없앤다** — 빈 dict 를 남기면 읽는 쪽이 갈릴 수 있다.

    res = _store().set_setting(
        KEY_PREFIX + key, payload,
        expected_version=(None if ver == 0 else ver),
        actor=actor or "app", source="app-pm-plan-hide",
        evidence=("정기점검 예정 숨김: " + why) if hidden else "정기점검 예정 되살림",
        idempotency_key=idempotency_key,
    )
    setting = res.get("setting") or {}
    return {
        "ok": True, "열쇠": key,
        "판": int(setting.get("record_version") or 1),
        "고친때": setting.get("updated_at"),
        "숨김": bool(hidden),
        "msg": ("목록에서 지웠습니다 — [지운 예정 보기]에서 되살릴 수 있습니다"
                if hidden else "되살렸습니다"),
    }


# ─────────────────────────────────────────────────────────────────────────────
def main():
    m, why = hidden_map()
    if why:
        print("숨긴 예정을 못 읽었습니다 — %s" % why)
        return 1
    print("목록에서 지운 정기점검 예정 %d건" % len(m))
    for key, rec in sorted(m.items(), key=lambda kv: str(kv[1].get("때") or "")):
        print("  %-24s %-18s %-10s  %s  (%s)"
              % (key, str(rec.get("캠프명") or "")[:18],
                 str(rec.get("점검예정일") or ""),
                 str(rec.get("왜") or "")[:40], str(rec.get("누가") or "")))
    return 0


if __name__ == "__main__":
    sys.exit(main())
