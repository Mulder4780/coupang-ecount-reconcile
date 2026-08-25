# -*- coding: utf-8 -*-
"""브라우저가 **지금 긁어야 할 것**을 한 곳에서 만든다 (2026-08-20 지시).

사용자 지시: "자료 긁어오는건 자료 변경이나 업로드 및 추가가 반영되면 실시간으로
백그라운드에서 대기하고 있다가 긁어오는 알고리즘 반영해"

빠져 있던 것은 '긁는 기능'이 아니다 — 그건 `grab_posts.js` 가 이미 한다. 빠진 것은
**무엇을 긁을지가 한 곳에 없다**는 것이었다. 실측 2026-08-20: 브라우저 쪽이 읽는
계획(`/api/collect_plan` → `comment_backfill.load_plan`)에는 90610953 의 **2건**만
들어 있었는데, 실제로 남은 브라우저 일은 **505건**이었다(미수집 416 · UI오염 74 ·
재수집 16). 계획을 만드는 곳이 넷이라(`recheck_plan`·`recollect`·`comment_backfill`·
UI오염 목록) 각자 제 파일로만 나갔고, **앱이 내려 주는 계획은 그중 하나뿐**이었다.
그래서 사람이 탭을 앞에 둬도 2건만 긁고 끝났다 — 오류도 안 나고 화면도 멀쩡하다([169]).

여기서 **새 판정을 만들지 않는다**([162]). 네 producer 를 그대로 불러 합치고,
거르는 것은 이미 있는 문 하나(`make_oneclick.screen`)를 그대로 쓴다([223] —
번호를 정하는 곳이 여럿이라 거르는 자리를 각자 두면 한 곳만 고쳐진다).

★ 순서가 곧 우선순위다([177]) — 그리고 **갈래마다 왜 그 갈래인지 한 줄로 말한다**.
★ 못 읽은 producer 는 **조용히 빼지 않는다**([169]) — `못읽음` 에 이름과 이유를 적는다.
  하나가 죽어도 나머지는 낸다(전부 못 내면 그날 수집이 통째로 선다).
★ 읽기 전용이다 — 캐시도 원장도 안 고치고 큐에도 안 넣는다. 쓰는 것은 제 파일 하나뿐.
"""
import io, os, sys, json, glob, argparse
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
for _p in (HERE, ROOT):
    if _p not in sys.path:
        sys.path.insert(0, _p)

QUEUE_PATH = os.path.join(ROOT, "reports", "밴드_수집대기열.json")

# 한 배치 한도는 수집기가 정한다 — 여기 숫자를 적으면 사본이 둘 된다([162]).
BATCH_MAX = 250

# 갈래 — (열쇠, 왜 이 갈래인가). 사람이 목록을 보고 "이건 왜 긁나"를 물을 수 있어야 한다([177]).
TIER_WHY = {
    "미수집": "캐시에 아예 없는 번호 — 없는 자료라 어떤 화면에도 안 뜬다",
    "오염":   "본문이 밴드 화면 UI 로 덮여 있다 — 있는데 틀린 자료다([172] 가드가 다시 받으면 되돌린다)",
    "재수집": "최근 30일 안에 **고쳐진** 글 — 개수도 날짜도 그대로라 티가 안 난다",
    "댓글":   "댓글을 한 번도 안 들여다본 글 — 취소 통보는 대부분 댓글로 온다",
}
# ★ 순서가 곧 우선순위다([177]) — 브라우저는 이 차례로 긁고, 사람이 도중에 멈춰도
#   앞쪽은 이미 들어와 있어야 한다.
# ★ **오염을 맨 뒤로 보낸다**(2026-08-25 지시: "왜 이렇게 오래 걸려").
#   실측: 대기열 795건 중 **오염이 685건(86%)** 인데 그것이 값어치 있는 110건보다
#   **앞에** 있었다. 번호당 최악 21초라 오염을 다 훑는 데 **3.9시간**이 걸리고,
#   그동안 재수집(고쳐진 글)·댓글(취소 통보)은 한 건도 안 들어온다.
#   그래서 "오래 걸리는데 맨날 잘못됐다고 한다" 가 된다.
# ★ 뒤로 보내는 근거는 짐작이 아니다 — [425] 실측으로 오염을 다시 긁어
#   **되살아난 것이 0건**이다(84789192 32번호 전부 · 90610953 237건 notime).
#   밴드가 오염 번호에는 **이웃 글 본문**을 그대로 돌려주기 때문이다.
# ⚠ **빼는 것이 아니라 미루는 것이다**([172]) — 다시 긁는 것이 되살리는 유일한
#   길이므로 오염도 여전히 내려간다. 자리만 뒤다.
TIER_ORDER = ["미수집", "재수집", "댓글", "오염"]


def _cache_bands():
    """캐시에 실재하는 밴드만. 유령 번호는 담지 않는다(2026-08-12 실사고)."""
    import convert_dump as CD          # plausible_band 는 여기 산다 — 이름을 짐작하지 않는다([165])
    out = []
    for p in sorted(glob.glob(os.path.join(HERE, "cache", "*.json"))):
        b = os.path.splitext(os.path.basename(p))[0]
        if b.isdigit() and CD.plausible_band(b):
            out.append(b)
    return out


def _dirty_nos(band, posts=None):
    """UI 오염 목록 — 근거는 **지금 캐시가 단 표시**다.

    ★ 2026-08-24 실사고(분담판 [221]): 예전에는 `reports/밴드_UI오염글_*.json`
      **만** 읽었는데 그 리포트는 한 번 만들어지고 안 갱신된다. 실측으로 8/20자
      74건만 실렸고 그것들은 **이미 되살아나** 오염 표시가 없었다 — 정작 지금
      오염인 **609건**은 어느 길에도 안 실렸고 화면에도 안 떴다([169]).
    ★ 리포트도 버리지 않는다 — 캐시에 **기록 자체가 없는** 번호는 리포트만 안다.
    """
    nos = set()
    for k, v in (posts or {}).items():
        if isinstance(v, dict) and v.get("contaminated"):
            try:
                nos.add(int(k))
            except (TypeError, ValueError):
                pass
    for p in glob.glob(os.path.join(ROOT, "reports", "밴드_UI오염글_*.json")):
        try:
            d = json.load(io.open(p, encoding="utf-8"))
        except Exception:
            continue
        b = d.get(str(band))
        if isinstance(b, dict):
            for g in (b.get("글") or []):
                try:
                    nos.add(int(g.get("번호")))
                except (TypeError, ValueError):
                    pass
    return sorted(nos)


def _tiers_for(band, posts, missed):
    """갈래별 번호. producer 하나가 죽어도 나머지는 낸다([169] — 죽은 것은 적는다)."""
    out = {k: [] for k in TIER_ORDER}

    try:
        import recheck_plan as RP
        # ★ `plan()` 의 인자 기본값은 `ahead=0` 이라 그대로 부르면
        #   **새 글 후보가 언제나 0 건**이다(2026-08-23 실측 사고).
        #   그러면 브라우저가 받는 목록에 새 글이 영영 안 실려,
        #   밴드가 며칠씩 멈춰 있어도 자동 경로로는 안 풀린다([169]).
        pl = RP.plan(band, posts, ahead=RP.default_ahead()) or {}
        out["미수집"] = sorted(set(pl.get("new") or []) | set(pl.get("gaps") or []))
    except Exception as e:
        missed.append({"갈래": "미수집", "왜": "%s: %s" % (type(e).__name__, e)})

    try:
        out["오염"] = _dirty_nos(band, posts)
    except Exception as e:
        missed.append({"갈래": "오염", "왜": "%s: %s" % (type(e).__name__, e)})

    try:
        import recollect as RC
        nos, _floor = RC.targets(band, posts)
        # ★ '이미 오늘 다시 받은 것'을 빼는 것까지가 대상이다 — 안 빼면 붙여넣기가
        #   매일 같은 것을 훑는다. 자르는 시각도 recollect 것을 그대로 빌린다([162]).
        #   ⚠ 여기를 try/except 로 감싸면 안 된다 — 실패해도 nos 가 그대로 남아
        #     **거르는 문이 조용히 없어진다**([169]). 실제로 첫 판이 그랬다:
        #     fresh_cut 을 None 으로 넘겨 TypeError 가 났는데 except 가 삼켜
        #     90610953 재수집이 1건이어야 할 자리에 109건으로 나왔다.
        from datetime import datetime as _dt, timedelta as _td
        fresh_cut = (_dt.now() - _td(hours=20)).timestamp() * 1000
        nos = RC._not_recollected(band, nos, fresh_cut)
        out["재수집"] = sorted(int(n) for n in nos)
    except Exception as e:
        missed.append({"갈래": "재수집", "왜": "%s: %s" % (type(e).__name__, e)})

    try:
        import comment_backfill as CB
        # ★ load_plan() 을 부르지 않는다 — 그 함수가 이 대기열을 합치도록 고쳤으므로
        #   여기서 부르면 **서로를 부른다**. 파일을 직접 읽어 고리를 끊는다.
        doc = json.load(io.open(CB.PLAN_PATH, encoding="utf-8"))
        b = (doc.get("bands") or {}).get(str(band)) or {}
        out["댓글"] = sorted(int(n) for n in (b.get("nos") or []))
    except Exception as e:
        missed.append({"갈래": "댓글", "왜": "%s: %s" % (type(e).__name__, e)})

    return out


def build(bands=None):
    import make_oneclick as MO
    import recheck_plan as RP
    doc = {"generated": datetime.now().isoformat(timespec="seconds"),
           "갈래설명": TIER_WHY, "bands": {}, "못읽음": []}
    for band in (bands or _cache_bands()):
        band = str(band)
        missed = []
        try:
            posts = RP.load(band) or {}
        except Exception as e:
            doc["못읽음"].append({"밴드": band, "갈래": "캐시", "왜": str(e)})
            continue
        tiers = _tiers_for(band, posts, missed)

        # ★ 거르는 자리는 하나다([223]) — 없는 번호를 사람·에이전트 손에 들려 보내지 않는다.
        seen, nos, kept, dropped_all = set(), [], {}, {}
        for t in TIER_ORDER:
            raw = [n for n in tiers[t] if n not in seen]
            keep, dropped, _why = MO.screen(band, raw, posts)
            for k, v in (dropped or {}).items():
                dropped_all.setdefault(k, 0)
                dropped_all[k] += len(v)
            kept[t] = keep
            for n in keep:
                seen.add(n)
                nos.append(n)
        doc["bands"][band] = {
            "nos": nos,                       # 전부. 자르는 것은 내려 주는 쪽이 한다.
            "tiers": {t: kept[t] for t in TIER_ORDER if kept[t]},
            "건수": {t: len(kept[t]) for t in TIER_ORDER},
            # 뺀 것은 숫자로 말한다([169]) — 조용히 빼면 '0건'이 '다 봤다'로 읽힌다.
            "거른것": dropped_all,
            "못읽음": missed,
        }
        if missed:
            doc["못읽음"].extend([dict(m, 밴드=band) for m in missed])
    return doc


def save(doc, path=None):
    path = path or QUEUE_PATH
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".%d.tmp" % os.getpid()
    with io.open(tmp, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, path)
    return path


def load(path=None):
    """읽는 쪽이 부른다. 없으면 **빈 것이 아니라 없음**을 돌려준다([169])."""
    try:
        with io.open(path or QUEUE_PATH, encoding="utf-8") as fh:
            d = json.load(fh)
        return d if isinstance(d, dict) else None
    except Exception:
        return None


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", action="store_true", help="대기열 파일을 새로 쓴다")
    ap.add_argument("--band", help="한 밴드만")
    ap.add_argument("--print", dest="show", action="store_true")
    a = ap.parse_args(argv)
    doc = build([a.band] if a.band else None)
    if a.write:
        save(doc)
    tot = 0
    for b, v in doc["bands"].items():
        tot += len(v["nos"])
        print("밴드 %s — %d건  %s" % (b, len(v["nos"]),
              " · ".join("%s %d" % (t, n) for t, n in v["건수"].items() if n)))
        if v["거른것"]:
            print("   거른 것 " + " · ".join("%s %d" % (k, n) for k, n in v["거른것"].items()))
    print("합계 %d건%s" % (tot, "" if not doc["못읽음"] else
          "  ★ 못 읽은 갈래 %d개 — 이 숫자는 전부가 아니다" % len(doc["못읽음"])))
    for m in doc["못읽음"]:
        print("   못읽음:", m)
    return 0


if __name__ == "__main__":
    sys.exit(main())
