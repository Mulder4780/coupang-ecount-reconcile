# -*- coding: utf-8 -*-
"""
recollect.py — **매일 08:00 재수집 회차**: 최근 30일 글만 다시 받아 바뀐 것을 찾아낸다.

사용자 지시(2026-08-08): "매일 08:00 회차에 최근 30일(예: 30일) 글만 재수집 대상으로
뽑아 → 캐시에 다시 넣고 → put_record가 바뀐 것만 record_rev에 남기고 → 바뀐 게 있으면
인계 문서 맨 위에 올린다."

무엇을 푸는 문제인가
  밴드 글은 **수정된다**(2026-08-04 확인). 상태가 바뀌면 같은 번호의 글을 고쳐 다시
  올린다 — 미실시가 완료로, 금액이 바뀌고, 사진이 붙는다. 그런데 지금까지의 수집은
  전부 **'없는 것을 채우는'** 방향이었다: recheck_plan 은 구멍·새 글·옛 수집분을
  고르고, band_sync 는 이미 아는 글을 만나면 그 자리에서 멈춘다.
  → **이미 받은 최근 글이 그 뒤에 고쳐지면 아무도 다시 안 본다.** 캐시는 첫 수집
    당시의 모습으로 굳고, 화면 숫자는 멀쩡한데 원본과 다르다. 밀림보다 조용한 사고다.
  이 회차가 그 한 갈래를 맡는다 — **최근 30일은 매일 다시 본다.**

왜 30일인가
  글이 고쳐지는 것은 대부분 그 달 안이다(완료 보고·금액 확정). 전량을 매일 다시
  긁으면 5,000글 × 5초 = 7시간이라 애초에 회차가 될 수 없다. 창은 `--days` 로 바꾼다.

네 단계 (그리고 **어디까지가 무인인가**)
  ① 대상 뽑기      — 캐시에서 최근 N일 글 번호. 무인.
  ② 캐시에 다시 넣기 — 떨어져 있는 덤프를 흡수(무인) + 아직 안 긁었으면 붙여넣기 파일 준비.
                      ★ 긁기 자체는 **사람 로그인**이 있어야 한다(절대규칙 3). 그래서
                        이 단계는 '한다/못 한다'가 아니라 **'사람 손 한 번을 남긴다'**이다.
                        여기서 거짓으로 성공을 적으면 ③④가 옛 캐시를 보고 "바뀐 것 없음"을
                        내놓는다 — 아무 일도 안 일어났는데 안심시키는, 가장 나쁜 결과다.
  ③ 바뀐 것만 기록  — datalake.ingest_band(since=기준일). put_record 가 해시로 판정해
                      **달라진 것만** record_rev 에 남긴다(같으면 아무것도 안 쓴다). 무인.
  ④ 인계 문서 맨 위  — reports/밴드_재수집.json 을 session_handoff 가 읽어 올린다. 무인.

  python band/recollect.py --run              # 08:00 회차 전체(스케줄러가 부르는 것)
  python band/recollect.py --plan             # 대상만 보기(아무것도 안 바꾼다)
  python band/recollect.py --run --days 60    # 창 넓히기
  python band/recollect.py --print            # 지난 회차 결과
  python band/recollect.py --ack              # 인계 문서 맨 위 배너 내리기(확인했다)

검증 [152].
"""
import argparse
import json
import os
import subprocess
import sys
from datetime import datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
sys.path.insert(0, HERE)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

DAYS = 30              # 사용자가 정한 창(2026-08-08). --days 로 바꾼다.
LIMIT = 400            # 한 회차 상한 — 글당 5초라 400이면 붙여넣기 33분쯤
PY = sys.executable


def _shared(*parts):
    """워크트리에서도 **본체 하나**를 본다(CLAUDE.md 워크트리 규칙)."""
    try:
        from worktree_state import shared
        return shared(*parts)
    except Exception:
        return os.path.join(ROOT, *parts)


STATE = _shared("reports", "밴드_재수집.json")


def load_state():
    try:
        return json.load(open(STATE, encoding="utf-8"))
    except Exception:
        return {}


def save_state(st):
    os.makedirs(os.path.dirname(STATE), exist_ok=True)
    json.dump(st, open(STATE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


# ── ① 대상 뽑기 ──────────────────────────────────────────────────────────────
def targets(band, posts, days=DAYS, today=None):
    """최근 `days` 일에 **작성된** 글 번호. 최신부터.

    ★ 거르는 것들은 recheck_plan 과 같은 근거를 쓴다 — 여기서만 다르게 판단하면
      한쪽은 긁으라 하고 다른 쪽은 긁지 말라 해서 사람이 무엇을 믿을지 모르게 된다.
        · created_at 없음  → 시각 없는 수확은 진짜 글이라는 증거가 없다(2026-08-07 사고)
        · deleted          → 열면 밴드 홈이 돌아온다. 영원히 실패한다
        · contaminated     → 남의 본문이 잡힌 기록. 삭제 판정(검증 [135])
        · absent           → 처음부터 없던 번호(유령). 긁는 행위 자체가 오염을 만든다
    """
    import datalake as D
    base = (datetime.strptime(str(today), "%Y-%m-%d") if today else datetime.now())
    floor_day = (base - timedelta(days=int(days))).strftime("%Y-%m-%d")
    out = []
    for no, p in (posts or {}).items():
        if not str(no).isdigit() or not isinstance(p, dict):
            continue
        if p.get("deleted") or p.get("contaminated") or p.get("absent"):
            continue
        day = D.band_day(p.get("created_at"))
        if day and day >= floor_day:
            out.append(int(no))
    return sorted(out, reverse=True), floor_day


def plan(days=DAYS, limit=LIMIT, today=None):
    """밴드별 대상. 캐시만 읽는다 — 아무것도 바꾸지 않는다."""
    import recheck_plan as RP
    bands = sorted(f[:-5] for f in os.listdir(RP.CACHE)
                   if f.endswith(".json") and f[:-5].isdigit())
    out, floor_day = {}, ""
    for band in bands:
        posts = RP.load(band)
        if posts is None:
            continue
        nos, floor_day = targets(band, posts, days, today)
        out[band] = {"수": len(nos), "번호": nos[:limit],
                     "넘침": max(0, len(nos) - limit)}
    return out, floor_day


# ── ② 캐시에 다시 넣기 ───────────────────────────────────────────────────────
def _run(args, timeout=1800):
    try:
        r = subprocess.run([PY] + args, cwd=ROOT, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=timeout,
                           env={**os.environ, "PYTHONIOENCODING": "utf-8"}, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        tail = [l for l in (r.stdout or "").splitlines() if l.strip()]
        return r.returncode == 0, (tail[-1] if tail else "")
    except Exception as e:
        return False, str(e)[:200]


def absorb():
    """**이미 떨어져 있는** 덤프를 캐시로 흡수한다. 여기서 밴드를 긁지는 않는다.

    사람이 붙여넣기를 돌렸다면 dump_*.json 이 다운로드 폴더에 있다. 그것을 Z: 로
    옮기고(download_intake) 캐시에 합친다(convert_dump). 덤프가 없으면 두 단계 모두
    할 일이 없어 곧 끝난다 — 그래서 조건 없이 매 회차 돌려도 싸다.
    """
    steps = []
    ok1, m1 = _run([os.path.join(ROOT, "download_intake.py"), "--apply"], 600)
    steps.append({"단계": "다운로드 흡수", "ok": ok1, "말": m1})
    ok2, m2 = _run([os.path.join(HERE, "convert_dump.py")], 900)
    steps.append({"단계": "덤프 → 캐시", "ok": ok2, "말": m2})
    return steps


def make_paste(band, nos, days):
    """사람 손 **한 번**으로 끝나게 붙여넣기 파일을 만든다.

    ★ 재수집은 `{keep:false}` 회차 JS 를 그대로 쓴다 — make_oneclick 과 갈래를
      나누지 않는다. 나누면 한쪽만 고쳐지고 다른 쪽은 조용히 옛 규칙으로 남는다.
    """
    import make_oneclick as MO
    js, note = MO.build(band, len(nos), nos=nos, why=f"최근 {days}일 재수집")
    if not js:
        return None, note
    path = os.path.join(HERE, f"재수집_붙여넣기_{band}.js")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(js)
    return path, note


# ── ③ 바뀐 것만 record_rev 로 ───────────────────────────────────────────────
def ingest(floor_day, started):
    """최근 창만 흡수. put_record 가 **달라진 것만** record_rev 에 남긴다.

    ★ 창을 거는 이유는 속도만이 아니다. 전량을 흡수하면 창 밖의 오래된 글도 같이
      비교돼, 이번 회차가 무엇을 발견했는지가 흐려진다. 재수집 회차의 답은
      "**최근 30일에서** 무엇이 바뀌었나" 한 문장이어야 한다.
    """
    import datalake as D
    con = D.connect()
    try:
        got = D.ingest_band(con, quiet=True, since=floor_day, why="재수집 회차")
        # 이번 회차에 쌓인 변경 이력만 — 시각 기준이라 앞 회차 것이 섞이지 않는다.
        got["변경상세"] = [
            {"언제": r["at"], "글": r["natural_key"], "작성일": r["biz_date"],
             "어떻게": r["어떻게"], "왜": r["why"]}
            for r in D.record_changes(con, kind="band_post", at_since=started, limit=200)
        ]
        return got
    finally:
        con.close()


# ── 회차 ─────────────────────────────────────────────────────────────────────
def run(days=DAYS, limit=LIMIT, today=None, do_absorb=True):
    started = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st = {"회차": started, "창일수": int(days), "손볼것": []}

    steps = absorb() if do_absorb else []
    st["흡수단계"] = steps

    tg, floor_day = plan(days, limit, today)
    st["기준일"] = floor_day
    got = ingest(floor_day, started)
    st["흡수"] = {k: got[k] for k in ("신규", "변경", "그대로", "버림")}
    st["새글"] = got["새글"]
    st["바뀐글"] = got["바뀐글"]
    st["변경상세"] = got["변경상세"]

    # 붙여넣기 파일은 **아직 안 긁은 밴드에만** 만든다. 오늘 이미 재수집된 밴드에
    # 또 파일을 만들어 두면 사람이 매일 같은 것을 붙여넣게 된다.
    fresh_cut = (datetime.now() - timedelta(hours=20)).timestamp() * 1000
    st["대상"] = {}
    for band, info in tg.items():
        nos = info["번호"]
        need = _not_recollected(band, nos, fresh_cut)
        rec = {"창안글": info["수"], "다시받을것": len(need), "넘침": info["넘침"]}
        if need:
            path, note = make_paste(band, need, days)
            rec["붙여넣기"] = path
            rec["안내"] = note
            st["손볼것"].append(
                "밴드 %s — 최근 %d일 중 %d건이 오늘 아직 안 받아졌다."
                " 로그인된 밴드 탭 콘솔(F12)에 `%s` 를 붙여넣으면 나머지는 자동이다."
                % (band, days, len(need), os.path.relpath(path, ROOT) if path else "?"))
        st["대상"][band] = rec

    # ★ 바뀐 것을 **다음 회차가 지우지 않게** 따로 보관한다. 인계 문서 배너는 사람이
    #   `--ack` 로 내리기 전까지 남아야 한다 — 아침에 뜬 경보가 다음 아침에 조용히
    #   사라지면, 아무도 안 본 채로 없어진 것과 같다.
    # ★ 다만 **새 글만으로는 배너를 올리지 않는다.** 새 글이 들어오는 것은 정상이고
    #   이미 밀림 표가 세고 있다. 그것으로 맨 위 칸을 켜면 매일 아침 켜져 있어서,
    #   정작 글이 고쳐진 날에 아무도 눈길을 안 준다. 배너는 **고쳐진 글** 전용이다.
    if st["바뀐글"]:
        _bands = sorted({str(c.get("글","")).partition("/")[0]
                         for c in (st["변경상세"] or [])} - {""})
        st["되돌아감"] = regressed(st["변경상세"], _cache_posts(_bands))
        st["최근변경"] = {"회차": started, "바뀐글": st["바뀐글"], "새글": st["새글"],
                          "되돌아감": st["되돌아감"],
                          "변경상세": st["변경상세"]}
        st["확인함"] = False
    else:
        old = load_state()
        if old.get("최근변경") and not old.get("확인함"):
            st["최근변경"] = old["최근변경"]     # 아직 아무도 안 봤다 — 그대로 둔다
            st["확인함"] = False
    save_state(st)
    _log(st)
    return st



# ── 되돌아감 ────────────────────────────────────────────────────────────────
def regressed(changes, band_posts=None):
    """완료였는데 지금은 완료가 아닌 글만 고른다 (2026-09-01 실사고).

    ★ 이 회차가 여태 할 수 있던 말은 **`본문 바뀜`** 뿐이라, **완료 -> 안내로
      되돌아간 것**과 **안내 -> 완료로 나아간 것**이 한 덩어리로 보였다.  앞은
      사고이고 뒤는 정상이다.  실측 2026-09-01: 바뀐 글 12건 중 **11건이
      되돌아감**이었고 그 프로젝트의 완료 글은 캐시 어디에도 없었다 —
      곧 밴드 쪽 완료 근거가 통째로 사라졌는데 화면은 `본문 바뀜` 이라고만 했다.
    ★ **판정을 새로 만들지 않는다**([162]) — 완료/안내는 `band_extract.parse_post`
      하나가 정한다.  여기서 낱말을 다시 적으면 그 글자가 바뀌는 날 **한 건도
      안 걸리면서 오류도 안 난다**([165]).
    ★ **옛 요약은 잘린 글자다** — 그래도 양식 머리(`♣ ［ … 완료 ]`)는 앞쪽에 있어
      살아 있다.  못 읽으면 **되돌아감이라 우기지 않는다**([169]).
    """
    try:
        import band_extract as BE
    except Exception:
        return None                      # 못 재면 '없다'가 아니라 '모름'이다([169])

    def _state(text, no, band):
        try:
            r = BE.parse_post(no, {"content": text or ""}, band) or {}
        except Exception:
            return None
        return r.get("진행상태") or None

    out = []
    for c in changes or []:
        key = str(c.get("글") or "")
        band, _, no = key.partition("/")
        how = str(c.get("어떻게") or "")
        if " -> " in how:
            old_s, new_s = how.split(" -> ", 1)
        elif "→" in how:
            old_s, new_s = how.split("→", 1)
        else:
            continue                     # 옛/새를 못 가른다 — 지어내지 않는다
        was = _state(old_s, no, band)
        now = None
        if band_posts is not None:
            p = (band_posts.get(band) or {}).get(no)
            if p is not None:
                now = _state(p.get("content") or "", no, band)
        if now is None:
            now = _state(new_s, no, band)
        if was == "작업완료" and now and now != "작업완료":
            out.append({"글": key, "작성일": c.get("작성일"),
                        "was": was, "now": now})
    return out


def _ensure_regressed(d):
    """자국에 `되돌아감` 이 없으면 **그 자리에서 계산**한다.

    ★ 옛 자국은 그 칸을 **안 물었을 뿐**이지 '되돌아감 없음'이 아니다([247]).
      그렇다고 매번 "확인 못 했다"만 말하면 회차를 다시 돌기 전까지 답이 없다.
    ★ **바뀐 글이 있을 때만** 캐시를 읽는다([168]) — 대개 없고, `--ack` 뒤에는
      아예 안 불린다.  Z: 는 한 번도 안 문다(밴드 캐시는 로컬이다).
    """
    if not isinstance(d, dict):
        return d
    if d.get("되돌아감") is not None:
        return d
    ch = d.get("변경상세") or []
    if not ch:
        return d
    bands = sorted({str(c.get("글", "")).partition("/")[0] for c in ch} - {""})
    d["되돌아감"] = regressed(ch, _cache_posts(bands))
    return d

def _cache_posts(bands):
    """되돌아감 판정에 쓸 **지금 캐시 본문**.  못 읽으면 그 밴드는 건너뛴다."""
    import recheck_plan as RP
    out = {}
    for b in bands:
        try:
            out[b] = RP.load(b) or {}
        except Exception:
            pass
    return out

def _not_recollected(band, nos, fresh_cut_ms):
    """이 번호들 중 **최근에 다시 받지 않은** 것. 판정 근거는 캐시의 captured_at.

    오늘 이미 받은 글을 또 목록에 넣으면 붙여넣기가 매일 같은 400건을 훑는다.
    """
    import recheck_plan as RP
    posts = RP.load(band) or {}
    out = []
    for n in nos:
        p = posts.get(str(n)) or {}
        try:
            cap = int(p.get("captured_at") or 0)
        except (TypeError, ValueError):
            cap = 0
        if cap < fresh_cut_ms:
            out.append(n)
    return out


def _log(st):
    """보관소 로그에도 한 줄 — 실패도 남긴다. DB 가 잠겨도 회차를 죽이지 않는다."""
    try:
        import datalake as D
        con = D.connect()
        try:
            D.log(con, "band", "recollect", ok=True,
                  detail={"창": st.get("창일수"), "기준일": st.get("기준일"),
                          "바뀜": len(st.get("바뀐글") or []),
                          "새글": len(st.get("새글") or []),
                          "손볼것": len(st.get("손볼것") or [])})
            con.commit()
        finally:
            con.close()
    except Exception:
        pass


def ack():
    st = load_state()
    st["확인함"] = True
    st["확인시각"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_state(st)
    return st


def banner():
    """인계 문서가 맨 위에 올릴 것. 볼 것이 없으면 None.

    session_handoff 가 이 한 함수만 부른다 — 판정 기준이 두 곳에 갈리지 않게.
    """
    st = load_state()
    if st.get("확인함"):
        return None
    rc = st.get("최근변경") or {}
    if not rc.get("바뀐글"):
        return None
    _ensure_regressed(rc)
    return {"회차": rc.get("회차") or st.get("회차", ""),
            "창일수": st.get("창일수", DAYS),
            "바뀐글": rc.get("바뀐글") or [], "새글": rc.get("새글") or [],
            "변경상세": rc.get("변경상세") or [],
            "되돌아감": rc.get("되돌아감") or []}


def show(st):
    print(f"밴드 재수집 회차 {st.get('회차','')} · 최근 {st.get('창일수')}일"
          f"(기준일 {st.get('기준일','?')} 이후)")
    for band, r in sorted((st.get("대상") or {}).items()):
        print(f"  밴드 {band}: 창 안 {r['창안글']}건 · 다시 받을 것 {r['다시받을것']}"
              + (f" · 상한 넘침 {r['넘침']}" if r.get("넘침") else "")
              + (f"\n      → {os.path.relpath(r['붙여넣기'], ROOT)}"
                 if r.get("붙여넣기") else ""))
    g = st.get("흡수") or {}
    print(f"  흡수: 새 글 {g.get('신규',0)} · **바뀐 글 {g.get('변경',0)}**"
          f" · 그대로 {g.get('그대로',0)}")
    _ensure_regressed(st)
    _rg = st.get("되돌아감")
    if _rg:
        print("  ★ 그중 **완료 -> 완료 아님으로 되돌아간 글 %d건**"
              " — 밴드 쪽 완료 근거가 사라졌다: %s"
              % (len(_rg), ", ".join(r["글"] for r in _rg[:8])))
    elif _rg is None and st.get("바뀐글"):
        print("  ※ 되돌아감 여부는 **확인 못 했다**(판정을 못 불렀다)")
    for c in (st.get("변경상세") or [])[:15]:
        # 어떻게 가 이제 무엇이 바뀜는지까지 담는다 — 콘솔은 줄이되
        # **자른 것은 말한다**([273]).  JSON · 인계에는 온전히 간다.
        _how = str(c["어떻게"])
        if len(_how) > 150:
            _how = _how[:150] + "…(%d자)" % len(_how)
        print(f"    · {c['글']} ({c['작성일']}) {_how}")
    for h in st.get("손볼것") or []:
        print(f"  ※ {h}")


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--run", action="store_true", help="08:00 회차 전체")
    ap.add_argument("--plan", action="store_true", help="대상만 보기(안 바꾼다)")
    ap.add_argument("--print", dest="show", action="store_true", help="지난 회차 결과")
    ap.add_argument("--ack", action="store_true", help="인계 문서 배너 내리기")
    ap.add_argument("--days", type=int, default=DAYS)
    ap.add_argument("--limit", type=int, default=LIMIT)
    ap.add_argument("--no-absorb", action="store_true",
                    help="덤프 흡수를 건너뛴다(다른 세션이 수집 중일 때)")
    a = ap.parse_args(argv)

    if a.ack:
        ack()
        print("확인 처리했습니다 — 인계 문서 맨 위 배너를 내립니다.")
        return 0
    if a.plan:
        tg, floor = plan(a.days, a.limit)
        print(f"최근 {a.days}일 재수집 대상 (기준일 {floor} 이후)")
        for band, info in sorted(tg.items()):
            print(f"  밴드 {band}: {info['수']}건"
                  + (f" (상한 {a.limit} 넘침 {info['넘침']})" if info["넘침"] else ""))
        return 0
    if a.show and not a.run:
        st = load_state()
        if not st:
            print("아직 회차를 돈 적이 없습니다 — `--run` 을 먼저 도세요.")
            return 0
        show(st)
        return 0
    show(run(a.days, a.limit, do_absorb=not a.no_absorb))
    return 0


CRASH = _shared("reports", "밴드_재수집_오류.json")


def _leave_trace(exc):
    """죽은 이유를 **디스크에 남긴다** (2026-08-12, `[228]` 이 드러낸 것).

    ★ 이 회차는 스케줄러에서 `pythonw.exe` 로 돈다 — 창이 없으니 **트레이스백이
      어디에도 안 남는다.** 실측: `쿠팡업무_밴드재수집` 이 매일 08:00 에 exit 1 로
      끝나고 있었는데, 그 사실조차 `schedule_watch`(`[228]`) 를 만들고서야 보였고
      **왜인지는 그때도 알 길이 없었다.** 자국이 없으면 다음 사람도 똑같이 못 고친다.
    ★ 성공하면 **지운다** — 옛 자국이 남아 있으면 이미 고쳐진 고장을 계속 보고한다.
    """
    import traceback
    try:
        os.makedirs(os.path.dirname(CRASH), exist_ok=True)
        with open(CRASH, "w", encoding="utf-8") as fh:
            json.dump({"시각": datetime.now().isoformat(timespec="seconds"),
                       "명령": " ".join(sys.argv[1:]) or "(인자 없음)",
                       "무엇": "%s: %s" % (type(exc).__name__, exc),
                       "자취": traceback.format_exc()[-4000:]}, fh, ensure_ascii=False, indent=1)
    except Exception:
        pass                    # 자국을 남기려다 종료를 막지 않는다


if __name__ == "__main__":
    try:
        rc = main()
    except SystemExit:
        raise
    except BaseException as exc:                   # 죽어도 이유는 남기고 죽는다
        _leave_trace(exc)
        raise
    if rc == 0:
        try:
            os.remove(CRASH)
        except OSError:
            pass
    sys.exit(rc)
