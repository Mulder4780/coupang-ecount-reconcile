# -*- coding: utf-8 -*-
"""
cross_signal.py — 카톡과 밴드 댓글을 **한 사건으로 묶는다** (읽기 전용)
=======================================================================
사용자 지시(2026-08-09): "카톡도 댓글이랑 연관지어서 생각하고 반영하는 알고리즘 구현해"

## 왜 필요한가 — 같은 사건이 두 군데로 나뉘어 온다
현장에서 일이 바뀌면(취소·연기·재방문·금액 변경) 그 말은 **어디로 올지 정해져 있지 않다.**
캠프 담당자는 카톡방에 쓰고, 기사는 밴드 글 댓글에 단다. 둘 중 한 곳만 보면
**같은 사건의 반쪽만 본다.**

지금까지 둘은 서로 모르는 채 돌았다:
  · `cancel_watch.py` 는 **밴드**만 본다(댓글이 들어오기 전에는 그마저 반쪽이었다)
  · `kakao_reconcile.py` 는 **카톡**만 본다
그래서 카톡에만 온 취소는 밴드 화면에 안 뜨고, 댓글에만 달린 취소는 카톡 대조에 안 뜬다.
**어느 쪽에도 오류는 안 난다.** 조용한 사고다.

## 무엇을 묶나 — 프로젝트번호가 있으면 그것이 최우선
짝을 지을 근거가 약하면 **엉뚱한 현장 둘을 한 사건으로 만든다.** 그건 못 묶는 것보다
나쁘다(가지 말아야 할 곳에 사람을 보내거나, 가야 할 곳을 지운다). 그래서:
  ① 양쪽의 **프로젝트NO가 정확히 같을 것**. 있으면 캠프 표기가 빠져도 묶는다.
  ② 프로젝트NO가 한쪽이라도 없을 때만 캠프+날짜(±3일)를 보고용으로 묶는다.
  ③ 어느 경우든 취소·연기·재방문 같은 사건 종류가 양쪽에 같아야 한다.

## ★ 아무것도 고치지 않는다
`cancel_watch` 가 원장 대조와 큐를 맡는다. 이 도구는 **두 원본이 같은 말을 하는지**만
보고한다. 판정을 두 곳에서 하면 언젠가 갈린다 — 오늘 배운 것이 그것이다.

실행:  python cross_signal.py            # reports/카톡_밴드_교차.md
       python cross_signal.py --json
"""
import sys, os, re, json, glob
from datetime import datetime, date, timedelta
from collections import defaultdict

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")
OUT_MD = os.path.join(REPORT_DIR, "카톡_밴드_교차.md")
OUT_JSON = os.path.join(REPORT_DIR, "카톡_밴드_교차.json")
CACHE = os.path.join(ROOT, "band", "cache")

NEAR_DAYS = 3

# 사건어 — '무슨 일이 났다'고 말하는 낱말. 종류를 나눠 두면 양쪽이 **같은 종류**를
# 말할 때만 묶을 수 있다. '취소'와 '연기'를 한 덩어리로 두면 서로 다른 사건이 붙는다.
EVENTS = {
    "취소": (r"취소", ),
    "연기": (r"연기", r"미루", r"보류", r"다음\s*주", r"일정\s*변경"),
    "재방문": (r"재방문", r"다시\s*방문", r"재작업", r"재점검"),
    "금액변경": (r"금액\s*변경", r"견적\s*변경", r"단가\s*변경", r"금액\s*수정"),
}


def events_in(text):
    t = str(text or "")
    return {k for k, pats in EVENTS.items() if any(re.search(p, t) for p in pats)}


def camp_core(camp):
    """괄호·공백 앞 이름만 — 카톡은 `송파5MB(감일동)` 을 `송파5MB` 로도 `송파5` 로도 쓴다."""
    return re.split(r"[(\s]", str(camp or ""))[0].strip()


def band_signals():
    """밴드 글 **본문과 댓글**에서 사건어를 뽑는다. 댓글은 글보다 나중이라 더 무겁다."""
    import band_extract
    out = []
    for f in sorted(glob.glob(os.path.join(CACHE, "*.json"))):
        b = os.path.basename(f)[:-5]
        if not (b.isdigit() and len(b) == 8):        # 유령 밴드는 안 본다
            continue
        with open(f, encoding="utf-8") as fh:
            posts = (json.load(fh).get("posts") or {})
        for no, p in posts.items():
            if not isinstance(p, dict) or p.get("deleted") or p.get("ghost") or p.get("dirty"):
                continue
            # ★ `band_day` 는 글이 아니라 **`created_at` 값**을 받는다(2026-08-09 실수).
            #   글을 통째로 넘기면 조용히 "" 를 돌려주고 **모든 글이 건너뛰어진다** —
            #   오류 없이 "밴드 신호 0" 이 나왔다. 날짜 한 곳으로 모아 둔 이유가 이것이다.
            from datalake import band_day
            day = band_day(p.get("created_at"))
            if not day:
                continue                              # 시각 없는 글은 순서를 못 세운다
            body = str(p.get("content") or p.get("body") or "")
            cmts = ""
            try:
                cmts = band_extract.comment_text(p) or ""
            except Exception:
                cmts = " ".join(str(c.get("content") or "")
                                for c in (p.get("comments") or []) if isinstance(c, dict))
            camp = camp_core(p.get("camp") or _camp_from(body))
            project_match = band_extract.RE_PRJ.search(body)
            project = project_match.group(1).upper() if project_match else ""
            for where, text in (("본문", body), ("댓글", cmts)):
                ev = events_in(text)
                if ev and (camp or project):
                    out.append({"출처": f"밴드 {where}", "밴드": b, "글번호": str(no),
                                "날짜": str(day)[:10], "캠프": camp,
                                "프로젝트NO": project,
                                "사건": sorted(ev), "글": text[:160]})
    return out


CAMP_PAT = re.compile(r"([가-힣]+\d*(?:MB|캠프|Sub-hub|Sub-FC|SPA)[^\s,]*)")


def _camp_from(text):
    m = CAMP_PAT.search(str(text or ""))
    return m.group(1) if m else ""


def kakao_signals():
    sys.path.insert(0, os.path.join(ROOT, "kakao"))
    import kakao_reconcile as K
    out = []
    for path in K.source_paths():
        try:
            msgs = K.parse_export(path)
        except Exception:
            continue
        for m in msgs:
            ev = events_in(m["text"])
            if not ev:
                continue
            camp = camp_core(_camp_from(m["text"]))
            from band_extract import RE_PRJ
            project_match = RE_PRJ.search(m["text"])
            project = project_match.group(1).upper() if project_match else ""
            if not camp and not project:
                continue
            out.append({"출처": "카톡", "파일": os.path.basename(path),
                        "날짜": m["date"].isoformat(), "보낸이": m.get("sender", ""),
                        "캠프": camp, "프로젝트NO": project,
                        "사건": sorted(ev), "글": m["text"][:160]})
    return out


def pair(bs, ks, near=NEAR_DAYS):
    """정확 프로젝트를 우선하고, 없을 때만 캠프·날짜로 보고용 연결한다."""
    matched, band_only = [], []
    used = set()
    for b in bs:
        bd = date.fromisoformat(b["날짜"])
        candidates = []
        for i, k in enumerate(ks):
            if i in used:
                continue
            if abs((date.fromisoformat(k["날짜"]) - bd).days) > near:
                continue
            common = set(b["사건"]) & set(k["사건"])
            if not common:
                continue
            b_project = str(b.get("프로젝트NO") or "").strip().upper()
            k_project = str(k.get("프로젝트NO") or "").strip().upper()
            exact_project = bool(b_project and k_project and b_project == k_project)
            same_camp = bool(b.get("캠프") and b.get("캠프") == k.get("캠프"))
            if exact_project or (not (b_project and k_project) and same_camp):
                candidates.append((0 if exact_project else 1,
                                   abs((date.fromisoformat(k["날짜"]) - bd).days),
                                   i, k, sorted(common), exact_project))
        hit = min(candidates, default=None, key=lambda item: (item[0], item[1], item[2]))
        if hit:
            _, _, index, kakao, common, exact_project = hit
            used.add(index)
            project = str(b.get("프로젝트NO") or kakao.get("프로젝트NO") or "")
            matched.append({"프로젝트NO": project, "캠프": b.get("캠프") or kakao.get("캠프"),
                            "연결근거": "프로젝트NO 정확 일치" if exact_project else "캠프+날짜",
                            "사건": common,
                            "밴드": f"{b['밴드']}/{b['글번호']} ({b['출처']}) {b['날짜']}",
                            "카톡": f"{kakao['파일']} {kakao['날짜']} {kakao['보낸이']}",
                            "밴드글": b["글"], "카톡글": kakao["글"]})
        else:
            band_only.append(b)
    kakao_only = [k for i, k in enumerate(ks) if i not in used]
    return matched, band_only, kakao_only


def run(near=NEAR_DAYS):
    bs, ks = band_signals(), kakao_signals()
    matched, b_only, k_only = pair(bs, ks, near)
    os.makedirs(REPORT_DIR, exist_ok=True)
    L = ["# 카톡 · 밴드 교차 확인", "",
         f"- 생성 {datetime.now():%Y-%m-%d %H:%M} · 밴드 신호 {len(bs)} · 카톡 신호 {len(ks)}",
         f"- 짝지어짐 **{len(matched)}** · 밴드에만 **{len(b_only)}** · 카톡에만 **{len(k_only)}**",
         "- 자동연결 1순위: 프로젝트NO 정확 일치 · 같은 사건 · 날짜(±%d일)" % near,
         "- 프로젝트NO가 없을 때의 캠프+날짜 연결은 보고 근거이며 자동 취소에는 쓰지 않습니다.",
         "- **이 도구는 아무것도 고치지 않습니다.** 원장 대조·큐는 cancel_watch 몫입니다.", ""]

    def tbl(title, rows, keys, why=""):
        L.append(f"## {title} — {len(rows)}건")
        if why:
            L.append(why)
        if not rows:
            L.append("없음")
            L.append("")
            return
        L.append("| " + " | ".join(keys) + " |")
        L.append("|" + "---|" * len(keys))
        for r in rows[:150]:
            L.append("| " + " | ".join(str(r.get(k, "")).replace("|", "/") for k in keys) + " |")
        if len(rows) > 150:
            L.append(f"\n… 외 {len(rows) - 150}건 (전체는 카톡_밴드_교차.json)")
        L.append("")

    tbl("두 곳이 같은 말을 한다", matched,
        ["프로젝트NO", "캠프", "연결근거", "사건", "밴드", "카톡"],
        "양쪽 원본이 같은 사건을 말합니다 — 근거가 가장 셉니다.")
    tbl("밴드에만 있다", b_only, ["캠프", "날짜", "사건", "출처", "글"],
        "카톡방에는 안 올라온 사건입니다. 캠프가 모르고 있을 수 있습니다.")
    tbl("카톡에만 있다", k_only, ["캠프", "날짜", "사건", "보낸이", "글"],
        "밴드 글·댓글에 안 남은 사건입니다. **기사에게 안 전달됐을 수 있습니다.**")

    with open(OUT_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump({"생성": datetime.now().isoformat(timespec="seconds"),
                   "밴드신호": len(bs), "카톡신호": len(ks),
                   "짝지어짐": matched, "밴드에만": b_only, "카톡에만": k_only},
                  f, ensure_ascii=False, indent=1)
    return matched, b_only, k_only, len(bs), len(ks)


def main():
    args = sys.argv[1:]
    near = int(args[args.index("--near") + 1]) if "--near" in args else NEAR_DAYS
    m, b, k, nb, nk = run(near)
    if "--json" in args:
        print(json.dumps({"밴드신호": nb, "카톡신호": nk, "짝지어짐": len(m),
                          "밴드에만": len(b), "카톡에만": len(k)}, ensure_ascii=False))
    else:
        print(f"밴드 신호 {nb} · 카톡 신호 {nk} → 짝지어짐 {len(m)} · "
              f"밴드에만 {len(b)} · 카톡에만 {len(k)}")
        print("리포트:", OUT_MD)
    return 0


if __name__ == "__main__":
    sys.exit(main())
