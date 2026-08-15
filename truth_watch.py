#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""화면이 말하는 것과 원본이 말하는 것을 매일 대 본다 — **조용한 어긋남 감시**.

사용자 지시(2026-08-13): **"위 같은 문제 잡아내는 기능 AI 추가해"**

무엇이 '위 같은 문제'인가 — 그날 실측 둘 다 **형님이 전화로 지적하고 나서야** 알았다:
  · 캘린더 미처리 131건 중 **41건(31%)이 이미 취소된 건**이었다(`[243]`).
  · 다녀온 현장이 **원장 완료일이 비어** 미처리로 서 있었다(`[244]`).
둘 다 **오류가 안 났다.** 화면은 멀쩡히 숫자를 보여 줬고 파일도 다 있었다. 그래서
아무도 안 봤다 — 이 프로젝트가 반복해 당한 바로 그 모양이다(`[169]`).

**그래서 이 회차가 하는 일은 계산이 아니라 질문이다.** 매일 스스로 묻는다:
  ① 화면이 '아직 안 끝났다'고 하는 건을 **원본도 그렇게 말하나**
  ② 한 갈래가 목록의 **절반을 넘지 않나**(넘으면 경보가 아니라 기준이 틀린 것, `[170]`)
  ③ 지금 "0건"인 것이 **없는 것인가 안 본 것인가**(`[169]`)
  ④ 값이 전부 같은 열이 있나 — **안 읽은 열은 빈칸과 구별이 안 된다**(`[165]`)

★ **판정을 새로 만들지 않는다**(`[162]`). 화면이 실제로 내보내는 그 함수
  (`app_server._calendar_work_events`)를 **그대로 불러** 결과를 센다. 여기서 다시
  판정하면 화면과 감시가 갈리고, 갈린 뒤에는 어느 쪽이 맞는지 아무도 모른다 —
  그리고 **감시자가 틀리면 아무도 그 사실을 모른다.**

★ **읽기 전용이다.** 아무것도 안 고치고 큐에도 안 넣고 엑셀도 안 연다.
  무엇이 맞는지는 사람만 안다(`typo_watch` 와 같은 자리). 자동으로 고치면
  "그때 화면이 정말 뭐라고 했나"를 잃는다.

결과: `reports/화면_사실대조.md` · `reports/화면_사실대조.json`
"""
from __future__ import annotations

import json
import os
import sys
from datetime import datetime, timedelta

if hasattr(sys.stdout, "reconfigure"):      # 무인 회차는 pythonw 라 stdout 이 None 이다([235])
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.abspath(__file__))
REPORT_MD = os.path.join(ROOT, "reports", "화면_사실대조.md")
REPORT_JSON = os.path.join(ROOT, "reports", "화면_사실대조.json")

# 한 갈래가 목록의 이만큼을 넘으면 그것은 경보가 아니라 **기준 이야기**다([170]).
DOMINANT = 0.5
# 근거가 이보다 오래되면 '없다'고 말할 수 없다([169]). 밴드는 매일 올라온다.
BAND_STALE_DAYS = 2


def _app():
    """화면이 쓰는 그 모듈을 그대로 부른다 — 사본을 만들지 않는다([162])."""
    p = os.path.join(ROOT, "webapp")
    if p not in sys.path:
        sys.path.insert(0, p)
    import app_server
    return app_server


def _today():
    return datetime.now().strftime("%Y-%m-%d")


def collect():
    """묻고 답을 모은다. 못 물어본 것은 **못 물어봤다고 적는다**([169])."""
    out = {"때": datetime.now().isoformat(timespec="seconds"), "물음": [], "못물음": []}
    try:
        A = _app()
        ev = A._calendar_work_events() or []
    except Exception as exc:
        out["못물음"].append("캘린더 집계를 못 불렀다: %s" % exc)
        return out

    미처리 = [e for e in ev if e.get("분류") in ("as_open", "pm_overdue")]
    out["미처리"] = len(미처리)
    out["밴드근거완료"] = sum(1 for e in ev if e.get("원장미기입"))

    # ── ① 화면이 '안 끝났다'는 것을 원본도 그렇게 말하나 ──────────────────
    #    화면이 이미 이유를 붙여 두므로(`[244]`) 여기서 다시 판정하지 않고 **센다**.
    사유 = {}
    이유없음 = 0
    for e in 미처리:
        r = e.get("미처리사유")
        if not r:
            이유없음 += 1
            continue
        사유[r] = 사유.get(r, 0) + 1
    out["사유"] = 사유
    if 이유없음:
        # 이유를 못 붙인 건이 있다는 것은 판정 어딘가가 이 길을 안 탄다는 뜻이다.
        out["물음"].append({
            "무엇": "미처리인데 이유가 안 붙은 건",
            "값": 이유없음,
            "왜": "화면이 '왜 아직인지'를 못 말하면 사람이 다녀온 현장을 다시 찾아간다",
            "등급": "경보" if 이유없음 else "이상없음"})

    # ── ② 한 갈래가 절반을 넘나 — 넘으면 경보가 아니라 기준 이야기다([170]) ──
    if 미처리:
        큰갈래, 큰수 = max(사유.items(), key=lambda kv: kv[1]) if 사유 else ("", 0)
        비율 = 큰수 / float(len(미처리))
        out["물음"].append({
            "무엇": "미처리 사유가 한 갈래에 몰렸나",
            "값": "%s %d건 (%.0f%%)" % (큰갈래 or "(사유 없음)", 큰수, 비율 * 100),
            "왜": "한 갈래가 절반을 넘으면 그 경보는 아무도 안 본다 — 기준부터 본다",
            "등급": "확인" if 비율 > DOMINANT else "이상없음"})

    # ── ③ 지금 0건인 것이 없는 것인가 안 본 것인가([169]) ────────────────
    try:
        idx = _app()._band_completion_index()
    except Exception as exc:
        idx = None
        out["못물음"].append("밴드·카톡 완료 색인을 못 읽었다: %s" % exc)
    if idx is not None:
        최신 = idx.get("최신") or ""
        늦음 = None
        if 최신:
            try:
                늦음 = (datetime.strptime(_today(), "%Y-%m-%d")
                        - datetime.strptime(최신, "%Y-%m-%d")).days
            except Exception:
                늦음 = None
        out["밴드최신"] = 최신
        out["밴드밀림일"] = 늦음
        if not idx.get("읽음"):
            # ★ 못 읽은 것을 '근거 없음'이라 하지 않는다. 그러면 이 감시자 자신이
            #   눈먼 채 "이상 없음"을 말한다.
            out["물음"].append({
                "무엇": "밴드·카톡 근거를 읽었나",
                "값": "못 읽음",
                "왜": "못 읽은 것을 '완료 근거 0건'으로 치면 감시자가 눈멀고도 조용하다",
                "등급": "경보"})
        elif 늦음 is None or 늦음 > BAND_STALE_DAYS:
            out["물음"].append({
                "무엇": "밴드 수집이 어디까지 왔나",
                "값": "%s (%s)" % (최신 or "모름",
                                   "%d일 밀림" % 늦음 if 늦음 is not None else "날짜 모름"),
                "왜": "'완료 글이 없다'와 '아직 안 긁었다'는 다른 말이다 — "
                      "밀린 동안의 미처리 숫자는 그만큼 많게 나온다",
                "등급": "확인"})
        else:
            out["물음"].append({
                "무엇": "밴드 수집이 어디까지 왔나",
                "값": "%s (%d일 전)" % (최신, 늦음),
                "왜": "", "등급": "이상없음"})

    # ── ④ 되돌아가면 안 되는 것 — 취소·완료가 미처리에 다시 섞이나([243]) ──
    섞임 = [e for e in 미처리
            if (e.get("진행상태") or e.get("점검상태") or "") in ("취소", "철회")]
    out["물음"].append({
        "무엇": "취소된 건이 미처리에 섞였나",
        "값": len(섞임),
        "왜": "2026-08-13 에 41건이 그 상태였다 — 되돌아가면 같은 사고다([243])",
        "등급": "경보" if 섞임 else "이상없음"})

    # ── ⑤ 안 읽은 열 — 값이 전부 비면 빈칸과 구별이 안 된다([165]) ────────
    try:
        works = _app().get_works() or {}
        for kind, cols in (("as", ("진행상태", "접수일자", "작업완료일", "캠프명",
                                   "프로젝트NO", "유상·무상·보험")),
                           ("pm", ("점검상태", "점검예정일", "실제점검일", "캠프명",
                                   "프로젝트NO"))):
            rows = works.get(kind) or []
            if not rows:
                continue
            for c in cols:
                찬 = sum(1 for r in rows if str(r.get(c) or "").strip())
                if 찬 == 0:
                    out["물음"].append({
                        "무엇": "%s 시트 '%s' 열이 %d행 전부 비었다" % (kind, c, len(rows)),
                        "값": 0,
                        "왜": "안 읽은 열은 빈칸과 구별할 수 없다 — 열 이름이 바뀌었을 수 있다([165])",
                        "등급": "경보"})
        # ── ⑥ 앱 DB에 없는데 화면엔 뜨는 행 — [89] 사각지대 ───────────────
        #   정기점검 '예정월' 계획처럼 아직 work_item 이 없는 행이다. 저장을 누르면
        #   거절되므로 화면이 '신규 등록' 으로 안내한다(inputForm 게이트, [89]).
        #   평소 값(실측 pm 70)이면 정상이지만, 갑자기 늘면 흡수가 밀렸다는 뜻일 수
        #   있어 매일 세어 눈에 둔다([169]). works 는 위에서 이미 불렀다 — 다시 안 부른다.
        for kind in ("as", "pm"):
            rows = works.get(kind) or []
            if not rows:
                continue
            missing = sum(1 for r in rows if not r.get("_store_id"))
            if missing:
                out["물음"].append({
                    "무엇": "%s 화면에 앱 DB 미등록 행 %d건 (총 %d)" % (kind, missing, len(rows)),
                    "값": missing,
                    "왜": ("예정·계획 등 아직 work_item 이 없는 행 — 저장 대신 신규 "
                           "등록으로 안내한다([89]). 갑자기 늘면 흡수 밀림을 의심한다"),
                    "등급": "이상없음"})
    except Exception as exc:
        out["못물음"].append("원장 열을 못 셌다: %s" % exc)

    return out


def render(d):
    L = ["# 화면이 말하는 것 ↔ 원본이 말하는 것", ""]
    L.append("만든 때: %s" % d.get("때", ""))
    L.append("")
    L.append("> 이 회차는 **아무것도 고치지 않는다.** 무엇이 맞는지는 사람만 안다.")
    L.append("> 판정은 화면이 쓰는 그 함수를 그대로 불러 세운 것이다 — 사본이 아니다.")
    L.append("")
    if d.get("못물음"):
        L.append("## ⚠ 못 물어본 것 (이상 없음이 아니다)")
        for m in d["못물음"]:
            L.append("- %s" % m)
        L.append("")
    급 = [q for q in d.get("물음", []) if q.get("등급") == "경보"]
    확 = [q for q in d.get("물음", []) if q.get("등급") == "확인"]
    무 = [q for q in d.get("물음", []) if q.get("등급") == "이상없음"]
    for 제목, 목록 in (("## ★ 먼저 볼 것", 급), ("## 확인해 볼 것", 확),
                       ("## 이상 없음", 무)):
        if not 목록:
            continue
        L.append(제목)
        for q in 목록:
            L.append("- **%s** — %s" % (q["무엇"], q["값"]))
            if q.get("왜"):
                L.append("  · %s" % q["왜"])
        L.append("")
    L.append("## 지금 숫자")
    L.append("- 미처리 %s건 · 밴드·카톡 근거로 완료 처리 %s건"
             % (d.get("미처리", "?"), d.get("밴드근거완료", "?")))
    for k, v in sorted((d.get("사유") or {}).items(), key=lambda kv: -kv[1]):
        L.append("  · %d건 — %s" % (v, k))
    return "\n".join(L) + "\n"


def main(argv=None):
    argv = list(argv if argv is not None else sys.argv[1:])
    d = collect()
    os.makedirs(os.path.dirname(REPORT_MD), exist_ok=True)
    with open(REPORT_JSON, "w", encoding="utf-8") as fh:
        json.dump(d, fh, ensure_ascii=False, indent=1)
    with open(REPORT_MD, "w", encoding="utf-8") as fh:
        fh.write(render(d))
    급 = sum(1 for q in d.get("물음", []) if q.get("등급") == "경보")
    확 = sum(1 for q in d.get("물음", []) if q.get("등급") == "확인")
    print("화면 사실대조 — 먼저 볼 것 %d · 확인 %d · 못 물어봄 %d → %s"
          % (급, 확, len(d.get("못물음") or []), os.path.basename(REPORT_MD)))
    if "--print" in argv:
        print(render(d))
    return 0                     # 감시자가 회차를 죽이지 않는다


if __name__ == "__main__":
    raise SystemExit(main())
