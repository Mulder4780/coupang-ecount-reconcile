# -*- coding: utf-8 -*-
"""
local_ai.py — 앱이 스스로 답한다 (클로드·코덱스 없이)
=====================================================
사용자 지시(2026-08-09): "클로드 코드 또는 코덱스 없이 앱 자체적으로 처리할 수 있는
AI 알고리즘 기능 넣어서 클로드코드 크래딧 사용 최대한 아낄 수 있게 설계하고
알고리즘에 반영해"

## 크레딧은 '작업'이 아니라 '되묻기'에서 샌다
지금까지의 자동화는 전부 **일을 하는** 쪽이었다(수집·대조·큐·반영). 그런데 크레딧은
거기서 안 샜다. 실제로 샌 자리는 이렇다 — 사람이 앱을 열어 이상한 숫자를 보고,
**앱이 이유를 못 말해서** 클로드에게 묻는다. 그러면 클로드는 매번 리포트를 열고
ERP 를 대 보고 원장을 뒤져서 **이미 디스크에 있는 사실**을 다시 조립한다.
실측된 예: "작업은 완료인데 왜 계산서 발행이 안된거지" 한 마디가 도구 호출 열몇 번이 됐고,
답은 결국 `ERP 4.세금계산서발행대기` 라는 **파일에 이미 적혀 있던 한 줄**이었다.

그러니 아껴야 할 것은 계산이 아니라 **왕복**이다. 이 파일은 그 왕복을 없앤다.

## 규칙 기반이다 — 이 프로젝트에서는 그게 더 세다
언어모델을 부르지 않는다. 부를 필요가 없다. 이 업무의 질문은 몇 가지 모양으로 반복되고,
답의 근거는 **전부 구조화된 파일**에 있다(리포트 JSON·색인). 모델을 부르면 돈이 들고,
느리고, 무엇보다 **지어낼 수 있다.** 규칙은 근거가 없으면 답을 못 하고, 못 하는 것이
이 프로젝트에서는 미덕이다(오기입 탐지 `[172]` 와 같은 이유다).

## 세 갈래로만 답한다 — 여기가 핵심이다
  ① **답함**   — 근거 파일이 있고 신선하다. 사실과 출처와 '다음에 무엇을 하나'를 준다.
  ② **모름**   — 규칙에 없는 질문이다. **지어내지 않는다.** 대신
                 **클로드에게 붙여넣을 문구**를 만들어 준다. 그 문구에는 관련 리포트
                 이름·수치가 이미 박혀 있어서, 클로드가 처음부터 뒤질 일이 없다.
                 못 답해도 **한 번의 왕복은 짧아진다.**
  ③ **낡음**   — 규칙은 있는데 근거 파일이 한도보다 오래됐다. 낡은 값을 자신 있게
                 말하는 것이 이 프로젝트에서 제일 위험하다(조용한 사고). 그래서
                 **몇 시간 전 자료인지 밝히고** 답한다. 숨기지 않는다.

## 스스로 자란다 — 모델 없이
못 답한 질문은 `reports/앱_자문기록.json` 에 쌓인다. `--stats` 가 그것을 세어
**"이 모양의 질문이 N번 왔는데 아직 규칙이 없다"** 를 보여 준다. 다음에 만들 규칙의
목록이 추측이 아니라 기록에서 나온다(UX 개선을 기록으로 하는 것과 같은 원리).
그리고 **얼마나 아꼈는지도 그 파일이 답한다** — 앱이 답한 비율이 곧 안 물어본 횟수다.

## 절대 하지 않는 것
  · **아무것도 고치지 않는다.** 읽기 전용이다. 큐에도 안 넣고 엑셀은 열지도 않는다.
    답변기가 원장을 건드리면 "물어봤을 뿐인데 값이 바뀌었다"가 된다.
  · **비싼 탐색을 하지 않는다.** Z: 재귀 glob 은 부르지 않는다(`[168]`). 리포트와
    이미 만들어진 색인만 읽는다. 질문 하나에 1초를 넘기면 사람은 그냥 클로드에게 묻는다.
  · **근거 없이 단정하지 않는다.** 근거 파일 이름을 답마다 같이 준다.

쓰는 법:
    python local_ai.py "왜 계산서가 안 나갔지"
    python local_ai.py "UJ2600021"
    python local_ai.py --stats
    python local_ai.py --selftest
앱에서는 `/api/ask?q=...` 로 같은 답이 나온다.

검증 [181].
"""
import os
import re
import io
import json
import glob
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")
LOG = os.path.join(REPORT_DIR, "앱_자문기록.json")
LOG_MAX = 500          # 기록은 굴러간다 — 무한히 쌓아 두면 읽는 쪽이 느려진다


# ─────────────────────────────────────────────────────────────────────────────
# 근거 — 어디서 읽나
#
# 한 곳에 모아 둔다. 규칙이 늘어도 "이 답의 근거가 어느 파일인가"가 한 화면에 보인다.
# `한도시간` 은 **그 자료가 몇 시간까지 쓸 만한가**다. 넘으면 답을 막지 않고
# '낡음'으로 표시해 같이 말한다 — 낡았다고 입을 다물면 사람은 클로드에게 묻는다.
# ─────────────────────────────────────────────────────────────────────────────
SOURCES = {
    "인계":     {"glob": "세션인계.json",          "한도시간": 2},
    "큐":       {"glob": "반영대기.json",          "한도시간": 6},
    "회차":     {"glob": ".daily_run.progress.json", "한도시간": 26},
    "오기입":   {"glob": "오기입_확인.json",        "한도시간": 48},
    "교차":     {"glob": "카톡_밴드_교차.json",     "한도시간": 48},
    "ERP색인":  {"glob": "ERP판매_프로젝트색인.json", "한도시간": 72},
    "PO근거":   {"glob": "po_objective_evidence.json", "한도시간": 48},
    "PO대조":   {"glob": "PO대조_*.md",             "한도시간": 48},
    "취소":     {"glob": "접수취소_확인.md",         "한도시간": 48},
}


def _latest(pattern):
    """이름에 시각이 박힌 리포트는 **가장 새것**을 고른다."""
    hits = glob.glob(os.path.join(REPORT_DIR, pattern))
    if not hits:
        return None
    return max(hits, key=lambda p: os.path.getmtime(p))


# ★ 한 번 읽은 근거는 붙들고 있는다 (`[168]` 을 또 밟았다).
#   `answer_pack()` 이 프로젝트마다 `a_project()` 를 부르고 그 안에서 `src("PO대조")`
#   가 **매번 md 를 다시 읽고 다시 파싱했다** — 2,014번. 그래서 꾸러미 하나 만드는 데
#   94.5초가 걸렸다. 비싼 읽기는 언제나 캐시 검사 **뒤에** 온다.
#   TTL 을 짧게 둔 이유는 질문 하나 안에서만 붙들면 충분하기 때문이다 —
#   길게 잡으면 리포트가 갱신됐는데도 옛 답을 주는 반대쪽 사고가 된다.
_SRC_CACHE = {}
SRC_TTL = 20.0


def _cache_get(name):
    # ★ 열쇠에 **폴더까지** 넣는다. 이름만으로 잡으면 `COUPANG_REPORT_DIR` 이 바뀌어도
    #   옛 폴더의 답을 그대로 준다 — 워크트리에서 본체 리포트를 보고 답하는 셈이라
    #   화면은 멀쩡하고 값만 남의 것이 된다(검증 [181] 이 실제로 이걸 잡았다).
    hit = _SRC_CACHE.get((REPORT_DIR, name))
    if not hit:
        return None
    at, val = hit
    if (datetime.now() - at).total_seconds() > SRC_TTL:
        return None
    return val


def src(name):
    """근거 하나를 읽어 온다.

    없으면 없다고, 낡았으면 몇 시간 됐는지까지 담아 돌려준다.
    **예외를 올리지 않는다** — 근거 하나가 깨졌다고 답변기 전체가 죽으면,
    사람은 앱을 못 믿고 다시 클로드에게 묻는다(그게 이 파일이 없애려는 왕복이다).
    """
    cached = _cache_get(name)
    if cached is not None:
        return cached
    spec = SOURCES.get(name) or {}
    path = _latest(spec.get("glob") or "")
    out = {"이름": name, "파일": None, "있음": False, "나이시간": None,
           "신선": False, "데이터": None}
    if not path or not os.path.exists(path):
        # '없음'도 기억한다 — 없는 것을 찾는 glob 이 제일 오래 돈다(`[168]`).
        _SRC_CACHE[(REPORT_DIR, name)] = (datetime.now(), out)
        return out
    out["파일"] = os.path.basename(path)
    out["있음"] = True
    try:
        age = (datetime.now().timestamp() - os.path.getmtime(path)) / 3600.0
        out["나이시간"] = round(age, 1)
        out["신선"] = age <= float(spec.get("한도시간") or 24)
    except OSError:
        pass
    if path.lower().endswith(".json"):
        try:
            out["데이터"] = json.load(io.open(path, encoding="utf-8"))
        except Exception:
            out["데이터"] = None
    else:
        try:
            out["데이터"] = io.open(path, encoding="utf-8", errors="replace").read()
        except OSError:
            out["데이터"] = None
    _SRC_CACHE[(REPORT_DIR, name)] = (datetime.now(), out)
    return out


def _age_note(*sources):
    """'몇 시간 전 자료인가' 한 줄. 낡은 것이 있으면 그것을 앞세운다."""
    stale = [s for s in sources if s.get("있음") and not s.get("신선")]
    if stale:
        s = stale[0]
        return "⚠ %s 는 %s시간 전 자료입니다 — 그 사이 바뀌었을 수 있습니다." % (
            s["파일"], s["나이시간"])
    fresh = [s for s in sources if s.get("있음")]
    if fresh:
        return "근거: " + " · ".join("%s(%s시간 전)" % (s["파일"], s["나이시간"])
                                     for s in fresh)
    return "근거 파일이 아직 없습니다."


# ─────────────────────────────────────────────────────────────────────────────
# 규칙들
#
# 각 규칙은 `(이름, 말투들, 답하는 함수)` 다. 함수는 답을 못 하겠으면 **None** 을
# 돌려준다 — 그러면 '모름' 갈래로 내려가 클로드 문구를 만든다. 억지로 문장을
# 만들어 내지 않는 것이 규칙이 지켜야 할 첫째 예의다.
# ─────────────────────────────────────────────────────────────────────────────
PRJ_PAT = re.compile(r"\b([A-Z]{2}\d{7})\b")
PO_PAT = re.compile(r"\bPO\s?(\d{6})\b", re.I)


def _who_next(state):
    """ERP 진행상태 한 줄을 '누가 무엇을 하면 되나'로 옮긴다.

    상태 코드를 그대로 보여 주면 사람은 그게 무슨 뜻인지 다시 묻는다.
    **딱지를 사람 이름과 행동으로 바꾸는 것**이 이 함수가 하는 전부다.
    """
    s = str(state or "")
    if s[:2] == "7.":
        return "끝났습니다(수금까지 완료)."
    if s[:2] == "6.":
        return "세금계산서까지 발행됐습니다. 남은 것은 수금입니다."
    if s[:2] == "4.":
        return ("**발행 대기** — PO 도 왔고 금액도 맞습니다. ERP 에서 다음 단계로 "
                "넘기기만 하면 됩니다(류지영).")
    if s[:2] == "5.":
        return "거래명세서 단계입니다. 다음은 세금계산서 발행입니다(류지영)."
    if s[:2] == "3.":
        return "오더처리 단계입니다 — 아직 명세서 전입니다."
    if not s:
        return ("**ERP 전표가 아예 없습니다.** 발행 대기와 다른 문제입니다 — "
                "등록부터 해야 합니다.")
    return "ERP 상태: " + s


def _po_sections():
    """PO 대조 리포트의 구역 머리글에서 갈래별 건수를 읽는다.

    `## A. 미청구 PO — 계산서 미발행 (★) — 5건` 같은 줄이다. 표를 파싱하지 않고
    **머리글만** 읽는다 — 표 모양은 바뀌어도 머리글은 사람이 읽으라고 쓴 것이라
    잘 안 바뀐다(ERP 원장류를 머리글로 가른 `[173]` 과 같은 판단이다).
    """
    s = src("PO대조")
    out = []
    for line in (s.get("데이터") or "").splitlines():
        m = re.match(r"^##\s+([A-Z])\.\s*(.+?)\s*—\s*(\d+)\s*건\s*$", line.strip())
        if m:
            out.append({"갈래": m.group(1), "이름": m.group(2).strip(),
                        "건수": int(m.group(3))})
    return s, out


def _po_verdict(prj=None, po=None):
    """번호 하나가 PO 대조 표 어디에 어떤 판정으로 들어 있나.

    표의 **마지막 칸이 판정**이다. 그 문장은 사람이 읽으라고 쓰인 것이므로
    다시 지어내지 않고 **그대로** 준다 — 옮겨 적으면 뜻이 흔들린다.
    """
    s = src("PO대조")
    txt = s.get("데이터") or ""
    if not txt:
        return None
    want = [x for x in (po, prj) if x]
    if not want:
        return None
    section = ""
    for line in txt.splitlines():
        if line.startswith("## "):
            section = line[3:].strip()
            continue
        if not line.startswith("|") or line.startswith("|---"):
            continue
        if not any(w in line for w in want):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 2 or cells[0] in ("PO번호", "정산ID"):
            continue
        return {"판정": cells[-1], "구역": section, "파일": s["파일"]}
    return None


def a_project(q):
    """프로젝트NO·PO번호가 질문에 있으면 그것부터 답한다.

    ★ 번호로 물을 수 있으면 금액으로 묻지 않는다(`[170]` 과 같은 원칙).
    질문에 번호가 들어 있는 것은 사람이 **이미 무엇을 볼지 정해 준 것**이다.
    """
    m = PRJ_PAT.search(q or "")
    po = PO_PAT.search(q or "")
    if not m and not po:
        return None
    idx = src("ERP색인")
    data = idx.get("데이터") or {}
    # ★ 색인 파일은 `{src, count, index}` 다 — 알맹이는 `index` 아래 있다.
    #   이걸 모르고 통째로 뒤지면 **모든 번호가 '없음'으로 나온다**. 오류는 안 난다.
    if isinstance(data, dict) and "index" in data:
        data = data.get("index") or {}
    rec = None
    prj = m.group(1) if m else None
    if prj and isinstance(data, dict):
        rec = data.get(prj)
    if rec is None and po:
        # PO 로 물었으면 색인을 PO 로 훑는다 — 색인은 이미 메모리에 있으니 싸다.
        want = "PO" + po.group(1)
        if isinstance(data, dict):
            for k, v in data.items():
                if want in str((v or {}).get("po") or ""):
                    prj, rec = k, v
                    break
    # PO 대조 리포트가 이 번호를 **이미 판정해 뒀으면** 그 문장을 그대로 준다.
    # 사람이 클로드에게 물었을 때 클로드가 하던 일이 정확히 이것이다 —
    # 리포트를 열어 그 줄을 찾아 읽어 주는 것.
    verdict = _po_verdict(prj, po.group(0) if po else None)

    if rec is None:
        if not idx["있음"] and not verdict:
            return None            # 근거 자체가 없다 — 모름으로 내린다
        who = prj or ("PO" + po.group(1))
        if verdict:
            return {"답": "%s — %s" % (who, verdict["판정"]),
                    "다음": "표 전체는 `reports/%s` 의 %s 구역에 있습니다."
                            % (verdict["파일"], verdict["구역"]),
                    "근거": _age_note(src("PO대조")), "확신": "높"}
        return {
            "답": "%s 를 ERP 판매조회 색인에서 찾지 못했습니다. ERP 전표가 아직 "
                  "없거나, 번호가 잘못 적혔을 수 있습니다." % who,
            "다음": "오기입 의심은 `reports/오기입_확인.md` 를 먼저 보십시오.",
            "근거": _age_note(idx),
            "확신": "중",
        }
    # ★ 색인의 칸 이름은 **영문**이다(state/supply/cust/po/date).
    #   한글 이름으로만 물으면 늘 빈칸이 나오고 화면은 '상태 없음'이라 적는다 —
    #   오류가 안 나는 종류의 잘못이다(`[165]`). 그래서 둘 다 물어본다.
    rec = rec or {}
    state = rec.get("state") or rec.get("상태") or rec.get("진행상태") or ""
    amt = rec.get("supply") or rec.get("금액") or rec.get("공급가액") or ""
    line = "%s — ERP **%s**" % (prj, state or "상태 없음")
    if amt:
        try:
            line += " · 공급가 %s원" % format(int(float(amt)), ",")
        except (TypeError, ValueError):
            line += " · %s" % amt
    if rec.get("cust"):
        line += " · %s" % rec["cust"]
    if rec.get("po"):
        line += " · %s" % rec["po"]
    if verdict:
        line += "\nPO 대조 판정: " + verdict["판정"]
    return {"답": line, "다음": _who_next(state), "근거": _age_note(idx, src("PO대조")),
            "확신": "높"}


def a_invoice(q):
    """"작업은 끝났는데 왜 계산서가 안 나갔나" — 실제로 제일 자주 온 질문이다.

    한 덩어리로 묶어 두면 화면이 **누구에게 넘길지**를 말해 주지 못한다.
    그래서 이유별로 갈라 답한다.
    """
    s, secs = _po_sections()
    if not s["있음"]:
        return None
    lines = ["'미발행' 은 한 가지가 아닙니다 — **갈래마다 할 사람이 다릅니다.**"]
    if secs:
        lines.append("")
        for x in secs:
            lines.append("  · **%s** — %d건" % (x["이름"], x["건수"]))
    # 미청구(A) 는 '진짜 계산서가 안 나간 것'이라 번호를 바로 보여 준다.
    #   나머지 갈래는 수만 말한다 — 다 늘어놓으면 아무도 안 본다(`[170]`).
    a_rows = []
    grab = False
    for line in (s.get("데이터") or "").splitlines():
        if line.startswith("## "):
            grab = line.startswith("## A.")
            continue
        if grab and line.startswith("|") and not line.startswith("|---"):
            cells = [c.strip() for c in line.strip().strip("|").split("|")]
            # 머리글 칸도 'PO' 로 시작한다(`PO번호`). **숫자가 붙은 것만** 번호다 —
            # 안 거르면 목록 맨 앞에 'PO번호' 가 실려 나간다.
            if cells and re.match(r"^PO\d{5,}$", cells[0], re.I):
                a_rows.append(cells[0])
    if a_rows:
        lines += ["", "**A. 미청구(계산서가 정말 안 나간 것)**: " + ", ".join(a_rows[:8])]
    lines += [
        "",
        "구별해야 할 것 — 화면 딱지 `미발행사유` 가 이대로 갈라 줍니다:",
        "  · **발행 대기(ERP 4단계)** — PO 도 왔고 금액도 맞습니다. "
        "류지영이 ERP 에서 넘기기만 하면 됩니다.",
        "  · **ERP 전표 없음** — 발행이 아니라 **등록**부터입니다(다른 사람 몫).",
        "  · **3.오더처리** — 아직 명세서 전입니다.",
    ]
    return {"답": "\n".join(lines),
            "다음": "번호 하나가 궁금하면 그 프로젝트NO 나 PO번호를 그대로 물어보십시오 "
                    "— 그 줄의 판정을 그대로 읽어 드립니다.",
            "근거": _age_note(s), "확신": "높"}


def a_round(q):
    """회차가 안 끝났나 / 왜 이렇게 오래 걸리나."""
    h = src("인계")
    p = src("회차")
    d = (h.get("데이터") or {}).get("일일대조") or {}
    prog = p.get("데이터") or {}
    if not d and not prog:
        return None
    bits = []
    if prog:
        bits.append("지금 단계: **%s** (%s · %s분째)" % (
            prog.get("단계") or "?", prog.get("상태") or "?", prog.get("머문분") or "?"))
    if d:
        if d.get("진행중"):
            bits.append("회차가 %s시간째 돌고 있습니다." % d.get("진행중"))
        if d.get("완주없음"):
            bits.append("마지막 완주가 없습니다 — 회차가 끝을 못 봤습니다.")
        if d.get("실패단계"):
            bits.append("실패한 단계: " + ", ".join(map(str, d["실패단계"])))
    nxt = ("**pid 가 살아 있고 단계가 최근에 바뀌었으면 느린 것이지 멈춘 것이 아닙니다** "
           "— 돌고 있는 회차를 죽이면 그날 대조가 통째로 빕니다. "
           "회차는 예산(기본 150분)을 넘으면 남은 단계를 건너뛰고 스스로 완주합니다.")
    return {"답": "\n".join(bits) or "회차 자국이 아직 없습니다.",
            "다음": nxt, "근거": _age_note(h, p), "확신": "높" if prog else "중"}


def a_queue(q):
    """엑셀에 언제 들어가나 / 몇 건 밀렸나."""
    s = src("큐")
    d = s.get("데이터") or {}
    if not d:
        return None
    txt = "대기 **%s건** · 다음 반영 **%s**" % (d.get("대기"), d.get("다음반영"))
    if d.get("남은분") is not None:
        txt += " (%s분 뒤)" % d.get("남은분")
    if d.get("밀린회차"):
        txt += "\n밀린 회차: " + ", ".join(map(str, d["밀린회차"]))
    return {"답": txt,
            "다음": "엑셀 반영은 11:00·15:00 두 회차뿐입니다. 지금 넣어야 하면 "
                    "앱 [엑셀 반영 예정] 카드의 '지금 바로 엑셀에 반영' 단추를 쓰십시오.",
            "근거": _age_note(s), "확신": "높"}


def a_collect(q):
    """수집이 밀렸나 — 조용한 사고의 대표 격이다."""
    s = src("인계")
    rows = (s.get("데이터") or {}).get("수집신선도") or []
    if not rows:
        return None
    late = [r for r in rows if r.get("밀림")]
    if not late:
        quiet = [r for r in rows if r.get("조용함")]
        return {"답": "밀린 원본이 없습니다(%d갈래 확인)." % len(rows) +
                      ("\n조용한 것: " + " · ".join(
                          "%s — %s" % (r.get("이름"), r.get("조용함")) for r in quiet[:3])
                       if quiet else ""),
                "다음": "'최신 글이 며칠 전'인 것과 '못 모은 글이 있다'는 다른 말입니다.",
                "근거": _age_note(s), "확신": "높"}
    lines = ["★ 밀린 것 %d갈래:" % len(late)]
    for r in late[:5]:
        lines.append("  · %s — 최신 %s (%s일 밀림, 한도 %s일)" % (
            r.get("이름"), r.get("최신"), r.get("밀린일"), r.get("한도")))
    return {"답": "\n".join(lines),
            "다음": "수집은 'CSOS 리서치 및 자료 수집' 세션 몫입니다 — "
                    "여기서 직접 긁지 말고 그쪽에 알리십시오.",
            "근거": _age_note(s), "확신": "높"}


def a_typo(q):
    s = src("오기입")
    d = s.get("데이터") or {}
    if not d:
        return None
    kinds = d.get("종류별") or {}
    return {"답": "원장 %s행 중 오기입 의심 **%s건**%s" % (
                d.get("원장행"), d.get("의심"),
                (" (" + " · ".join("%s %s" % (k, v) for k, v in kinds.items() if v) + ")")
                if any(kinds.values()) else ""),
            "다음": "상세는 `reports/오기입_확인.md`. **자동으로 고치지 않습니다** — "
                    "무엇이 맞는지는 사람만 압니다.",
            "근거": _age_note(s), "확신": "높"}


def a_cancel(q):
    s = src("교차")
    d = s.get("데이터") or {}
    if not d:
        return None
    return {"답": "밴드 신호 %s · 카톡 신호 %s → 짝지어짐 **%s** · 밴드에만 %s · "
                  "**카톡에만 %s**" % (d.get("밴드신호"), d.get("카톡신호"),
                                      d.get("짝지어짐"), d.get("밴드에만"),
                                      d.get("카톡에만")),
            "다음": "마지막 '카톡에만' 이 제일 값어치 있습니다 — **기사에게 안 전달됐을 수 "
                    "있는 것**들입니다. 상세는 `reports/카톡_밴드_교차.md`.",
            "근거": _age_note(s), "확신": "중"}


def a_todo(q):
    """'지금 뭐부터 하면 되나' — 사람이 제일 먼저 묻는 것.

    ★ 2026-08-19 형님 지시("챗봇 헛소리 자꾸 하는데 고도화 작업 진행해")로 **본문을
      갈았다.** 전에는 `다음할일`(= `ecount/AGENTS.md` '대기 항목' 발췌)을 그대로 읊고
      그 위에 **파일 나이**를 '0.0시간 전'이라 찍었다. 실측 그 글은 v420 시절이고
      실제 원장은 v603 이었다 — `session_handoff` 가 30분마다 같은 발췌를 다시 퍼
      담으므로 **파일은 늘 새것이고 내용만 넉 달 묵는다.** 지어내는 것이 아니라
      **낡은 글을 새것이라고 말하는** 쪽이라 더 나쁘다(`[169]` 중 안심시키는 거짓).
      머리글은 실시간 값(반영 대기 484건)인데 본문은 그 blob('반영 대기 0건')이라
      **제 답 안에서 두 숫자가 어긋나기까지** 했다.
    ★ **판정을 새로 만들지 않는다**(`[162]`) — 인계 문서의 '먼저 처리할 것'을 만드는
      그 `session_handoff.blockers()` 를 **그대로 빌린다.** 여기서 다시 세면 같은
      물음에 두 답이 생기고, 갈린 뒤에는 어느 쪽이 맞는지 아무도 모른다.
      그 함수는 넘겨받은 스냅샷만 읽는다 — Z: 를 안 훑는다(`[168]` · 실측 0.001초).
    ★ **갈래는 안 건드렸다**(`[172]`) — '미처리' 질문이 여기로 오는 것은 설계다
      (PROBES). 고친 것은 낡은 내용에 새 시각을 찍던 자리 하나다.
    ★ **발췌를 조용히 빼지도 않는다**(`[169]`) — 있다는 사실과 "이 글이 언제
      것인지는 모른다"를 같이 적고, 출처가 붙은 자리를 가리킨다.
    """
    s = src("인계")
    d = s.get("데이터") or {}
    if not d:
        return None
    bits = []
    if d.get("큐잔량"):
        bits.append("반영 대기 %s건" % d["큐잔량"])
    if d.get("미커밋"):
        bits.append("미커밋 %s" % d["미커밋"])
    if d.get("미푸시"):
        bits.append("미푸시 %s" % d["미푸시"])
    if d.get("점유"):
        bits.append("점유 %d건" % len(d["점유"]))
    if (d.get("앱서버") or {}).get("옛코드"):
        bits.append("★ 앱 서버가 옛 코드로 돌고 있습니다 — `python webapp/restart_server.py`")

    # 인계 문서와 **같은 판정**을 빌린다. 못 부르면 '0건'이 아니라 '못 셈'이다(`[169]`).
    bl, blind = [], ""
    try:
        import session_handoff as _H
        bl = _H.blockers(d)
    except Exception as e:
        blind = "먼저 처리할 것을 **못 셌습니다**(%s) — '없다'는 뜻이 아닙니다." % (
            type(e).__name__)

    lines = []
    if blind:
        lines += [blind, "  `python session_handoff.py --check`"]
    elif bl:
        lines.append("먼저 처리할 것 **%d건** — 위에서부터:" % len(bl))
        for why, how in bl[:4]:
            lines += ["  · %s" % why, "    `%s`" % how]
        if len(bl) > 4:
            lines.append("  · …그 밖 %d건 — `python session_handoff.py --check`"
                         % (len(bl) - 4))
    else:
        lines.append("먼저 처리할 것이 없습니다 — 새 작업을 시작해도 됩니다.")

    # AGENTS.md 발췌: 빼지 않되 **새것이라고도 말하지 않는다**.
    nxt = [x for x in (d.get("다음할일") or []) if str(x).strip()]
    if nxt:
        lines += ["", "참고: `ecount/AGENTS.md` '대기 항목' 발췌도 %d줄 있습니다 — "
                      "**그 글이 언제 것인지는 이 파일이 말해 주지 않습니다**"
                      "(파일 나이 ≠ 내용 나이). 출처가 붙은 자리에서 보십시오: "
                      "`python session_handoff.py --check` 의 "
                      "'## 다음 할 일 (AGENTS.md 발췌)'." % len(nxt)]

    if blind:
        head = "먼저 처리할 것을 못 셌습니다"
    else:
        head = "먼저 처리할 것 **%d건**" % len(bl)
    if bits:
        head += " · 걸린 것: " + " · ".join(bits)
    return {"답": head, "다음": "\n".join(lines),
            # 근거 시각은 이제 **본문이 실제로 읽은 값**의 나이다 — 30분마다 다시
            # 계산되는 스냅샷이라 파일 나이가 곧 내용 나이다.
            "근거": _age_note(s), "확신": "중" if blind else "높"}


def a_server(q):
    s = src("인계")
    d = (s.get("데이터") or {}).get("앱서버") or {}
    if not d:
        return None
    if d.get("옛코드"):
        return {"답": "★ 앱 서버가 **옛 코드**로 돌고 있습니다. 화면이 멀쩡히 숫자를 "
                      "보여 줘도 그건 고치기 전 코드의 답입니다.",
                "다음": "`python webapp/restart_server.py` 한 줄입니다.",
                "근거": _age_note(s), "확신": "높"}
    # 서버가 최신인데도 "왜 안 바뀌나"를 물었다면 남는 원인 둘을 마저 답한다 —
    # ① 화면 캐시(갱신 단추) ② 엑셀 보관본은 회차 전이라 아직 옛 시점이다.
    nxt = "화면 위 갱신 단추(또는 새로고침)를 눌러 보십시오 — 화면이 옛 값을 들고 있을 수 있습니다."
    qd = (src("큐").get("데이터") or {})
    if qd.get("대기"):
        nxt += ("\n엑셀(보관본)을 보고 있다면 아직 회차 전입니다 — 대기 %s건, 다음 생성 %s. "
                "정본(앱 DB)에는 이미 저장돼 있습니다." % (qd.get("대기"), qd.get("다음반영")))
    return {"답": "앱 서버는 최신 코드로 돌고 있습니다.", "다음": nxt,
            "근거": _age_note(s), "확신": "높"}


# 말투는 **넉넉하게** 잡는다. 사람은 같은 것을 열 가지로 묻는다.
def a_input_store(q):
    """앱에 입력하면 어디에 저장되나 — 정본 확정(2026-08-10·11) 뒤 가장 잦은 혼동.

    실제 기록: "여기에 입력하면 db에 저장되나?" 가 모름으로 떨어졌다. 답의 근거는
    정본 규칙 문서라 정적이지만, 보관본 대기 건수는 큐에서 살아 있는 값을 붙인다."""
    txt = ("**예 — 저장 버튼 응답 전에 앱 뒤 SQLite 정본에 즉시 저장됩니다**"
           "(감사로그와 한 트랜잭션·멱등키).\n"
           "엑셀(관리대장)은 그 DB 에서 밖으로만 만드는 **읽기 전용 보관본**입니다 — "
           "11:00·15:00 회차(또는 사람 지시 즉시 생성)가 만듭니다.\n"
           "반대로 **엑셀에 손으로 적은 값은 정본에 들어가지 않습니다**"
           "(2026-08-11 컷오버 — 손입력은 감지해 앱 재입력을 안내합니다).")
    s = src("큐")
    d = s.get("데이터") or {}
    if d.get("대기") is not None:
        txt += "\n지금 보관본 대기 **%s건** · 다음 생성 %s." % (d.get("대기"), d.get("다음반영"))
    return {"답": txt,
            "다음": "방금 저장한 건이 화면에 없으면 갱신 단추를 누르십시오. 그래도 없으면 "
                    "저장이 실패한 것입니다 — 다시 입력하고, 반복되면 클로드에 알리십시오.",
            "근거": (_age_note(s) if d else "근거: CLAUDE.md 정본 규칙(2026-08-10·11 확정)"),
            "확신": "높"}


INTENTS = [
    ("프로젝트조회", [r"[A-Z]{2}\d{7}", r"PO\s?\d{6}"], a_project),
    # 말투를 좁게 잡으면 **제일 자주 오는 질문이 조용히 모름으로 떨어진다.**
    # 실제로 그랬다 — "계산서 발행이 안된거지" 는 '안'이 '발행' **뒤에** 온다.
    # 그래서 순서를 강요하지 않는다: 계산서와 부정어가 **한 문장에 같이** 있으면 잡는다.
    ("계산서미발행", [r"(세금)?계산서.{0,20}(안|못|미|왜|언제)",
                      r"(안|못|미|왜).{0,20}(세금)?계산서",
                      r"미발행", r"발행.{0,10}(안|못)\s*(됐|됨|되|나)"], a_invoice),
    ("회차상태", [r"미완주", r"회차.*(안|왜|멈|늦|오래)", r"대조.*(안 돌|멈)",
                  r"몇 시간째", r"daily.?run"], a_round),
    ("엑셀반영", [r"(엑셀|원장).*(반영|들어가|언제)", r"큐.*(몇|얼마|남)",
                  r"반영.*(대기|언제)", r"대기.*건"], a_queue),
    # 2026-08-11 기록에서: "여기에 입력하면 db에 저장되나?" 가 모름으로 떨어졌다.
    # 정본 컷오버 뒤 가장 잦은 혼동이라 갈래를 준다. 엑셀반영 **뒤**에 둔다 —
    # 엑셀 낱말이 있는 저장 질문은 위 갈래가 먼저 잡는 것이 맞다.
    ("입력저장", [r"(입력|저장|적).{0,14}(db|DB|디비|정본|어디)",
                  r"(db|DB|디비).{0,10}(저장|들어가|기록)", r"저장되나"], a_input_store),
    ("수집밀림", [r"수집.*(밀|안 들어|늦)", r"밴드.*(밀|최신|언제까지)",
                  r"자료.*(최신|밀)", r"원본.*(밀|안 들어)"], a_collect),
    ("오기입", [r"오타", r"오기입", r"잘못 적", r"틀린 (값|번호)"], a_typo),
    ("취소", [r"취소", r"연기", r"재방문", r"카톡.*밴드", r"교차"], a_cancel),
    # 실측 "근데 왜 안바껴있어"(2026-08-11) — 구어체 '바껴'·붙여쓰기까지 넉넉히 잡는다.
    ("앱서버", [r"(앱|서버).*(옛|낡|안 바뀌|재시작)", r"화면이 안 바뀌",
               r"안 ?바(뀌|껴)", r"왜 ?그대로"], a_server),
    ("지금할일", [r"(뭐|무엇).*(하면|해야|먼저)", r"할 일", r"상태.*(어때|어떻)",
                  r"먼저 처리", r"미처리", r"안 ?끝난.{0,6}(정리|뭐)"], a_todo),
]


def classify(q):
    """어느 규칙의 질문인가. 여러 개가 걸리면 **먼저 선언된 것**이 이긴다.

    번호가 든 질문이 맨 앞인 이유가 그것이다 — 번호는 사람이 이미 무엇을
    볼지 정해 준 것이므로, 낱말로 미루어 짐작한 것보다 세다.
    """
    t = q or ""
    for name, pats, fn in INTENTS:
        for p in pats:
            if re.search(p, t, re.I):
                return name, fn
    return None, None


# ─────────────────────────────────────────────────────────────────────────────
# 못 답할 때 — 여기가 진짜 절약이다
# ─────────────────────────────────────────────────────────────────────────────
def escalation(q, intent=None):
    """클로드에게 **그대로 붙여넣을** 문구를 만든다.

    답을 못 하는 것과 사람을 빈손으로 보내는 것은 다르다. 클로드가 어차피 열어 볼
    파일과 지금 수치를 **미리 붙여** 주면, 열몇 번 걸릴 왕복이 한 번으로 준다.
    못 답한 질문에서도 크레딧이 절약되는 유일한 길이다.
    """
    facts = []
    for name in ("인계", "큐", "회차"):
        s = src(name)
        if not s["있음"]:
            continue
        d = s.get("데이터") or {}
        if name == "큐" and isinstance(d, dict):
            facts.append("- 반영 대기 %s건, 다음 반영 %s" % (d.get("대기"), d.get("다음반영")))
        elif name == "회차" and isinstance(d, dict):
            facts.append("- 회차 단계 '%s' (%s)" % (d.get("단계"), d.get("상태")))
        elif name == "인계" and isinstance(d, dict):
            facts.append("- 관리대장/큐잔량 %s · 점유 %d건 · 미커밋 %s"
                         % (d.get("큐잔량"), len(d.get("점유") or []), d.get("미커밋")))
        facts.append("  (근거 %s, %s시간 전)" % (s["파일"], s["나이시간"]))
    body = [
        "%s" % (q or "").strip(),
        "",
        "— 앱이 먼저 확인한 것(다시 뒤지지 마십시오):",
    ] + (facts or ["- (리포트가 아직 없습니다)"]) + [
        "",
        "관련 리포트: reports/세션인계.md · reports/오기입_확인.md · "
        "reports/카톡_밴드_교차.md · 최신 PO대조_*.md",
    ]
    if intent:
        body.insert(1, "(앱 분류: %s — 규칙은 있는데 근거가 모자랐습니다)" % intent)
    return "\n".join(body)


def ask(q, log=True):
    """질문 하나에 답한다. **예외를 밖으로 내보내지 않는다.**"""
    q = (q or "").strip()
    t0 = datetime.now()
    intent, fn = classify(q)
    res, err = None, None
    if fn is not None:
        try:
            res = fn(q)
        except Exception as e:      # 규칙 하나가 깨져도 답변기는 산다
            err = "%s: %s" % (type(e).__name__, e)
    ms = int((datetime.now() - t0).total_seconds() * 1000)
    if res:
        out = {"질문": q, "분류": intent, "답함": True, "답": res.get("답", ""),
               "다음": res.get("다음", ""), "근거": res.get("근거", ""),
               "확신": res.get("확신", "중"), "ms": ms, "클로드문구": ""}
    else:
        out = {"질문": q, "분류": intent, "답함": False,
               "답": ("규칙은 있는데 근거 파일이 모자랍니다." if intent
                      else "이 모양의 질문은 아직 규칙이 없습니다."),
               "다음": "아래 문구를 클로드에게 그대로 붙여넣으십시오 — "
                       "앱이 확인한 사실이 이미 들어 있어 왕복이 짧아집니다.",
               "근거": "", "확신": "없음", "ms": ms,
               "클로드문구": escalation(q, intent)}
    if err:
        out["오류"] = err
    if log:
        log_ask(out)
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 기록 — 얼마나 아꼈나, 다음에 무엇을 만들면 되나
# ─────────────────────────────────────────────────────────────────────────────
def log_ask(out):
    """기록은 **실패해도 조용히 넘어간다** — 기록하려다 답을 못 주면 본말전도다."""
    try:
        os.makedirs(REPORT_DIR, exist_ok=True)
        rows = []
        if os.path.exists(LOG):
            try:
                rows = json.load(io.open(LOG, encoding="utf-8")).get("기록") or []
            except Exception:
                rows = []
        rows.append({"때": datetime.now().isoformat(timespec="seconds"),
                     "질문": out.get("질문", "")[:200],
                     "분류": out.get("분류"), "답함": bool(out.get("답함")),
                     "ms": out.get("ms")})
        rows = rows[-LOG_MAX:]
        tmp = LOG + ".tmp"
        with io.open(tmp, "w", encoding="utf-8") as f:
            json.dump({"갱신": datetime.now().isoformat(timespec="seconds"),
                       "기록": rows}, f, ensure_ascii=False, indent=1)
        os.replace(tmp, LOG)
    except Exception:
        pass


def stats():
    """앱이 몇 %를 답했나 · 못 답한 질문은 어떤 모양인가.

    뒤엣것이 **다음에 만들 규칙의 목록**이다. 추측이 아니라 기록에서 나온다.
    """
    rows = []
    if os.path.exists(LOG):
        try:
            rows = json.load(io.open(LOG, encoding="utf-8")).get("기록") or []
        except Exception:
            rows = []
    n = len(rows)
    ok = sum(1 for r in rows if r.get("답함"))
    miss = {}
    for r in rows:
        if not r.get("답함"):
            key = r.get("분류") or "(분류 없음)"
            miss[key] = miss.get(key, 0) + 1
    return {"질문수": n, "앱이답함": ok,
            "비율": (round(ok * 100.0 / n, 1) if n else 0.0),
            "못답한갈래": sorted(miss.items(), key=lambda kv: -kv[1])[:10],
            "평균ms": (round(sum(r.get("ms") or 0 for r in rows) / n, 1) if n else 0)}


# 자가점검이 볼 것 — **어느 갈래로 가야 하는지까지** 적는다.
#
# ★ 처음 만든 자가점검은 "터지지 않으면 통과"였다. 그래서 "왜 계산서 발행이
#   안된거지"(이 프로젝트에서 제일 자주 온 질문)가 **모름으로 떨어졌는데도
#   '모두 통과'라고 말했다.** 계기가 0을 내면 아무도 의심하지 않는다(`[169]`).
#   그러니 기대하는 갈래를 적어 두고, 거기로 안 가면 실패라고 말한다.
# `None` 은 **모름이어야 맞는** 질문이다 — 규칙이 아무 데나 걸리는 것도 고장이다.
PROBES = [
    ("UJ2600021", "프로젝트조회"),
    ("PO327948 어떻게 됐어", "프로젝트조회"),
    ("왜 계산서 발행이 안된거지", "계산서미발행"),
    ("작업 끝났는데 세금계산서가 왜 안 나가", "계산서미발행"),
    ("회차가 왜 미완주야", "회차상태"),
    ("엑셀 언제 반영돼", "엑셀반영"),
    ("수집 밀렸어?", "수집밀림"),
    ("오타 있어?", "오기입"),
    ("취소된 거 있나", "취소"),
    ("지금 뭐부터 하면 돼", "지금할일"),
    ("지금 미처리된건 정리", "지금할일"),
    # 2026-08-19 형님이 실제로 물어 헛소리를 받은 그 문장이다 — 갈래가
    # 바뀌면 이 고침이 통째로 안 걸린다. 내용까지는 `[320]` 이 잰다.
    ("미처리건은 어디서 확인해?", "지금할일"),
    ("여기에 입력하면 db에 저장되나?", "입력저장"),
    ("앱 화면이 안 바뀌는데", "앱서버"),
    ("근데 왜 안바껴있어", "앱서버"),
    ("김밥천국 메뉴 알려줘", None),
    ("오늘 날씨 어때", None),
]


def selftest():
    """규칙이 살아 있나.

    보는 것 셋: ① 터지지 않는다 ② **기대한 갈래로 간다** ③ 못 답하면
    반드시 클로드 문구가 있다(빈손으로 돌려보내지 않는다).
    """
    bad = []
    for q, want in PROBES:
        try:
            r = ask(q, log=False)
        except Exception as e:
            bad.append("%s → 터짐(%s)" % (q, e))
            continue
        if not isinstance(r, dict) or "답" not in r:
            bad.append("%s → 모양이 이상하다" % q)
            continue
        if r.get("분류") != want:
            bad.append("%s → 갈래 '%s' (기대 '%s')" % (q, r.get("분류"), want))
        if not r["답함"] and not r.get("클로드문구"):
            bad.append("%s → 못 답했는데 클로드 문구도 없다" % q)
    return bad


# ─────────────────────────────────────────────────────────────────────────────
# PC 가 꺼져 있어도 답한다 — 미리 만든 답 꾸러미
#
# 사용자 지시(2026-08-09): "만약 이 컴퓨터가 꺼져있어 연결되어있지 않더라도
# 앱 자체적으로 처리할 수 있는 알고리즘 구현해"
#
# ★ **규칙을 JS 로 옮겨 적지 않는다.** 같은 판단을 두 곳에서 하면 언젠가 갈리고,
#   갈린 뒤에는 어느 쪽이 맞는지 아무도 모른다(이 프로젝트가 여러 번 겪은 일이다).
#   그래서 파이썬이 **답을 미리 만들어** 싣고, JS 는 그것을 **읽기만** 한다.
#   분류 말투마저 데이터로 실어 보낸다 — 규칙이 사는 곳은 끝까지 이 파일 하나다.
#
# ★ 그리고 **언제 만든 답인지 반드시 같이 싣는다.** 꾸러미는 PC 가 꺼져 있는 동안
#   낡는다. 낡은 답을 시각 없이 보여 주는 것이 여기서 제일 위험하다 —
#   화면은 멀쩡한데 어제 상태를 오늘 것처럼 말한다.
# ─────────────────────────────────────────────────────────────────────────────
PACK_MAX_NUMBERS = 4000       # 번호 표 상한. 넘으면 **몇 개를 뺐는지 적어 싣는다**


def answer_pack():
    """폰이 PC 없이 쓸 답 꾸러미. `cloud_publish` 가 잠긴 사본에 넣는다."""
    made = datetime.now()

    # ① 전체 질문 답 — 어차피 데이터 전체를 보는 답이라 미리 만들어 두면 그만이다.
    answers = {}
    for name, _pats, fn in INTENTS:
        if name == "프로젝트조회":
            continue                      # 번호마다 다르므로 ② 로 간다
        try:
            r = fn("")
        except Exception:
            r = None
        answers[name] = ({"답": r.get("답", ""), "다음": r.get("다음", ""),
                          "근거": r.get("근거", ""), "확신": r.get("확신", "중")}
                         if r else None)

    # ② 번호 표 — 프로젝트NO·PO번호로 물었을 때의 답을 **문장까지 만들어** 둔다.
    #    여기서 문장을 안 만들고 값만 실으면 JS 가 `_who_next` 를 다시 써야 한다.
    nums, dropped = {}, 0
    idx = src("ERP색인")
    data = idx.get("데이터") or {}
    if isinstance(data, dict) and "index" in data:
        data = data.get("index") or {}
    for prj in sorted(data if isinstance(data, dict) else {}):
        if len(nums) >= PACK_MAX_NUMBERS:
            dropped += 1
            continue
        try:
            r = a_project(prj)
        except Exception:
            r = None
        if not r:
            continue
        nums[prj] = {"답": r.get("답", ""), "다음": r.get("다음", ""),
                     "근거": r.get("근거", "")}
        # 같은 답을 PO 번호로도 찾을 수 있게 이름표를 더 단다(표를 두 벌 만들지 않는다).
        for m in PO_PAT.finditer(str((data.get(prj) or {}).get("po") or "")):
            nums.setdefault("PO" + m.group(1), {"참조": prj})

    return {
        "만든때": made.isoformat(timespec="seconds"),
        # 분류 말투를 **데이터로** 보낸다 — JS 가 규칙을 다시 쓰지 않게.
        "규칙": [{"이름": n, "말투": list(p)} for n, p, _f in INTENTS],
        "답": answers,
        "번호": nums,
        "번호뺀수": dropped,          # 조용히 자르지 않는다
        "안내": ("이 답은 PC 가 마지막으로 만든 것입니다. PC 가 꺼져 있는 동안에는 "
                 "낡을 수 있으니 만든 시각을 함께 보십시오."),
        # 못 답할 때 폰이 만들어 줄 문구의 머리말. 규칙이 없는 질문도 빈손으로
        # 돌려보내지 않는다 — 여기서도 왕복을 짧게 만드는 것이 요점이다.
        # ★ 번호를 알아보는 말투도 **여기서 준다.** 폰이 제 손으로 적으면 그것도 사본이고,
        #   사본은 언젠가 갈린다. 번호 모양이 바뀌는 날 폰만 못 알아보게 된다.
        "번호말투": {"프로젝트": PRJ_PAT.pattern, "PO": PO_PAT.pattern},
        "탈출머리말": "— 앱(폰 사본)이 먼저 확인한 것:",
    }


def main():
    import sys
    # 윈도우 콘솔(cp949)에서 — 같은 문자가 UnicodeEncodeError 로 죽지 않게 한다
    # (실사고 2026-08-11: --stats 가 요약 첫 줄만 찍고 죽었다). 인코딩은 그대로 두고
    # 못 찍는 글자만 ?로 바꾼다 — utf-8 강제는 한글 콘솔을 통째로 깨뜨린다.
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print(__doc__.strip().splitlines()[0])
        print('쓰기: python local_ai.py "질문"  |  --stats  |  --selftest')
        return 0
    if args[0] == "--stats":
        s = stats()
        print("질문 %d건 · 앱이 답함 %d건 (%.1f%%) · 평균 %sms"
              % (s["질문수"], s["앱이답함"], s["비율"], s["평균ms"]))
        if s["못답한갈래"]:
            print("아직 규칙이 없는 갈래(다음에 만들 것):")
            for k, v in s["못답한갈래"]:
                print("  · %s — %d회" % (k, v))
        return 0
    if args[0] == "--pack":
        pk = answer_pack()
        print("답 꾸러미: 규칙 %d · 미리만든답 %d · 번호 %d%s (만든때 %s)"
              % (len(pk["규칙"]), sum(1 for v in pk["답"].values() if v),
                 len(pk["번호"]), (" · 뺀 것 %d" % pk["번호뺀수"]) if pk["번호뺀수"] else "",
                 pk["만든때"]))
        return 0
    if args[0] == "--selftest":
        bad = selftest()
        print("자가점검: " + ("모두 통과" if not bad else "문제 %d건 — %s"
                              % (len(bad), bad)))
        return 1 if bad else 0
    r = ask(" ".join(args))
    print(("[%s]" % r["분류"]) if r["분류"] else "[모름]", "확신:", r["확신"])
    print(r["답"])
    if r.get("다음"):
        print("\n▶ " + r["다음"])
    if r.get("근거"):
        print("  " + r["근거"])
    if r.get("클로드문구"):
        print("\n" + "-" * 60 + "\n" + r["클로드문구"] + "\n" + "-" * 60)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
