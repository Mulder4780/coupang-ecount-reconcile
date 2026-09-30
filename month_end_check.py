# -*- coding: utf-8 -*-
"""월말 점검 — **적기에 했나** 와 **매출로 이어졌나** 를 한 장에서 답한다.

2026-09-28 김형래 이사(CFO) 지시:
  "중요한 것은 정기점검 및 돌발 AS가 요구 일정에 맞춰 적기에 수행되고, 지연 없이
   매출 발생으로 연결되고 있는지를 지속적으로 관리하는 것입니다.
   월말 매출 누락이나 지연사항이 없는지 점검 바랍니다."

★ 왜 새로 만드는가 — 두 물음에 답할 재료는 **이미 다 있었다**(정기점검 기한 ·
  청구 사다리 · 세금계산서 미발행 경과). 없던 것은 그 둘을 **나란히 놓고 보는
  자리**다. 따로 보면 "점검은 다 했는데 돈이 왜 안 들어오지" 를 아무도 안 묻는다.

★ 판정을 새로 만들지 않는다([162]). 여기는 **모아서 보여 주기만** 한다:
  · 정기점검 기한 → `work_flow.pm_due` (분기 기준 · 2026-09-28 형님 지시)
  · 청구 사다리  → 앱 DB `정산` 갈래의 상태 낱말 그대로(낱말을 짓지 않는다 · [166])
  · 계산서 미발행 → `reports/세금계산서_미발행_경과.json`
  여기서 다시 판정하면 같은 물음에 두 답이 나온다.

★ 읽기만 한다. 엑셀도 ERP 도 한 글자 안 고친다 — 발행·청구는 사람 몫이다.

★ 못 읽은 것을 '없음'이라 적지 않는다([169]). 근거를 못 읽으면 **'확인 못함'** 으로
  남긴다 — 0건과 '안 봤다'는 다른 사실이다.

  python month_end_check.py                # 이번 달
  python month_end_check.py --month 2026-08
  python month_end_check.py --md           # reports/월말점검_YYYYMM.md 로도 남긴다
"""
from __future__ import annotations

import argparse
import io
import json
import os
import sqlite3
import sys
from datetime import date, datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import work_flow

DB = os.path.join(ROOT, "db", "app_store.db")
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")

# 돌발AS 가 '늦었다'고 볼 날수. 대표 보고 기준과 같은 값을 쓴다 — 여기서 새로 정하면
# 같은 건이 화면마다 다른 색으로 보인다([162]).
AS_DELAY_DAYS = int(os.environ.get("COUPANG_AS_DELAY_DAYS") or 3)

# 작업이 끝난 뒤 청구가 이만큼 안 움직이면 '지연'이라 적는다. 넉넉히 잡는다 —
# 경보가 대부분이면 아무도 안 본다([170]).
#
# ★ 이 값은 **'정산 기록이 없다'에도 같이 걸린다.** 정산 행은 작업이 끝난 뒤에
#   만들어지므로, 이달 완료를 그날로 '매출 누락'이라 부르면 **매달 수십 건이
#   거짓으로** 뜬다(실측 2026-09-28: 그렇게 세면 91건인데 전부 이달 완료 건이었다).
BILL_DELAY_DAYS = int(os.environ.get("COUPANG_BILL_DELAY_DAYS") or 30)

# 돌발AS 지연을 **두 갈래로** 가르는 값. 조치가 다르기 때문이다([289]):
#   · 90일 이내 → 진짜 현장 지연일 수 있다. 기사·일정을 봐야 한다.
#   · 90일 초과 → 접수가 2023~2025년인데 아직 '신규접수'다. 현장이 밀린 것이 아니라
#     **완료 기록이 안 들어온 것**일 가능성이 크다(실측 2026-09-28: 104건 중 79건).
#     뭉쳐서 '지연 104건'이라 보고하면 사람이 엉뚱한 데를 고치러 간다([172]).
#   ★ 어느 쪽인지 **단정하지 않는다**([169]) — 갈라서 보여 주고 사람이 판단한다.
AS_STALE_DAYS = int(os.environ.get("COUPANG_AS_STALE_DAYS") or 90)

# 청구 사다리 — 관리대장 드롭다운 낱말 그대로다(billing_status.py 와 같은 사다리).
# 아직 돈이 안 된 단계가 앞의 넷이다.
NOT_YET_BILLED = ("", "작업완료", "금액확정", "거래명세서발행")

# 일이 더 갈 곳이 아니라서 세지 않는 상태. 낱말은 work_flow.CANCELLED 에서 빌린다.
PM_SKIP = ("AS전환", "점검불가")


def _rows(kind):
    """앱 DB(정본)에서 한 갈래를 읽는다. 칸은 work_field 에 흩어져 있다.

    못 읽으면 **None** 이다 — 빈 목록으로 돌려주면 '0건'과 구별이 안 된다([169])."""
    if not os.path.exists(DB):
        return None
    try:
        c = sqlite3.connect("file:%s?mode=ro" % DB.replace("\\", "/"), uri=True)
    except Exception:
        return None
    try:
        out = {}
        for wid, st, pno, camp in c.execute(
                "select id,status,project_no,camp_name from work_item "
                "where kind=? and deleted_at is null", (kind,)):
            out[wid] = {"_id": wid, "상태": st or "", "프로젝트NO": pno or "",
                        "캠프명": camp or ""}
        if out:
            qs = ",".join("?" * len(out))
            for wid, k, v in c.execute(
                    "select work_id,field_key,value_json from work_field "
                    "where work_id in (%s)" % qs, tuple(out)):
                try:
                    v = json.loads(v)
                except Exception:
                    pass
                if v not in (None, ""):
                    out[wid][k] = v
        return list(out.values())
    except Exception:
        return None
    finally:
        c.close()


def _d(v):
    return str(v or "").strip()[:10]


def _days(a, b):
    """두 날짜 사이 날수. 못 읽으면 0 — 없는 숫자를 지어내지 않는다([169])."""
    try:
        return (datetime.strptime(b[:10], "%Y-%m-%d")
                - datetime.strptime(a[:10], "%Y-%m-%d")).days
    except Exception:
        return 0


def _in_month(v, month):
    return _d(v)[:7] == month


def erp_projects():
    """ERP 판매 자료에 실제로 찍힌 프로젝트번호 집합. **못 읽으면 None**([169]).

    판정을 새로 만들지 않는다([162]) — `erp_sales_index` 가 만든 정본 색인을 읽기만 한다.
    빈 집합과 None 은 다른 사실이다: 빈 집합은 '한 건도 없다', None 은 '안 봤다'이고
    부르는 쪽이 물러날지 말지를 그것으로 정한다.
    """
    try:
        import erp_sales_index as E
        d = json.load(io.open(E.canon_index_path(), encoding="utf-8"))
        idx = d.get("index") or {}
        if not isinstance(idx, dict):
            return None
        return set(idx)
    except Exception:
        return None


def _skip_pm(r):
    st = str(r.get("점검상태") or "").strip()
    return st in PM_SKIP or work_flow.is_cancelled(r, "pm")


def on_time(month, today):
    """① 요구 일정에 맞춰 했나 — 정기점검은 **분기 기한**, 돌발AS는 접수 경과일."""
    out = {"정기점검": {}, "돌발AS": {}, "확인못함": []}

    pm = _rows("정기점검")
    if pm is None:
        out["확인못함"].append("앱 DB를 못 읽어 정기점검을 세지 못했습니다")
    else:
        over, inq, done = [], [], 0
        for r in pm:
            real = _d(r.get("실제점검일"))
            if real:
                if _in_month(real, month):
                    done += 1
                continue
            if _skip_pm(r):
                continue
            due = work_flow.pm_due(r.get("점검예정일"), today)
            item = {"프로젝트NO": r.get("프로젝트NO"), "캠프명": r.get("캠프명"),
                    "점검예정일": _d(r.get("점검예정일")), "기한": due["기한"],
                    "분기": due["분기"], "경과일": due["경과일"], "근거": due["근거"]}
            if due["갈래"] == "경과":
                over.append(item)
            elif due["갈래"] == "분기내":
                inq.append(item)
        over.sort(key=lambda x: -x["경과일"])
        out["정기점검"] = {"이달완료": done, "기한경과": over, "분기내": inq,
                       "기준": "분기 기한 (2026-09-28 형님 지시 · 전상희 매니저 보고서)"}

    ass = _rows("돌발AS")
    if ass is None:
        out["확인못함"].append("앱 DB를 못 읽어 돌발AS를 세지 못했습니다")
    else:
        late, done = [], 0
        for r in ass:
            fin = _d(r.get("작업완료일"))
            if fin:
                if _in_month(fin, month):
                    done += 1
                continue
            if work_flow.is_cancelled(r, "as"):
                continue
            got = _d(r.get("접수일자"))
            if not got:
                continue
            n = _days(got, today)
            if n > AS_DELAY_DAYS:
                late.append({"프로젝트NO": r.get("프로젝트NO"), "캠프명": r.get("캠프명"),
                             "접수일자": got, "경과일": n,
                             "진행상태": r.get("진행상태") or "", "사유": r.get("비고") or ""})
        late.sort(key=lambda x: -x["경과일"])
        # 조치가 다르므로 갈라서 준다([289]) — 뭉치면 사람이 엉뚱한 데를 고친다.
        recent = [x for x in late if x["경과일"] <= AS_STALE_DAYS]
        stale = [x for x in late if x["경과일"] > AS_STALE_DAYS]
        out["돌발AS"] = {"이달완료": done, "지연": recent, "오래된미마감": stale,
                      "기준": "접수 후 %d일 초과 미완료 (%d일 넘은 건은 따로 센다)"
                             % (AS_DELAY_DAYS, AS_STALE_DAYS)}
    return out


def revenue_link(month, today):
    """② 지연 없이 매출로 이어졌나 — 끝난 일이 청구 사다리를 올라갔나."""
    out = {"누락후보": [], "청구지연": [], "계산서미발행": None,
           "정산자료밖": [], "정산자료끝": "", "확인못함": []}

    settle = _rows("정산")
    if settle is None:
        out["확인못함"].append("앱 DB를 못 읽어 청구 단계를 세지 못했습니다")
        return out
    step = {}
    for r in settle:
        k = str(r.get("프로젝트NO") or "").strip()
        if k:
            step[k] = r
    # ★★ **없는 것과 안 본 것을 가른다**([169]). 정산 자료는 프로젝트 번호 어디까지만
    #    들어와 있다 — 실측 2026-09-28 에 UJ2601323 이 마지막이었다. 그 뒤 번호를
    #    '매출 누락'이라 부르면 **89건이 통째로 거짓 경보**가 되고, CFO 께 없는
    #    구멍을 보고하는 셈이 된다([172] 틀린 지목은 못 잡는 것보다 나쁘다).
    #    번호는 UJ + 연도 + 일련번호라 글자 순서가 곧 시간 순서다.
    newest = max(step) if step else ""
    out["정산자료끝"] = newest

    # ★★ 2026-09-30 실측으로 **더 나은 근거**를 찾았다 — 정산 기록(관리대장에서 옮겨온
    #    것)은 7~8월에 멈췄지만 **ERP 판매 색인은 오늘까지 산다.** 물어야 할 것은
    #    '관리대장에 적혔나'가 아니라 **'ERP 에 매출로 잡혔나'** 다.
    #    그날 실측: 같은 8·9월을 정산 기준으로 세면 누락 89건인데 ERP 기준으로는 22건이었다.
    #    ⚠ 색인을 **못 읽으면 예전 근거로 물러난다**([169]) — 못 읽었다고 '매출 없음'이라
    #      말하지 않는다. 어느 근거로 판정했는지는 화면이 그대로 적는다.
    erp = erp_projects()
    out["ERP색인"] = {"건수": len(erp) if erp is not None else None,
                    "쓴근거": "ERP 판매 색인" if erp else "정산 기록(관리대장 이관분)"}
    if erp is None:
        out["확인못함"].append(
            "ERP 판매 색인을 못 읽어 정산 기록으로만 판정했습니다 "
            "(python erp_sales_index.py 로 다시 만듭니다)")

    finished = []
    for kind, dk in (("돌발AS", "작업완료일"), ("정기점검", "실제점검일")):
        rs = _rows(kind)
        if rs is None:
            out["확인못함"].append("앱 DB를 못 읽어 %s 완료 건을 세지 못했습니다" % kind)
            continue
        for r in rs:
            if _in_month(r.get(dk), month):
                finished.append((kind, _d(r.get(dk)), r))

    for kind, fin, r in finished:
        # 무상·보험 건은 애초에 매출이 아니다 — 누락이라 부르면 거짓 경보다([170]).
        paid = str(r.get("유상·무상·보험") or "").strip()
        if paid and "유상" not in paid:
            continue
        k = str(r.get("프로젝트NO") or "").strip()
        s = step.get(k)
        n = _days(fin, today)
        common = {"업무": kind, "프로젝트NO": k, "캠프명": r.get("캠프명"),
                  "완료일": fin, "유상무상": paid or "(미기입)", "경과일": n}
        if not s:
            # ERP 색인을 읽을 수 있으면 그것이 더 센 근거다 — 거기에 있으면 매출은 잡혔다.
            if erp is not None:
                if k in erp:
                    continue                      # ERP 에 있다 = 누락이 아니다
                if n > BILL_DELAY_DAYS or True:   # ERP 근거는 갓 끝난 건도 바로 말할 수 있다
                    out["누락후보"].append(dict(
                        common, 왜="ERP 판매 자료에 이 프로젝트가 안 보입니다(완료 %d일째)" % n,
                        근거="ERP 판매 색인"))
                continue
            # 정산 자료가 아직 안 닿은 번호대다 — '없다'가 아니라 '모른다'([169]).
            if newest and k > newest:
                out["정산자료밖"].append(dict(
                    common, 왜="정산 자료가 %s 까지만 들어와 있어 확인할 수 없습니다" % newest))
                continue
            # ★ 정산 행은 작업이 끝난 뒤에 만들어진다. 갓 끝난 건을 '누락'이라 부르면
            #   매달 수십 건이 거짓으로 뜬다([170]) — 기다릴 만큼 기다린 뒤에만 센다.
            if n > BILL_DELAY_DAYS:
                out["누락후보"].append(dict(
                    common, 왜="완료 %d일째인데 정산 기록이 없습니다" % n))
            continue
        st = str(s.get("상태") or "").strip()
        if st in NOT_YET_BILLED:
            if n > BILL_DELAY_DAYS:
                out["청구지연"].append(dict(
                    common, 청구단계=st or "(빈칸)", 경과일=n,
                    왜="완료 %d일째 청구 단계가 '%s' 에 머물러 있습니다" % (n, st or "빈칸")))
    out["누락후보"].sort(key=lambda x: x["완료일"])
    out["청구지연"].sort(key=lambda x: -x["경과일"])
    out["정산자료밖"].sort(key=lambda x: x["완료일"])

    # 계산서 미발행은 이미 세는 곳이 있다 — 빌려 온다([162]).
    p = os.path.join(REPORT_DIR, "세금계산서_미발행_경과.json")
    try:
        d = json.load(io.open(p, encoding="utf-8"))
        out["계산서미발행"] = {"건수": d.get("coupang_total"), "잰때": d.get("generated_at"),
                         "근거": d.get("basis"), "목록": (d.get("rows") or [])[:20]}
    except Exception as e:
        out["확인못함"].append("세금계산서 미발행 경과를 못 읽었습니다 (%s)" % type(e).__name__)
    return out


def build(month=None, today=None):
    today = str(today or date.today().isoformat())[:10]
    month = str(month or today[:7])[:7]
    return {"만든때": datetime.now().isoformat(timespec="seconds"), "달": month,
            "오늘": today,
            "지시": "2026-09-28 김형래 이사(CFO) — 적기 수행 · 지연 없는 매출 연결",
            "적기수행": on_time(month, today), "매출연결": revenue_link(month, today)}


def render(d):
    L = []
    A, B = d["적기수행"], d["매출연결"]
    L.append("# 월말 점검 — %s" % d["달"])
    L.append("")
    L.append("기준일 %s · %s" % (d["오늘"], d["지시"]))
    L.append("")
    L.append("## 1. 적기에 수행되고 있나")
    L.append("")
    pm = A.get("정기점검") or {}
    if pm:
        L.append("**정기점검** — 이달 완료 %d건 · 기한 경과 **%d건** · 분기 안 %d건"
                 % (pm.get("이달완료", 0), len(pm.get("기한경과") or []),
                    len(pm.get("분기내") or [])))
        L.append("")
        L.append("- 기한은 **분기**입니다. 예정일 하루가 지난 것은 지연이 아닙니다 (%s)."
                 % pm.get("기준", ""))
        for x in (pm.get("기한경과") or [])[:15]:
            L.append("  - ★ %s %s — %s 기한 %s, %d일 경과"
                     % (x["프로젝트NO"], x["캠프명"], x["분기"], x["기한"], x["경과일"]))
        if pm.get("분기내"):
            L.append("  - (참고) 분기 안이라 아직 기한 내인 건 %d개 — 경고 아님"
                     % len(pm["분기내"]))
        L.append("")
    a = A.get("돌발AS") or {}
    if a:
        stale = a.get("오래된미마감") or []
        L.append("**돌발AS** — 이달 완료 %d건 · 지연 **%d건** (%s)"
                 % (a.get("이달완료", 0), len(a.get("지연") or []), a.get("기준", "")))
        for x in (a.get("지연") or [])[:15]:
            L.append("  - ★ %s %s — 접수 %s, %d일째 미완료 (%s)"
                     % (x["프로젝트NO"], x["캠프명"], x["접수일자"], x["경과일"],
                        x["진행상태"] or "상태 미기입"))
        L.append("")
        if stale:
            L.append("**오래된 미마감 %d건** — 접수가 %d일을 넘었는데 아직 '접수' 상태입니다."
                     % (len(stale), AS_STALE_DAYS))
            L.append("  현장이 밀린 것인지 **완료 기록만 안 들어온 것**인지 갈라야 합니다 — "
                     "여기서는 단정하지 않습니다.")
            for x in stale[:10]:
                L.append("  - %s %s — 접수 %s, %d일째 (%s)"
                         % (x["프로젝트NO"], x["캠프명"] or "(캠프명 없음)", x["접수일자"],
                            x["경과일"], x["진행상태"] or "상태 미기입"))
            if len(stale) > 10:
                L.append("  - … 그 밖 %d건" % (len(stale) - 10))
            L.append("")
    for w in A.get("확인못함") or []:
        L.append("- (확인 못함) %s" % w)
    L.append("")
    L.append("## 2. 지연 없이 매출로 이어지고 있나")
    L.append("")
    miss = B.get("누락후보") or []
    근거 = (B.get("ERP색인") or {}).get("쓴근거") or "정산 기록"
    if 근거.startswith("ERP"):
        L.append("**매출 누락 후보 %d건** — 완료했는데 **ERP 판매 자료에 안 보입니다**." % len(miss))
        L.append("  (근거: %s · %s건) — 'ERP에 안 보인다'는 '매출이 없다'가 아니라 "
                 "받아 둔 화면에 안 찍혔다는 뜻입니다. 확정은 담당자 확인 뒤에 합니다."
                 % (근거, (B.get("ERP색인") or {}).get("건수")))
    else:
        L.append("**매출 누락 후보 %d건** — 완료 %d일이 넘었는데 정산 기록이 없습니다."
                 % (len(miss), BILL_DELAY_DAYS))
    if not miss:
        L.append("  (갓 끝난 건은 정산이 아직 안 만들어진 것이라 여기 안 셉니다.)")
    for x in miss[:15]:
        L.append("  - ★ %s %s %s — 완료 %s, %d일 경과 · %s"
                 % (x["업무"], x["프로젝트NO"], x["캠프명"], x["완료일"], x["경과일"],
                    x["유상무상"]))
    L.append("")
    out_of = B.get("정산자료밖") or []
    if out_of:
        L.append("**확인 못한 %d건** — 정산 자료가 %s 까지만 들어와 있습니다."
                 % (len(out_of), B.get("정산자료끝") or "?"))
        L.append("  이 건들은 **매출이 빠진 것이 아니라 아직 대조를 못 한 것**입니다. "
                 "정산 자료를 최신까지 받으면 갈립니다.")
        L.append("")
    L.append("**청구 지연 %d건** — 완료 %d일이 넘도록 청구 단계가 안 올라갔습니다."
             % (len(B.get("청구지연") or []), BILL_DELAY_DAYS))
    for x in (B.get("청구지연") or [])[:15]:
        L.append("  - ★ %s %s %s — %s · %d일 경과"
                 % (x["업무"], x["프로젝트NO"], x["캠프명"], x["청구단계"], x["경과일"]))
    L.append("")
    ti = B.get("계산서미발행")
    if ti:
        L.append("**세금계산서 미발행 %s건** (%s 기준)" % (ti.get("건수"), ti.get("잰때")))
        for r in (ti.get("목록") or [])[:10]:
            L.append("  - %s %s %s원 — %s"
                     % (r.get("일자"), r.get("전표번호"),
                        format(int(r.get("금액") or 0), ","), r.get("적요") or ""))
        L.append("")
    for w in B.get("확인못함") or []:
        L.append("- (확인 못함) %s" % w)
    L.append("")
    L.append("---")
    L.append("이 표는 읽기만 합니다. 발행·청구는 담당자가 ERP 에서 직접 합니다.")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser(description="월말 점검 — 적기 수행 · 매출 연결")
    ap.add_argument("--month", help="YYYY-MM (기본: 이번 달)")
    ap.add_argument("--today", help="기준일 (시험용)")
    ap.add_argument("--json", action="store_true", help="JSON 으로 낸다")
    ap.add_argument("--md", action="store_true", help="reports/ 에 마크다운으로도 남긴다")
    a = ap.parse_args()
    d = build(a.month, a.today)
    if a.json:
        print(json.dumps(d, ensure_ascii=False, indent=2))
        return 0
    txt = render(d)
    print(txt)
    if a.md:
        os.makedirs(REPORT_DIR, exist_ok=True)
        p = os.path.join(REPORT_DIR, "월말점검_%s.md" % d["달"].replace("-", ""))
        io.open(p, "w", encoding="utf-8", newline="").write(txt)
        print("\n남겼습니다: %s" % p)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
