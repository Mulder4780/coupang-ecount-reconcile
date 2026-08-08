# -*- coding: utf-8 -*-
"""
po_reconcile.py — 쿠팡 발행 PO ↔ 관리대장 자동 대조
=====================================================
쿠팡(아리바/포털)에서 내려받은 PO 목록 엑셀을 inbox/에 넣으면(파일명에 'PO' 포함)
관리대장 06(PO번호·발행일)·21(PO정산연결)과 대조해 4유형을 검출한다.

  [유형A] 쿠팡PO 있는데 원장 미등록      : 받은 PO를 관리대장에 안 옮김 (누락)
  [유형B] 원장 PO번호가 쿠팡 목록에 없음  : 번호 오기입 가능성
  [유형C] 금액 불일치                    : 쿠팡 PO 금액 ≠ 원장 정산 공급가액 합
  [유형D] 미연결 제안                    : PO필요·번호없는 정산 ↔ 미등록 PO 금액 매칭 후보
          └ 금액이 양방향 유일 일치하면 자동입력 큐(06 PO번호·PO발행일, 빈 칸만)에 적재

파서는 머리글 자동탐지 + 전체 셀에서 PO번호 패턴(PO+숫자) 추출이라
쿠팡 내보내기 형식이 달라도 동작한다. 원장은 read-only. 결과는 reports/.

실행:  python po_reconcile.py                       # inbox에서 'PO' 파일 자동 탐지
       python po_reconcile.py --file 경로 [--master 경로]   # 테스트용
"""
import sys, os, re, csv, glob, json
from datetime import datetime
from collections import defaultdict, Counter

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from ecount_reconcile import read_ledger, load_config, _num, _d, supply_from_statement

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INBOX_DIR = os.path.join(BASE_DIR, "inbox")
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(BASE_DIR, "reports")
PO_PAT = re.compile(r"\bPO[\s-]?(\d{3,})\b", re.I)


def norm_po(s):
    m = PO_PAT.search(str(s or ""))
    return f"PO{m.group(1)}" if m else ""


def parse_po_export(path):
    """쿠팡 PO 목록 엑셀 → [{po, date, amount, desc}] (형식 유연 파싱)"""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    pos = {}
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        # 머리글 탐지(선택적): 금액·일자 열 위치 힌트
        j_amt = j_date = None
        for r in rows[:15]:
            names = {str(c).strip(): j for j, c in enumerate(r) if c is not None}
            for n, j in names.items():
                if j_amt is None and any(k in n for k in ("금액", "공급가", "합계", "Amount", "amount", "Total")):
                    j_amt = j
                if j_date is None and any(k in n for k in ("일자", "날짜", "발행", "Date", "date")):
                    j_date = j
            if j_amt is not None:
                break
        for r in rows:
            if r is None:
                continue
            blob = " ".join(str(c) for c in r if c is not None)
            po = norm_po(blob)
            if not po:
                continue
            amt = _num(r[j_amt]) if j_amt is not None and j_amt < len(r) else None
            if amt is None:  # 금액 열을 못 찾으면 행에서 가장 큰 숫자 추정
                nums = [_num(c) for c in r if _num(c) and _num(c) > 1000]
                amt = max(nums) if nums else None
            dt = ""
            if j_date is not None and j_date < len(r):
                dt = _d(r[j_date])[:10]
            if not dt:
                m = re.search(r"(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})", blob)
                dt = m.group(1).replace("/", "-").replace(".", "-") if m else ""
            cur = pos.get(po)
            if not cur or (amt and not cur.get("amount")):
                pos[po] = {"po": po, "date": dt, "amount": amt,
                           "desc": blob[:120].replace("\n", " ")}
    wb.close()
    return list(pos.values())


_ERP_SUPPLY = None


def erp_supply_index():
    """프로젝트NO → ERP 판매조회 공급가액. 한 번만 만든다(Z: 재귀 탐색이라 비싸다)."""
    global _ERP_SUPPLY
    if _ERP_SUPPLY is None:
        try:
            import erp_sales_index
            idx, _ = erp_sales_index.build()
            _ERP_SUPPLY = {k: (v.get("supply") or 0) for k, v in idx.items()}
        except Exception as e:
            print(f"  (ERP 판매조회 색인 생략: {e})")
            _ERP_SUPPLY = {}
    return _ERP_SUPPLY


_ERP_BY_PO = None


def erp_sum_by_po():
    """PO번호 → 그 PO 로 끊긴 ERP 전표 공급가액 합.

    ★ 차액이 '얼마나' 안 맞는지보다 **왜** 안 맞는지가 중요하다(2026-08-08).
      쿠팡금액 = ERP합 인데 원장합만 모자라면 그건 금액 오류가 아니라
      **원장에 그 PO 로 연결된 정산행이 빠진 것**이다 — 고칠 곳이 아주 다르다.
    """
    global _ERP_BY_PO
    if _ERP_BY_PO is None:
        out = defaultdict(int)
        try:
            import erp_sales_index
            idx, _ = erp_sales_index.build()
            for v in idx.values():
                for token in PO_PAT.finditer(str(v.get("po") or "")):
                    out["PO" + token.group(1)] += (v.get("supply") or 0)
        except Exception as e:
            print(f"  (ERP PO 합계 생략: {e})")
        _ERP_BY_PO = dict(out)
    return _ERP_BY_PO


def supply_of(r):
    """정산 한 행의 공급가액 — **앱 화면과 같은 사다리**를 쓴다.

    ★ 2026-08-08: 여기가 원장 06시트 공급가액 하나만 보던 탓에 유형C(금액 불일치)가
      51건 중 44건으로 나왔다. 근거는 전부 `원장공급가액합: 0` — 그 칸은 사람 손
      입력이라 비어 있을 뿐이었다. 정작 ERP 로 대면 원 단위까지 맞았다
      (PO327948 쿠팡 7,551,500 = ERP 합 7,551,500).
      앱 `app_server` 는 이미 실제작업 → ERP → 명세서 순으로 채워 보여 주고 있었으니
      **화면과 대조기가 서로 다른 금액을 보고 있었다.** 경보가 44/51 이면 아무도
      안 본다 — 조용한 사고의 반대편이지만 결과는 같다.
    """
    v = r.get("원장_공급가액")
    if v:
        return int(v)
    erp = erp_supply_index().get(str(r.get("프로젝트NO") or ""))
    if erp:
        return int(erp)
    conv = supply_from_statement(r.get("원장_거래명세서합계"))
    if conv is not None:
        return int(conv)
    return 0


def ledger_po_view(master):
    """원장 측 PO 시각: PO번호 → [정산행], PO없는 유상 정산 목록"""
    recs = read_ledger(master)
    by_po, no_po = defaultdict(list), []
    for sid, r in sorted(recs.items()):
        po = norm_po(r.get("원장_PO번호"))
        if po:
            by_po[po].append(r)
        elif r.get("비용구분") == "유상" and r.get("원장_공급가액") and \
                str(r.get("원장_PO필요여부", "")).startswith("필요"):
            no_po.append(r)
    return by_po, no_po


def main():
    args = sys.argv[1:]
    cfg = load_config()
    master = cfg["reconcile"]["master_xlsx"]
    if "--master" in args:
        master = args[args.index("--master") + 1]
    if "--file" in args:
        files = [args[args.index("--file") + 1]]
    else:
        from inbox_scan import pick
        files = pick("po")                     # 정본 폴더 전체를 내용으로 판별
    if not files:
        sys.exit("inbox/ 에 쿠팡 PO 목록이 없습니다. PO 목록 엑셀을 inbox/ 에 넣어주세요"
                 "(파일명은 아무거나 괜찮습니다).")

    coupang = []
    for f in files:
        part = parse_po_export(f)
        coupang += part
        print(f"'{os.path.basename(f)}': PO {len(part)}건 파싱")
    cp_by_no = {p["po"]: p for p in coupang}

    by_po, no_po = ledger_po_view(master)
    tol = cfg["reconcile"].get("금액허용오차", 0)

    # ERP가 이미 계산서를 끊었는지부터 본다.
    # 쿠팡 PO 대부분은 '정기점검 24건'처럼 여러 건을 묶은 것이라 06시트 PO번호 칸에
    # 안 적힌다. 그걸 전부 '원장 미등록'이라고 하면 94건이 쏟아져 무엇이 진짜 누락인지
    # 알 수 없다(2026-07-27). **계산서가 안 끊긴 PO만** 진짜 미청구다.
    erp_amts = set()
    erp_amt_counts = Counter()
    try:
        import erp_bundle as _EB
        for _d in _EB.load_erp()[1]:
            erp_amts.add(_d["amt"])
            erp_amts.add(round(_d["amt"] * 1.1))
            erp_amt_counts[_d["amt"]] += 1
            erp_amt_counts[round(_d["amt"] * 1.1)] += 1
    except Exception as e:
        print(f"  (ERP 계산서 대조 생략: {e})")

    def invoiced(amount):
        if not amount or not erp_amts:
            return False
        a = int(amount)
        return a in erp_amts or round(a / 1.1) in erp_amts

    po_amount_counts = Counter(int(p["amount"]) for p in cp_by_no.values() if p.get("amount"))

    def uniquely_invoiced(amount):
        """금액만 같은 다른 PO·계산서를 직원 완료 근거로 오인하지 않는다."""
        if not amount:
            return False
        a = int(amount)
        return (po_amount_counts[a] == 1
                and (erp_amt_counts[a] == 1 or erp_amt_counts[round(a / 1.1)] == 1))

    # 밴드에 사람이 적어 둔 PO별 처리 상태를 읽는다.
    # 오종현 매니저 확인(2026-07-27): "PO 작업 내역은 밴드 매출처업무에 1월부터 다 적어 둔다."
    # 이걸 안 보면 **이미 발행한 PO를 미청구로 보고**하게 된다(PO344599가 실제로 그랬다).
    try:
        import po_band_status as _PB
        band_po = _PB.scan()
    except Exception as e:
        print(f"  (밴드 PO 상태 읽기 생략: {e})")
        band_po = {}

    A, B, C, OK, D = [], [], [], [], []
    billed = 0
    billed_rows = {}
    billed_ambiguous = {}
    for po, p in sorted(cp_by_no.items()):
        lrows = by_po.get(po)
        if not lrows:
            b = band_po.get(po) or {}
            st = b.get("상태") or []
            amount_billed = invoiced(p["amount"])
            band_billed = "발행완료" in st
            if amount_billed or band_billed:
                billed += 1          # 문제 목록에서는 발행 정황으로 제외한다.
                if band_billed:
                    billed_rows[po] = {**p, "basis": "밴드 세금계산서 발행완료"}
                elif uniquely_invoiced(p["amount"]):
                    billed_rows[po] = {**p, "basis": "ERP 계산서 금액 유일 일치"}
                else:
                    billed_ambiguous[po] = {**p, "basis": "ERP 계산서 금액 비유일 일치"}
                continue
            note = ""
            if st:
                note = " · 밴드: " + ",".join(st)
            if b.get("프로젝트"):
                note += " · " + ",".join(b["프로젝트"][:2])
            A.append({"PO번호": po, "쿠팡금액": p["amount"], "발행일": p["date"], "내용": p["desc"][:60],
                      "프로젝트NO": ",".join(b.get("프로젝트", [])[:3]),
                      "밴드상태": ",".join(st),
                      "판정": "미청구 — 쿠팡 PO는 받았는데 계산서가 발행되지 않았습니다" + note})
            continue
        led_sum = sum(supply_of(r) for r in lrows)
        raw_sum = sum((r.get("원장_공급가액") or 0) for r in lrows)
        ids = ",".join(r["정산ID"] for r in lrows)
        if p["amount"] is not None and abs(led_sum - p["amount"]) > tol:
            erp_sum = erp_sum_by_po().get(po, 0)
            if erp_sum and abs(erp_sum - p["amount"]) <= tol:
                # 쿠팡 = ERP 인데 원장만 모자라다 → 금액이 틀린 게 아니라 연결이 빠졌다.
                verdict = ("원장 연결 누락 — 쿠팡·ERP 금액은 일치하는데 "
                           "이 PO 로 묶인 원장 정산행이 모자랍니다")
            elif erp_sum:
                verdict = "쿠팡·ERP·원장 세 값이 모두 다릅니다 — 금액 확인 필요"
            else:
                verdict = "ERP 전표 없음 — 원장 금액과만 비교했습니다"
            C.append({"PO번호": po, "정산ID": ids, "쿠팡금액": p["amount"], "원장공급가액합": led_sum,
                      "원장직접입력합": raw_sum, "ERP전표합": erp_sum,
                      "금액출처": "원장" if raw_sum == led_sum else "원장+ERP·명세서 보완",
                      "차액": round((p["amount"] or 0) - led_sum), "판정": verdict})
        else:
            OK.append(po)
    for po, lrows in sorted(by_po.items()):
        if po not in cp_by_no:
            B.append({"PO번호": po, "정산ID": ",".join(r["정산ID"] for r in lrows),
                      "원장공급가액합": sum(supply_of(r) for r in lrows),
                      "판정": "쿠팡 목록에 없음 — 번호 오기입 또는 목록 기간 밖"})

    # 유형D: 미등록 PO ↔ PO없는 유상 정산 — 금액 유일 매칭이면 자동입력 후보
    # ★ 여기만 supply_of() 를 쓰지 않는다(2026-08-08). A~C 는 **경보**라 금액 출처를
    #   넓혀도 잃는 것이 없지만, D 는 06시트에 PO번호를 **써 넣는** 길이다.
    #   짐작으로 채운 금액으로 짝을 지으면 틀린 PO번호가 원장에 박히고, 그건 빈 칸보다
    #   나쁘다. 자동으로 쓰는 자리는 **사람이 직접 적은 금액**만 근거로 삼는다.
    unmatched_po = [p for p in coupang if p["po"] not in by_po and p["amount"]]
    queue_items = []
    for p in unmatched_po:
        cands = [r for r in no_po if abs((r.get("원장_공급가액") or 0) - p["amount"]) <= tol]
        if cands:
            unique = (len(cands) == 1 and
                      sum(1 for q in unmatched_po if q["amount"] and
                          abs(q["amount"] - (cands[0].get("원장_공급가액") or 0)) <= tol) == 1)
            D.append({"PO번호": p["po"], "쿠팡금액": p["amount"], "발행일": p["date"],
                      "후보정산": ",".join(r["정산ID"] for r in cands[:5]),
                      "판정": "유일 매칭 → 자동입력" if unique else f"후보 {len(cands)}건 — 수동 확인"})
            if unique:
                sid = cands[0]["정산ID"]
                queue_items.append({"sheet": "06_거래서류청구수금", "key_col": "정산ID", "key": sid,
                                    "col": "PO번호", "value": p["po"], "vtype": "text",
                                    "evidence": f"쿠팡 PO목록 금액 유일매칭({p['amount']:,})", "only_if_empty": True})
                if p["date"]:
                    queue_items.append({"sheet": "06_거래서류청구수금", "key_col": "정산ID", "key": sid,
                                        "col": "PO발행일", "value": p["date"], "vtype": "date",
                                        "evidence": "쿠팡 PO목록 발행일", "only_if_empty": True})
    if queue_items and "--no-queue" not in args:
        try:
            from ledger_writer import queue_add
            print("자동입력 큐 적재:", queue_add(queue_items), "건 (PO번호·발행일)")
        except Exception as e:
            print("(큐 적재 실패:", e, ")")

    os.makedirs(REPORT_DIR, exist_ok=True)
    base = os.path.join(REPORT_DIR, f"PO대조_{datetime.now():%Y%m%d_%H%M}")
    with open(base + ".md", "w", encoding="utf-8") as f:
        f.write("# 쿠팡 PO ↔ 관리대장 대조\n\n")
        f.write(f"- 생성 {datetime.now():%Y-%m-%d %H:%M} / 쿠팡 PO {len(coupang)}건 · 원장 PO {len(by_po)}건 · 정상 일치 {len(OK)}건\n")
        f.write(f"- PO필요·번호없는 유상 정산: {len(no_po)}건\n")
        for title, rows_ in [("A. 미청구 PO — 계산서 미발행 (★)", A), ("B. 쿠팡 목록에 없는 원장 PO (오기입 의심)", B),
                             ("C. 금액 불일치", C), ("D. 연결 제안 (미등록 PO ↔ 무PO 정산)", D)]:
            f.write(f"\n## {title} — {len(rows_)}건\n\n")
            if rows_:
                cols = list(rows_[0].keys())
                f.write("| " + " | ".join(cols) + " |\n|" + "---|" * len(cols) + "\n")
                for r in rows_:
                    f.write("| " + " | ".join(str(r[c]) for c in cols) + " |\n")
    allrows = ([{"유형": "A", **r} for r in A] + [{"유형": "B", **r} for r in B]
               + [{"유형": "C", **r} for r in C] + [{"유형": "D", **r} for r in D])
    if allrows:
        keys = []
        for r in allrows:
            for k in r:
                if k not in keys:
                    keys.append(k)
        with open(base + ".csv", "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=keys); w.writeheader(); w.writerows(allrows)

    # 객관 자료로 끝난 담당자 업무를 별도 근거 파일에 남긴다. 문제 CSV는 미해결만
    # 남기므로 완료 이력의 정본으로 사용할 수 없다. 다음 staff_completion 단계가
    # 이 파일을 읽어 "류지영/오종현/유현민 완료"를 SQLite에 멱등 반영한다.
    recognized = datetime.now().date().isoformat()
    source_names = ", ".join(sorted({os.path.basename(path) for path in files}))
    staff_entries = []
    staff_retractions = []
    for po, item in sorted(cp_by_no.items()):
        staff_entries.append({
            "owner": "오종현", "task_kind": "po_source", "record_id": po,
            "project": item.get("project") or "",
            "completed_on": item.get("date") or recognized,
            "basis": f"쿠팡 PO 원본 확인 · {source_names}",
        })
    for po in sorted(OK):
        item = cp_by_no[po]
        for owner, task_kind in (("유현민", "po_system_verified"),
                                 ("류지영", "po_amount_verified")):
            staff_entries.append({
                "owner": owner, "task_kind": task_kind, "record_id": po,
                "project": item.get("project") or "", "completed_on": recognized,
                "basis": "쿠팡 PO 금액과 원장 공급가액 합계 일치",
            })
    for po, item in sorted(billed_rows.items()):
        for owner, task_kind in (("유현민", "po_system_verified"),
                                 ("류지영", "billing_verified")):
            staff_entries.append({
                "owner": owner, "task_kind": task_kind, "record_id": po,
                "project": item.get("project") or "", "completed_on": recognized,
                "basis": item["basis"],
            })
    for po in sorted(billed_ambiguous):
        for owner, task_kind in (("유현민", "po_system_verified"),
                                 ("류지영", "billing_verified")):
            staff_retractions.append({
                "owner": owner, "task_kind": task_kind, "record_id": po,
                "reason": "금액만 같고 PO·계산서 유일 연결이 아님",
            })
    try:
        from ledger_writer import atomic_json_dump
        atomic_json_dump({
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "sources": sorted({os.path.basename(path) for path in files}),
            "entries": staff_entries,
            "retractions": staff_retractions,
            "counts": {"오종현": len(cp_by_no), "유현민": len(OK) + len(billed_rows),
                       "류지영": len(OK) + len(billed_rows)},
        }, os.path.join(REPORT_DIR, "po_objective_evidence.json"))
    except Exception as e:
        print("(담당자 객관완료 근거 저장 실패:", e, ")")
    print(f"A(미청구) {len(A)} / B(쿠팡목록에없음) {len(B)} / C(금액불일치) {len(C)} / "
          f"D(연결제안) {len(D)} / 정상 {len(OK)} / 계산서발행됨 {billed}")
    print("리포트:", base + ".md")


if __name__ == "__main__":
    main()
