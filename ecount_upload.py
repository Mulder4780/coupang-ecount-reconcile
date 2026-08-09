# -*- coding: utf-8 -*-
"""
ecount_upload.py — 관리대장 → 이카운트 매출전표 전송 **대기 목록 미리보기**(전송 안 함)
====================================================================================
관리대장 06_거래서류청구수금의 유상 정산 건 중 ERP에 등록해야 할 것을 골라 리포트로 낸다.

★ 실전송 금지 (2026-08-05 사용자 지시)
  - **이 도구는 ERP에 전표를 등록하지 않는다.** `--post` 를 줘도 미리보기까지만 하고 멈춘다.
    되돌릴 수 없는 등록을 앱·AI가 대신하지 않는다 — 사람이 ERP 화면에서 직접 한다.
  - `--post` 인자와 아래 전송 코드는 **옛 스케줄러·바로가기가 조용히 전송하지 못하도록**
    남겨 둔 차단 지점이다. 되살리려면 사용자 지시가 다시 있어야 한다.
  - (참고) 트래픽 원칙: 이카운트는 저장 1회/10초·연속 오류 30건 제한이 있다.

실행:
    python ecount_upload.py               # 전송 대기 목록 미리보기 (API 호출 0회)
"""
import sys, os, json, time, csv
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from ecount_client import EcountClient, load_config, EcountError
from ecount_reconcile import read_ledger  # 원장 read-only 로더 재사용

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.path.join(BASE_DIR, "reports")
STATE_PATH = os.path.join(BASE_DIR, ".uploaded_slips.json")


def load_state():
    try:
        return json.load(open(STATE_PATH, "r", encoding="utf-8"))
    except Exception:
        return {}


def save_state(st):
    json.dump(st, open(STATE_PATH, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def yyyymmdd(s):
    s = (s or "").strip()
    return s.replace("-", "")[:8] if s else ""


def build_payload(rec, up):
    """정산 1건 → SaveInvoiceAuto BulkDatas (사용자 제공 스펙 기준)."""
    supply = int(round(rec["원장_공급가액"] or 0))
    vat = int(round(supply * 0.1))
    return {
        "InvoiceAutoList": [{
            "BulkDatas": {
                "TRX_DATE": yyyymmdd(rec.get("작업완료일")),
                "ACCT_DOC_NO": "",
                "TAX_GUBUN": up.get("TAX_GUBUN", "11"),
                "S_NO": "",
                "CUST": up.get("CUST", ""),
                "CUST_DES": "",
                "SUPPLY_AMT": str(supply),
                "VAT_AMT": str(vat),
                "ACCT_NO": "",
                "CR_CODE": up.get("CR_CODE", ""),
                "DR_CODE": "",
                "REMARKS_CD": "",
                "REMARKS": f"{rec['정산ID']} {rec.get('캠프명','')} {rec.get('업무구분','')} (관리대장 자동등록)",
                "SITE_CD": "",
                "PJT_CD": rec.get("프로젝트NO", ""),
                "ITEM1_CD": "", "ITEM2_CD": "", "ITEM3_CD": "",
                "ITEM4": "", "ITEM5": "", "ITEM6": "", "ITEM7": "", "ITEM8": "",
            }
        }]
    }


def main():
    args = sys.argv[1:]
    do_post = "--post" in args
    limit = None
    if "--limit" in args:
        limit = int(args[args.index("--limit") + 1])

    cfg = load_config()
    up = cfg.get("upload", {})
    interval = float(up.get("전송간격_초", 10.5))

    # 1) 대상 선정: 유상 & 공급가액>0 & 작업완료일 존재 & 미업로드
    recs = read_ledger(cfg["reconcile"]["master_xlsx"])
    state = load_state()
    targets, skipped = [], {"유상 아님(무상·보험·미입력)": 0, "금액없음": 0, "완료일없음": 0, "업로드완료": 0,
                            "ERP기등록추정(명세서번호있음)": 0}
    for sid, r in sorted(recs.items()):
        if r.get("비용구분") != "유상":
            skipped["유상 아님(무상·보험·미입력)"] += 1; continue
        if not r.get("원장_공급가액"):
            skipped["금액없음"] += 1; continue
        if not yyyymmdd(r.get("작업완료일")):
            skipped["완료일없음"] += 1; continue
        if sid in state:
            skipped["업로드완료"] += 1; continue
        # ★ 이중계상 방지: 06시트 거래명세서번호(=ERP 일자-No. 형식)가 이미 있으면
        #   해당 건은 ERP에 전표가 존재할 가능성이 높아 전송에서 제외한다.
        #   (실존 여부는 거래처별계정별원장 7월 내보내기를 inbox 대조로 확정 가능)
        if str(r.get("원장_거래명세서번호") or "").strip() and "--include-billed" not in sys.argv:
            skipped["ERP기등록추정(명세서번호있음)"] += 1; continue
        targets.append(r)
    if limit:
        targets = targets[:limit]

    # 2) 필수 회계코드 검증
    missing = [k for k in ("CUST", "CR_CODE") if not (up.get(k) or "").strip()]

    mode = "실전송(POST)" if do_post else "미리보기(dry-run)"
    print(f"모드: {mode} / 대상 {len(targets)}건 / 제외 {skipped}")
    if missing:
        print(f"★ 미설정 회계코드: {', '.join(missing)} — config의 upload 절에 입력 필요 (실전송 불가)")

    # 3) 미리보기 리포트 (dry-run이든 post든 항상 생성)
    os.makedirs(REPORT_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    base = os.path.join(REPORT_DIR, f"전표업로드_{'실전송' if do_post else '미리보기'}_{stamp}")
    with open(base + ".csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["정산ID", "TRX_DATE", "PJT_CD(프로젝트NO)", "SUPPLY_AMT", "VAT_AMT", "REMARKS",
                    "TAX_GUBUN", "CUST", "CR_CODE", "전송결과"])
        for r in targets:
            p = build_payload(r, up)["InvoiceAutoList"][0]["BulkDatas"]
            w.writerow([r["정산ID"], p["TRX_DATE"], p["PJT_CD"], p["SUPPLY_AMT"], p["VAT_AMT"],
                        p["REMARKS"], p["TAX_GUBUN"], p["CUST"] or "(미설정)", p["CR_CODE"] or "(미설정)",
                        "" if do_post else "dry-run"])
    total = sum(int(round(r["원장_공급가액"])) for r in targets)
    print(f"공급가액 합계: {total:,}원 / 리포트: {base}.csv")

    if not do_post:
        print("\n→ 미리보기까지만 제공합니다. ERP 실제 등록은 사람이 ERP 화면에서 직접 하세요.")
        return
    # ★ 실전송 금지(2026-08-05 사용자 지시) — 되돌릴 수 없는 ERP 등록은 앱·AI가 하지 않는다.
    #   --post 를 남겨 둔 것은 옛 스케줄러·바로가기가 조용히 전송하지 않고 여기서 멈추게 하기 위함이다.
    sys.exit("실전송은 제공하지 않습니다(2026-08-05 지시). 위 미리보기 리포트를 보고 ERP에서 직접 등록하세요.")
    if missing:
        sys.exit("실전송 중단: 필수 회계코드 미설정")
    if not targets:
        print("전송할 건이 없습니다."); return

    # 4) 실전송 — 트래픽 보호: 간격 준수·연속오류 3회 중단
    cli = EcountClient(cfg)
    cli.login()
    dom = cfg["endpoints"]["invoice_auto"]
    consec_err = 0
    ok = 0
    for i, r in enumerate(targets):
        payload = build_payload(r, up)
        url = f"{cli._root()}/OAPI/V2/{dom['domain']}/{dom['method']}?SESSION_ID={cli.session_id}"
        try:
            resp = cli._post(url, payload)
            data = resp.get("Data") or {}
            # 공식 Result 스펙: SuccessCnt/FailCnt/SlipNos(전표번호, 실패시 공백)/TRACE_ID/QUANTITY_INFO/EXPIRE_DATE
            success = str(data.get("SuccessCnt", "")).strip() == "1"
            raw_slip = data.get("SlipNos", "")
            if isinstance(raw_slip, list):
                raw_slip = raw_slip[0] if raw_slip else ""
            slip_no = str(raw_slip or "").strip()
            m = __import__("re").match(r"^(\d{4})(\d{2})(\d{2})-(\d+)$", slip_no)
            if m:                                   # 20260720-2 → 2026/07/20-2 (원장 표기 규격)
                slip_no = f"{m.group(1)}/{m.group(2)}/{m.group(3)}-{m.group(4)}"
            if data.get("EXPIRE_DATE"):
                print(f"  ⚠ API 버전 서비스 종료 예정일: {data['EXPIRE_DATE']}")
            if success:
                ok += 1; consec_err = 0
                state[r["정산ID"]] = {"ts": datetime.now().isoformat(), "supply": r["원장_공급가액"],
                                      "prj": r.get("프로젝트NO"), "slip_no": slip_no,
                                      "test": bool(cfg["auth"].get("IS_TEST"))}
                save_state(state)
                print(f"[{i+1}/{len(targets)}] {r['정산ID']} 등록 OK → ERP 전표번호 {slip_no or '(미회신)'}")
                # 확정 사실 → 관리대장 자동입력 큐 (ledger_writer가 빈 칸만 채움)
                try:
                    from ledger_writer import queue_add
                    t = build_payload(r, up)["InvoiceAutoList"][0]["BulkDatas"]["TRX_DATE"]
                    ups = []
                    if slip_no:
                        ups.append({"sheet": "06_거래서류청구수금", "key_col": "정산ID", "key": r["정산ID"],
                                    "col": "거래명세서번호", "value": slip_no, "vtype": "text",
                                    "evidence": "ECOUNT SaveInvoiceAuto 응답 SlipNos", "only_if_empty": True})
                        ups.append({"sheet": "06_거래서류청구수금", "key_col": "정산ID", "key": r["정산ID"],
                                    "col": "거래명세서발행일", "value": f"{t[:4]}-{t[4:6]}-{t[6:8]}", "vtype": "date",
                                    "evidence": "ECOUNT 전표 TRX_DATE", "only_if_empty": True})
                    if ups:
                        queue_add(ups)
                except Exception as e:
                    print("  (자동입력 큐 적재 실패:", e, ")")
            else:
                consec_err += 1
                print(f"[{i+1}/{len(targets)}] {r['정산ID']} 실패 — 원문 응답:")
                print(json.dumps(resp, ensure_ascii=False)[:1200])
        except EcountError as e:
            consec_err += 1
            print(f"[{i+1}/{len(targets)}] {r['정산ID']} 오류: {str(e)[:150]}")
        if consec_err >= 3:
            print("★ 연속 오류 3회 — 트래픽 보호를 위해 중단합니다. 원인 확인 후 재실행하세요(성공분은 중복 전송되지 않음).")
            break
        if i < len(targets) - 1:
            time.sleep(interval)
    try:
        if data.get("QUANTITY_INFO"):
            print("허용량 현황:", data["QUANTITY_INFO"])
    except NameError:
        pass
    print(f"\n완료: 성공 {ok}/{len(targets)}건")
    print("ERP 전표번호(SlipNos)는 .uploaded_slips.json 에 정산ID별로 기록됨 — 원장↔ERP 영구 연결키.")


if __name__ == "__main__":
    main()
