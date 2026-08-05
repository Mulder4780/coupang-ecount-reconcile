# -*- coding: utf-8 -*-
"""
customer_index.py — ERP 거래처등록 → **거래처코드 색인**, 캠프명과 잇는다

사용자 지시(2026-08-05): "거래처 코드 전수 조사 및 앱과 엑셀에 추가해서 표기해."
앱 화면에는 캠프명(강서1MB(가양A))만 있고 ERP 거래처코드(CU177)가 없어, ERP·씨앗과
대조할 때마다 사람이 눈으로 찾아야 했다.

원천: ERP `거래처등록` 엑셀(ESA001M — 거래처코드·거래처명·주소·담당자·보유장비명).
이름이 무작위라 **머리글로 판별**한다(거래처코드+거래처명이 2행에 있는 파일).

캠프명 ↔ 거래처명 잇기
  · 1순위 완전 일치         강서1MB(가양A) = 강서1MB(가양A)
  · 2순위 괄호/공백 제거 후 일치   강서1MB(가양A) → 강서1MB가양A
  · 3순위 캠프 핵심코드 일치       강서1MB / M_강릉1 / 송파5캠프 같은 앞부분
  맞는 게 **하나일 때만** 확정한다. 여럿이면 후보로 남기고 추측하지 않는다.

산출: reports/거래처코드_색인.json  {캠프명 → {code, erp_name, addr, manager, equip}}
      reports/거래처코드_색인.csv   (사람이 엑셀에서 보는 용도)

  python customer_index.py
"""
import csv
import glob
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

OUT_JSON = os.path.join(ROOT, "reports", "거래처코드_색인.json")
OUT_CSV = os.path.join(ROOT, "reports", "거래처코드_색인.csv")


def norm(s):
    """비교용 정규화 — 괄호·공백·기호를 지우고 대문자로."""
    return re.sub(r"[\s()\[\]_·\-/]", "", str(s or "")).upper()


def core(s):
    """캠프 핵심코드: '강서1MB(가양A)' → '강서1MB', 'M_강릉1' → 'M강릉1'."""
    s = str(s or "").split("(")[0]
    return norm(s)


def load_customers():
    import warnings
    warnings.filterwarnings("ignore")
    import openpyxl
    import source_dirs as S
    cands = [p for p in glob.glob(os.path.join(S.ERP_DIR, "**", "*.xlsx"), recursive=True)
             if not os.path.basename(p).startswith(("~$", "ESD007E"))]
    cands.sort(key=os.path.getmtime, reverse=True)
    best = None
    for p in cands[:80]:
        try:
            wb = openpyxl.load_workbook(p, read_only=False, data_only=True)
            ws = wb.active
            head = [str(c or "") for c in next(ws.iter_rows(min_row=2, max_row=2, values_only=True))]
            j = "|".join(head)
            # ★ 2026-08-05: '거래처코드'만 보면 거래처관리대장(104행) 같은 다른 화면이
            #   먼저 걸린다. 거래처**등록**은 주소·담당자 열이 함께 있고 행이 압도적으로 많다.
            if "거래처코드" not in j or "거래처명" not in j or "주소" not in j:
                wb.close()
                continue
            if best and ws.max_row <= best[1]:
                wb.close()
                continue
            idx = {h: i for i, h in enumerate(head)}
            rows = []
            for r in ws.iter_rows(min_row=3, values_only=True):
                r = ["" if c is None else str(c).strip() for c in r]
                g = lambda k: (r[idx[k]] if k in idx and len(r) > idx[k] else "")
                code, name = g("거래처코드"), g("거래처명")
                if not code or not name:
                    continue
                rows.append({"code": code, "name": name, "addr": g("주소"),
                             "manager": g("담당자"), "email": g("Email"),
                             "tel": g("연락처"), "equip": g("보유장비명")})
            wb.close()
            if not best or len(rows) > len(best[2]):
                best = (os.path.basename(p), ws.max_row, rows)
        except Exception:
            continue
    return (best[2], best[0]) if best else ([], None)


def camps_from_ledger():
    """관리대장에서 실제 쓰이는 캠프명을 모은다(프로젝트NO 와 함께)."""
    from ecount_reconcile import read_ledger, load_config
    recs = read_ledger(load_config()["reconcile"]["master_xlsx"])
    out = {}
    for r in recs.values():
        camp = str(r.get("캠프명") or "").strip()
        if not camp:
            continue
        d = out.setdefault(camp, {"건수": 0, "프로젝트": set()})
        d["건수"] += 1
        if r.get("프로젝트NO"):
            d["프로젝트"].add(str(r["프로젝트NO"]))
    return out


def main():
    custs, src = load_customers()
    if not custs:
        print("거래처등록 원본을 찾지 못함 — ERP '거래처등록' 화면 엑셀이 필요하다")
        return 1
    by_exact, by_norm, by_core = {}, {}, {}
    for c in custs:
        by_exact.setdefault(c["name"], []).append(c)
        by_norm.setdefault(norm(c["name"]), []).append(c)
        by_core.setdefault(core(c["name"]), []).append(c)

    camps = camps_from_ledger()
    linked, multi, none = {}, {}, []
    for camp, info in camps.items():
        for table, key, how in ((by_exact, camp, "완전일치"),
                                (by_norm, norm(camp), "정규화"),
                                (by_core, core(camp), "핵심코드")):
            hit = table.get(key) or []
            if len(hit) == 1:
                c = hit[0]
                linked[camp] = {"code": c["code"], "erp_name": c["name"], "how": how,
                                "addr": c["addr"], "manager": c["manager"],
                                "tel": c["tel"], "email": c["email"], "equip": c["equip"],
                                "건수": info["건수"]}
                break
            if len(hit) > 1:
                multi[camp] = {"how": how, "codes": [x["code"] for x in hit][:6],
                               "names": [x["name"] for x in hit][:6], "건수": info["건수"]}
                break
        else:
            none.append({"camp": camp, "건수": info["건수"]})

    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    json.dump({"src": src, "customers": len(custs), "camps": len(camps),
               "linked": linked, "ambiguous": multi, "unmatched": none},
              open(OUT_JSON, "w", encoding="utf-8"), ensure_ascii=False)
    with open(OUT_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["캠프명", "거래처코드", "ERP거래처명", "연결방식", "원장건수",
                    "주소", "담당자", "연락처", "Email", "보유장비"])
        for camp, v in sorted(linked.items()):
            w.writerow([camp, v["code"], v["erp_name"], v["how"], v["건수"],
                        v["addr"], v["manager"], v["tel"], v["email"], v["equip"]])
        for camp, v in sorted(multi.items()):
            w.writerow([camp, "(후보 여럿)", " / ".join(v["names"]), v["how"], v["건수"],
                        "", "", "", "", ""])
        for x in sorted(none, key=lambda r: -r["건수"]):
            w.writerow([x["camp"], "(ERP에 없음)", "", "", x["건수"], "", "", "", "", ""])
    # ★ 사용자 지시(2026-08-05) "앱과 엑셀에 추가해서 표기": 관리대장 본체는 수식·차트가
    #   있어 열을 늘리지 않는다(절대규칙). 대신 **별도 엑셀**을 같은 폴더에 만들어
    #   사람이 바로 열어 보게 한다 — 확인필요현황 엑셀과 같은 방식이다.
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "거래처코드"
        head = ["캠프명", "거래처코드", "ERP거래처명", "연결방식", "원장건수",
                "주소", "담당자", "연락처", "Email", "보유장비"]
        ws.append(head)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="2F5597")
            c.alignment = Alignment(horizontal="center", vertical="center")
        for camp, v in sorted(linked.items()):
            ws.append([camp, v["code"], v["erp_name"], v["how"], v["건수"],
                       v["addr"], v["manager"], v["tel"], v["email"], v["equip"]])
        for camp, v in sorted(multi.items()):
            ws.append([camp, "(후보 여럿)", " / ".join(v["names"]), v["how"], v["건수"]])
        for x in sorted(none, key=lambda r: -r["건수"]):
            ws.append([x["camp"], "(ERP에 없음)", "", "", x["건수"]])
        for col, w in zip("ABCDEFGHIJ", (26, 12, 26, 10, 9, 34, 12, 15, 24, 30)):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        from ecount_reconcile import load_config as _lc, resolve_master as _rm
        out_x = os.path.join(os.path.dirname(_rm(_lc()["reconcile"]["master_xlsx"])),
                             "쿠팡_거래처코드_최신.xlsx")
        wb.save(out_x)
        print(f"  엑셀: {out_x}")
    except Exception as e:
        print(f"  ! 엑셀 생성 실패: {str(e)[:70]}")

    print(f"거래처 {len(custs)}개 · 원장 캠프 {len(camps)}개 → "
          f"코드 확정 {len(linked)} · 후보여럿 {len(multi)} · ERP에 없음 {len(none)} "
          f"→ reports/거래처코드_색인.json/csv")
    return 0


if __name__ == "__main__":
    sys.exit(main())
