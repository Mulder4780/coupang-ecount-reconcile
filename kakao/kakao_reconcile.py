# -*- coding: utf-8 -*-
"""
kakao_reconcile.py — 카카오톡 대화 내보내기(.txt) ↔ 관리대장 자동 대조
========================================================================
카카오는 채팅방 읽기 API를 제공하지 않으므로(보내기 API만 존재),
공식 기능인 [대화 내보내기]로 저장한 .txt를 kakao/inbox/ 에 넣으면 자동 처리한다.
(PC 카톡: 채팅방 → ☰ → 대화 내용 → 내보내기 / 모바일: 채팅방 설정 → 대화 내보내기)

지원 형식(자동 감지):
  PC:    --------------- 2026년 7월 20일 월요일 ---------------
         [유현민] [오후 2:59] 메시지 내용
  모바일: 2026년 7월 20일 오후 2:59, 유현민 : 메시지 내용

대조: 관리대장 02(돌발AS)·04(정기점검) 완료 건 ↔ 카톡 메시지
  1순위 프로젝트NO 포함 / 2순위 캠프명 핵심부 + 날짜 ±3일 / 보조 발신자=담당기사
원장은 read-only. 결과는 reports/ 에만 출력.

실행:  python kakao/kakao_reconcile.py            # kakao/inbox/*.txt 전체
       python kakao/kakao_reconcile.py --file 경로 [--master 경로]   # 테스트용
"""
import sys, os, re, csv, json, glob
from datetime import datetime, date

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE_DIR)
CONFIG_PATH = os.path.join(ROOT, "config", "ecount_config.json")
INBOX_DIR = os.path.join(BASE_DIR, "inbox")
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")
HDR_ROW, FIRST = 4, 5

cfg = json.load(open(CONFIG_PATH, encoding="utf-8"))
DATE_TOL = int(cfg.get("kakao", {}).get("date_tolerance_days", 3))

RE_PC_DATE = re.compile(r"-+\s*(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일.*-+")
RE_PC_MSG = re.compile(r"^\[([^\]]+)\]\s*\[(오전|오후)\s*(\d{1,2}):(\d{2})\]\s*(.*)$")
RE_MO_MSG = re.compile(r"^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일\s*(오전|오후)\s*(\d{1,2}):(\d{2}),\s*([^:]+?)\s*:\s*(.*)$")
RE_MO_DATE = re.compile(r"^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일\s*\S요일\s*$")


def parse_export(path):
    """카톡 내보내기 txt → [{date, time, sender, text}] (여러 줄 메시지는 직전 메시지에 병합)"""
    msgs, cur_date = [], None
    # 내보내기 인코딩은 보통 UTF-8, 구버전은 cp949
    for enc in ("utf-8-sig", "utf-8", "cp949"):
        try:
            lines = open(path, encoding=enc).read().splitlines()
            break
        except UnicodeDecodeError:
            continue
    else:
        raise UnicodeDecodeError("kakao", b"", 0, 0, "지원 인코딩 아님")
    for ln in lines:
        m = RE_PC_DATE.match(ln) or RE_MO_DATE.match(ln)
        if m:
            g = m.groups()
            cur_date = date(int(g[0]), int(g[1]), int(g[2]))
            continue
        m = RE_PC_MSG.match(ln)
        if m and cur_date:
            sender, ap, hh, mm, text = m.groups()
            h = int(hh) % 12 + (12 if ap == "오후" else 0)
            msgs.append({"date": cur_date, "time": f"{h:02d}:{mm}", "sender": sender.strip(), "text": text})
            continue
        m = RE_MO_MSG.match(ln)
        if m:
            y, mo, d, ap, hh, mm, sender, text = m.groups()
            cur_date = date(int(y), int(mo), int(d))
            h = int(hh) % 12 + (12 if ap == "오후" else 0)
            msgs.append({"date": cur_date, "time": f"{h:02d}:{mm}", "sender": sender.strip(), "text": text})
            continue
        if msgs and ln.strip() and not ln.startswith("저장한 날짜"):
            msgs[-1]["text"] += " " + ln.strip()      # 여러 줄 메시지 병합
    return msgs


def camp_core(camp):
    return re.split(r"[(\s]", str(camp or ""))[0].strip()


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def read_rows(master):
    import openpyxl
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    spec = {
        "02_돌발AS접수": ("접수ID", "프로젝트NO", "캠프명", "담당기사", "작업완료일", "진행상태"),
        "04_정기점검":   ("점검ID", "프로젝트NO", "캠프명", "담당기사", "실제점검일", "점검상태"),
    }
    out = []
    for sheet, cols in spec.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        hdr = next(ws.iter_rows(min_row=HDR_ROW, max_row=HDR_ROW, values_only=True))
        idx = {str(h).strip(): i for i, h in enumerate(hdr) if h is not None}
        def gv(row, name):
            return row[idx[name]] if name in idx and idx[name] < len(row) else None
        for row in ws.iter_rows(min_row=FIRST, values_only=True):
            # ID 열은 수식이라 새로 추가된 행은 엑셀을 열기 전까지 캐시값이 없다(None).
            # 그런 행도 대조 대상이어야 하므로 프로젝트NO를 대체 키로 쓴다.
            rid = gv(row, cols[0]) or gv(row, cols[1])
            done = to_date(gv(row, cols[4]))
            if not rid or not done:
                continue
            out.append({"시트": sheet, "ID": str(rid), "프로젝트NO": str(gv(row, cols[1]) or ""),
                        "캠프명": str(gv(row, cols[2]) or ""), "담당기사": str(gv(row, cols[3]) or ""),
                        "완료일": done})
    wb.close()
    return out


def main():
    args = sys.argv[1:]
    # config에는 처음 만든 파일명(v20)이 그대로 적혀 있다. 실제로 봐야 할 건 **최신본**이다.
    # 다른 도구들은 전부 resolve_master를 쓰는데 여기만 빠져 있어 파일 없음으로 죽었다.
    sys.path.insert(0, ROOT)
    from ecount_reconcile import resolve_master
    master = resolve_master(cfg["reconcile"]["master_xlsx"])
    if "--master" in args:
        master = args[args.index("--master") + 1]
    if "--file" in args:
        files = [args[args.index("--file") + 1]]
    else:
        files = [f for f in glob.glob(os.path.join(INBOX_DIR, "*.txt"))]
    if not files:
        sys.exit("kakao/inbox/ 에 카톡 내보내기 .txt가 없습니다. (PC 카톡: 채팅방 → ☰ → 대화 내용 → 내보내기)")

    msgs = []
    for f in files:
        part = parse_export(f)
        room = os.path.splitext(os.path.basename(f))[0]
        for m in part:
            m["room"] = room
        msgs += part
        print(f"'{os.path.basename(f)}': 메시지 {len(part)}건 파싱")

    rows = read_rows(master)
    print(f"카톡 메시지 {len(msgs)}건 / 원장 완료건 {len(rows)}건 대조")

    matched_keys = set()
    results = []
    for r in rows:
        best, how = None, ""
        core = camp_core(r["캠프명"])
        for m in msgs:
            near = abs((m["date"] - r["완료일"]).days) <= DATE_TOL
            if r["프로젝트NO"] and r["프로젝트NO"] in m["text"]:
                best, how = m, "프로젝트NO"
                break
            if near and core and len(core) >= 3 and core in m["text"] and best is None:
                best, how = m, "캠프명+날짜"
        sender_ok = bool(best and r["담당기사"] and r["담당기사"] in best["sender"])
        if best:
            matched_keys.add(id(best))
        results.append({**{k: r[k] for k in ("시트", "ID", "프로젝트NO", "캠프명", "담당기사")},
                        "완료일": r["완료일"].isoformat(),
                        "카톡보고": "확인" if best else "미확인",
                        "매칭근거": how + ("+기사일치" if sender_ok else ""),
                        "메시지일": best["date"].isoformat() if best else "",
                        "발신자": best["sender"] if best else "",
                        "방": best["room"] if best else ""})

    miss = [r for r in results if r["카톡보고"] == "미확인"]
    os.makedirs(REPORT_DIR, exist_ok=True)
    base = os.path.join(REPORT_DIR, f"카톡대조_{datetime.now():%Y%m%d_%H%M}")
    cols = list(results[0].keys()) if results else []
    with open(base + ".csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols); w.writeheader(); w.writerows(results)
    with open(base + ".md", "w", encoding="utf-8") as f:
        f.write("# 카카오톡 ↔ 관리대장 대조 리포트\n\n")
        f.write(f"- 생성 {datetime.now():%Y-%m-%d %H:%M} / 완료건 {len(results)} / 카톡보고 확인 {len(results)-len(miss)} / 미확인 {len(miss)}\n\n")
        f.write("## 카톡 보고 미확인 건\n\n| ID | 캠프 | 기사 | 완료일 |\n|---|---|---|---|\n")
        for r in miss:
            f.write(f"| {r['ID']} | {r['캠프명']} | {r['담당기사']} | {r['완료일']} |\n")
    print(f"확인 {len(results)-len(miss)} / 미확인 {len(miss)}")
    print("리포트:", base + ".md")
    return results


if __name__ == "__main__":
    main()
