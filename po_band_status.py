# -*- coding: utf-8 -*-
"""
po_band_status.py — 밴드에 적힌 **PO별 상태·프로젝트NO**를 읽어 온다
================================================================================
오종현 매니저 확인(2026-07-27 통화): "PO 작업 내역은 밴드 매출처업무에 다 적혀 있다.
1월부터 전부 기재해 둔다." 실제로 확인해 보니 PO 발주글마다 이렇게 적혀 있다.

    Coupang이(가) 새 구매 오더(PO336120)를 전송했습니다.
    ★ 총금액 : 38,000,000 KRW
    ★ 품  목 : 송파3Sub-FC 메자닌리프트 2PL 1EA
    ★ 쿠팡오더 No. : PO336120/PR468820
    ★ 프로젝트 No. : UJ2600211

게다가 글머리에 **처리 상태**가 붙는다:
    ✅ 쿠팡오더처리 + 세금계산서 발행대기
    ⭐ 세금계산서 발행 2건 발행 … 2026.03.25 세금계산서 발행 완료
    ※ 쿠팡오더 금액 안맞아도 처리해도 된다고 확인 받음

이 상태를 안 읽으면 **이미 발행한 PO를 '미청구'라고 보고**하게 된다
(PO344599가 실제로 그랬다 — 3/25 발행 완료인데 미청구로 잡혔다).

  python po_band_status.py            # PO별로 밴드가 뭐라고 적어 뒀는지 보여준다
"""
import sys, os, re, json, glob

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

CACHE = os.path.join(ROOT, "band", "cache")
KAKAO = os.path.join(ROOT, "kakao", "inbox")
PO_RE = re.compile(r"PO\d{5,}")
UJ_RE = re.compile(r"UJ\d{7}")
DATE_RE = re.compile(r"(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})")

# 글에 적힌 처리 상태 — 앞에 있는 것일수록 확정적이다
STATES = [
    ("발행완료", re.compile(r"세금계산서\s*발행\s*완료|발행\s*완료")),
    ("발행대기", re.compile(r"세금계산서\s*발행\s*대기|발행대기")),
    ("오더처리", re.compile(r"쿠팡\s*오더\s*처리\s*완료|쿠팡오더처리")),
    ("금액이상", re.compile(r"금액\s*안맞|금액\s*불일치|금액\s*상이")),
]


def _bodies():
    out = []
    for f in glob.glob(os.path.join(CACHE, "*.json")):
        b = os.path.basename(f)
        if b.startswith(("dump_", "raw_")):
            continue
        try:
            d = json.load(open(f, encoding="utf-8"))
        except Exception:
            continue
        out += [(p.get("content") or "") for p in (d.get("posts") or {}).values()]
    for f in glob.glob(os.path.join(KAKAO, "*.txt")):
        try:
            out.append(open(f, encoding="utf-8", errors="replace").read())
        except OSError:
            pass
    return out


def scan():
    """PO번호 → {프로젝트들, 상태들, 발행일, 근거문구}"""
    info = {}
    for body in _bodies():
        # 발주·발행 글은 보통 한 덩어리다. 문단 단위로 잘라 PO와 그 주변만 본다.
        for blk in re.split(r"(?=Coupang이\(가\) 새 구매 오더|⭐|✅)", body):
            pos = PO_RE.findall(blk)
            if not pos:
                continue
            states = [n for n, rx in STATES if rx.search(blk)]
            ujs = UJ_RE.findall(blk)
            dm = DATE_RE.search(blk)
            issued = ("%s-%02d-%02d" % (dm.group(1), int(dm.group(2)), int(dm.group(3)))
                      if (dm and "발행완료" in states) else "")
            for po in set(pos):
                rec = info.setdefault(po, {"프로젝트": [], "상태": [], "발행일": "", "근거": ""})
                for u in ujs:
                    if u not in rec["프로젝트"]:
                        rec["프로젝트"].append(u)
                for s in states:
                    if s not in rec["상태"]:
                        rec["상태"].append(s)
                if issued and not rec["발행일"]:
                    rec["발행일"] = issued
                if states and not rec["근거"]:
                    rec["근거"] = re.sub(r"\s+", " ", blk)[:110]
    return info


def main():
    info = scan()
    print(f"밴드·카톡에서 확인한 PO {len(info)}건")
    done = {k: v for k, v in info.items() if "발행완료" in v["상태"]}
    wait = {k: v for k, v in info.items() if "발행대기" in v["상태"] and k not in done}
    odd = {k: v for k, v in info.items() if "금액이상" in v["상태"]}
    print(f"  발행완료 {len(done)} · 발행대기 {len(wait)} · 금액이상 메모 {len(odd)}")
    if "--all" in sys.argv:
        for po, v in sorted(info.items()):
            print(f"  {po} 상태={','.join(v['상태']) or '-'} 프로젝트={','.join(v['프로젝트'][:3]) or '-'}"
                  f" 발행일={v['발행일'] or '-'}")
    return info


if __name__ == "__main__":
    main()
