# -*- coding: utf-8 -*-
"""
erp_pdf_export.py — ERP 내보내기(엑셀)를 **PDF 사본**으로 굳혀 원본 저장소에 남긴다
================================================================================
사용자 지시(2026-08-04): "ERP 모든 데이터 다운로드 받아서 PDF 또는 이미지 파일로
저장하고 원본 데이터 저장소에 전부 반영해."

**왜 PDF 사본인가**: 엑셀은 열 때마다 수식·서식이 흔들리고, ERP 재조회 시점에 따라
숫자가 달라진다. PDF 는 "그때 ERP 가 이렇게 보여 줬다"를 고정한다 — 사후 대조·감사에
쓰는 것은 결국 이 고정본이다. 엑셀 원본은 그대로 두고 짝만 만든다.

**변환 방법**: 설치된 Excel 을 PowerShell COM 으로 띄워 ExportAsFixedFormat 한다
(pywin32 없이 동작). LibreOffice 도 없고 파이썬 PDF 라이브러리도 안 쓴다 —
엑셀이 그리는 그대로가 사람이 보던 화면이기 때문이다.

저장 위치: 원본과 같은 날짜 폴더 아래 `PDF/` (예: 1. ERP 내보내기/2026/08/2026-08-04/PDF)
파일명: `<종류>_<원본이름>.pdf` — ERP 다운로드 이름이 무작위(0NSKITA3APTYVRL)라
        inbox_scan 의 내용 분류를 붙여 사람이 찾을 수 있게 한다.

실행
  python erp_pdf_export.py            # 새로 생긴 것만 변환
  python erp_pdf_export.py --all      # 이미 있어도 다시 만든다
  python erp_pdf_export.py --status   # 변환 현황만
"""
import glob
import os
import subprocess
import sys
import time
import io
import json
import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

KIND_LABEL = {
    "sales": "판매조회", "tax": "매출계산서현황", "taxinv": "계산서진행단계",
    "hometax": "홈택스전자계산서", "ledger": "거래처별계정별원장",
    "ledger_acct": "계정별원장", "journal": "분개장", "cashbook": "현금출납장",
    "receivable": "거래처별채권", "collect": "수금현황", "quote": "견적서",
    "purchase": "구매조회", "unknown": "기타",
}


# ── 진도 자국 ──────────────────────────────────────────────────────────
# 2026-09-09 실사고: 이 단계가 29분을 쓰고 시간초과(1800s)로 끊겼는데
# **어디에 썼는지 물을 데가 없었다**. 자식 stdout 은 파이프에 물리면 블록
# 버퍼라 죽을 때 버퍼째 사라지고([427]), 그래서 남은 것이 종료코드뿐이었다.
# ★ 죽어도 남는다. 그게 요점이다([180]).
TRACE = os.path.join(ROOT, "reports", ".erp_pdf_진행.json")
_TRACE_EVERY_S = 5.0   # 상한은 개수가 아니라 **시간**으로 정한다([437])
_TRACE = {"때": "", "시작": "", "구간": "", "구간초": 0.0,
          "원본": 0, "본파일": 0, "대상": 0,
          "분류초": 0.0, "폴더초": 0.0, "대조초": 0.0, "pid": os.getpid()}
_SEG_T0 = [0.0]
_LAST = [0.0]


def _note(force=False):
    """자국 하나로 이 단계를 죽이지 않는다 — 못 남겨도 PDF 는 계속 만든다."""
    now = time.time()
    if not force and now - _LAST[0] < _TRACE_EVERY_S:
        return
    _LAST[0] = now
    _TRACE["때"] = datetime.datetime.now().isoformat(timespec="seconds")
    _TRACE["구간초"] = round(now - _SEG_T0[0], 1)
    try:
        tmp = TRACE + ".tmp"
        with io.open(tmp, "w", encoding="utf-8") as f:
            json.dump(_TRACE, f, ensure_ascii=False)
        os.replace(tmp, TRACE)          # 반쪽 파일을 안 남긴다([171])
    except Exception:
        pass


def _seg(name):
    if not _TRACE["시작"]:
        _TRACE["시작"] = datetime.datetime.now().isoformat(timespec="seconds")
    _SEG_T0[0] = time.time()
    _TRACE["구간"] = name
    _note(True)


def _sources():
    from source_dirs import ERP_DIR
    out = []
    for p in sorted(glob.glob(os.path.join(ERP_DIR, "**", "*.xlsx"), recursive=True)):
        base = os.path.basename(p)
        if base.startswith("~$") or os.sep + "PDF" + os.sep in p:
            continue
        out.append(p)
    return out


def _target(src):
    """원본과 같은 날짜 폴더 아래 PDF/ 로. 종류를 앞에 붙여 이름으로 찾을 수 있게."""
    folder = os.path.join(os.path.dirname(src), "PDF")
    stem = os.path.splitext(os.path.basename(src))[0]
    try:
        import inbox_scan
        kind = inbox_scan.classify_cached(src) or "unknown"
    except Exception:
        kind = "unknown"
    label = KIND_LABEL.get(kind, kind or "기타")
    return os.path.join(folder, f"{label}_{stem}.pdf")


def _convert(pairs):
    """PowerShell 로 Excel 한 번만 띄워 여러 파일을 연속 변환한다.

    파일마다 Excel 을 새로 띄우면 파일당 3~5초가 더 든다. 한 인스턴스를 재사용하고
    실패한 파일은 건너뛴 뒤 계속 간다 — 한 장 때문에 전체가 멈추면 안 된다.
    """
    if not pairs:
        return []
    # ★ 파일이 많으면 -Command 인자가 윈도우 한계(32KB)를 넘어 WinError 206 으로 죽는다
    #   (2026-08-04, 거래명세서 37개 흡수 때 실측). 10개씩 나눠 Excel 을 여러 번 띄운다 —
    #   느려지지만 한 번에 다 넣어 통째로 실패하는 것보다 낫다.
    if len(pairs) > 10:
        out = []
        for i in range(0, len(pairs), 10):
            out += _convert(pairs[i:i + 10])
        return out
    lines = [
        "$ErrorActionPreference='Continue'",
        "$xl = New-Object -ComObject Excel.Application",
        "$xl.Visible=$false; $xl.DisplayAlerts=$false",
        "$done=@()",
    ]
    for src, dst in pairs:
        s = src.replace("'", "''")
        d = dst.replace("'", "''")
        lines += [
            "try {",
            f"  $wb = $xl.Workbooks.Open('{s}', 0, $true)",
            # 가로 방향 + 폭 1장 맞춤: ERP 표는 열이 많아 세로로 뽑으면 잘린다.
            "  foreach ($ws in $wb.Worksheets) {",
            "    $ws.PageSetup.Orientation = 2",
            "    $ws.PageSetup.Zoom = $false",
            "    $ws.PageSetup.FitToPagesWide = 1",
            "    $ws.PageSetup.FitToPagesTall = $false",
            "  }",
            f"  $wb.ExportAsFixedFormat(0, '{d}')",
            "  $wb.Close($false)",
            f"  $done += '{d}'",
            "} catch { $wb = $null }",
        ]
    lines += [
        "$xl.Quit()",
        "[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null",
        "$done -join \"`n\"",
    ]
    script = "\n".join(lines)
    res = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                         capture_output=True, text=True, timeout=1800, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    return [ln.strip() for ln in (res.stdout or "").splitlines() if ln.strip().endswith(".pdf")]


def run(force=False, quiet=False):
    # ★ 세 갈래를 갈라 잰다 — 조치가 갈래마다 다르다([289]).
    #   분류초 = _target() 안 inbox_scan.classify_cached (캐시가 빗나가면 Z: 워크북을 연다)
    #   폴더초 = os.makedirs · 대조초 = exists + getmtime
    _seg("원본 목록 만들기(glob)")
    srcs = _sources()
    _TRACE["원본"] = len(srcs)
    _seg("고를 것 가리기")
    pairs = []
    for i, src in enumerate(srcs):
        t = time.time()
        dst = _target(src)
        _TRACE["분류초"] = round(_TRACE["분류초"] + time.time() - t, 1)
        t = time.time()
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        _TRACE["폴더초"] = round(_TRACE["폴더초"] + time.time() - t, 1)
        t = time.time()
        skip = (not force and os.path.exists(dst)
                and os.path.getmtime(dst) >= os.path.getmtime(src))
        _TRACE["대조초"] = round(_TRACE["대조초"] + time.time() - t, 1)
        _TRACE["본파일"] = i + 1
        _note()
        if skip:
            continue
        pairs.append((src, dst))
    _TRACE["대상"] = len(pairs)
    _seg("PDF 변환(Excel 띄우기)")
    made = _convert(pairs)
    _seg("끝")
    ok = sum(1 for _, d in pairs if os.path.exists(d))
    if not quiet:
        print(f"ERP PDF 사본 — 원본 {len(srcs)}개 · 대상 {len(pairs)}개 · 생성 {ok}개")
        print(f"  구간초 — 분류 {_TRACE['분류초']} · 폴더 {_TRACE['폴더초']}"
              f" · 대조 {_TRACE['대조초']} (자국 {TRACE})")
        for _, d in pairs[:8]:
            mark = "✓" if os.path.exists(d) else "✗"
            print(f"  {mark} {os.path.basename(d)}")
        if len(pairs) > 8:
            print(f"  … 외 {len(pairs) - 8}개")
    return {"sources": len(srcs), "targets": len(pairs), "made": ok, "files": made}


def status():
    srcs = _sources()
    have = [s for s in srcs if os.path.exists(_target(s))]
    print(f"ERP 원본 {len(srcs)}개 중 PDF 사본 {len(have)}개")
    for s in srcs:
        t = _target(s)
        print(f"  {'✓' if os.path.exists(t) else ' '} {os.path.basename(t)}")


def main():
    if "--status" in sys.argv:
        status()
        return
    run(force="--all" in sys.argv)


if __name__ == "__main__":
    main()
