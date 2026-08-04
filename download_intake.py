# -*- coding: utf-8 -*-
"""download_intake.py — 다운로드·바탕화면에 떨어진 업무 파일을 '0. 원본 자료'로 흡수한다.

사용자 지시(2026-07-31): "모든 자료는 '0. 원본 자료'에 다운로드 받거나 복사해서 저장.
세션이 바뀌어도 다른 PC에서 작업해도 모두 반영되게."

왜 필요한가 — 브라우저(ERP Excel·밴드 덤프)와 카톡 내보내기는 **Downloads/바탕화면에**
떨어진다. 오늘 실제로 사람이(그리고 AI가) 손으로 옮겼다. 손으로 옮기는 것은 세션이
바뀌면 잊힌다 — 스케줄러에 넣은 것만 산다.

원칙
  · 판별은 파일명이 아니라 **내용**(inbox_scan.classify)으로 한다 — ERP 내보내기는
    무작위 파일명(URIWDQIBP1N9PWU.xlsx)으로 온다.
  · 목적지는 source_dirs 가 정한 곳 하나뿐이다. 여기서 경로를 새로 정하지 않는다.
  · Z: 로 **이동**한다(복사 아님) — Downloads 에 사본이 남으면 다음에 또 흡수한다.
    Z: 가 끊겨 있으면 아무것도 하지 않는다(다음 실행이 처리).
  · 여러 PC 반영: 목적지가 공유폴더(Z:)라서, 옮겨지는 순간 모든 PC·세션이 본다.
    무엇을 언제 옮겼는지는 reports/download_intake.json 에 남긴다.
  · 최근 7일 파일만 본다 — Downloads 의 해묵은 잡동사니를 건드리지 않는다.
"""
from __future__ import annotations

import glob
import json
import os
import re
import shutil
import sys
import time
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
REPORT = os.path.join(ROOT, "reports", "download_intake.json")
RECENT_DAYS = 7

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _sources():
    home = os.path.expanduser("~")
    return [os.path.join(home, "Downloads"), os.path.join(home, "Desktop"),
            # OneDrive 가 바탕화면을 가로채는 PC 도 있다 — 있는 것만 실제로 쓰인다
            os.path.join(home, "OneDrive", "Desktop"),
            os.path.join(home, "OneDrive", "바탕 화면")]


def _recent(path):
    try:
        return (time.time() - os.path.getmtime(path)) < RECENT_DAYS * 86400
    except OSError:
        return False


def _dated(base):
    now = datetime.now()
    return os.path.join(base, f"{now.year}", f"{now.month:02d}", now.strftime("%Y-%m-%d"))


def _erp_filename(name):
    """이카운트 내보내기 파일명인가 — 무작위 대문자 15자, 또는 프로그램코드(ETA002R 등).

    이카운트는 화면에 따라 두 가지로 이름을 준다. 둘 다 사람이 지을 이름은 아니라서
    이 형태면 ERP 산출물로 봐도 안전하다.
    """
    stem = os.path.splitext(str(name or ""))[0]
    return bool(re.fullmatch(r"[A-Za-z0-9]{12,20}", stem)
                or re.fullmatch(r"E[A-Z]*\d{3,6}[A-Z]?", stem)
                or re.fullmatch(r"ECTAX\d+[A-Z]?", stem))


def plan_moves():
    """무엇을 어디로 옮길지 계산만 한다(이동은 apply 에서)."""
    import source_dirs as S
    from inbox_scan import classify

    if not os.path.isdir(S.ORIGIN_ROOT):
        return None            # Z: 끊김 — 판단 불가면 아무것도 하지 않는다

    moves = []
    for src_dir in _sources():
        if not os.path.isdir(src_dir):
            continue
        # 1) 카톡 내보내기 — 이름이 곧 판별이다(KakaoTalk_*.txt)
        for p in glob.glob(os.path.join(src_dir, "KakaoTalk*.txt")):
            if _recent(p):
                moves.append((p, _dated(S.KAKAO_DIR), "카톡 내보내기"))
        # 2) 밴드 브라우저 덤프 — dump_<밴드번호>.json
        for p in glob.glob(os.path.join(src_dir, "dump_*.json")):
            if _recent(p):
                moves.append((p, os.path.join(S.BAND_DIR, "브라우저덤프",
                                              datetime.now().strftime("%Y-%m-%d")), "밴드 덤프"))
        # 3) 엑셀 — 내용으로 판별한다. 업무 파일로 판별된 것만 가져간다.
        for p in glob.glob(os.path.join(src_dir, "*.xlsx")):
            if not _recent(p) or os.path.basename(p).startswith("~$"):
                continue
            try:
                kind = classify(p)
            except Exception:
                continue
            if kind in ("taxinv", "ledger", "sales", "tax", "stmt", "slips", "hometax", "po", "receipt"):
                base = S.COUPANG_DIR if kind == "po" else (S.RECEIPT_DIR if kind == "receipt" else S.ERP_DIR)
                moves.append((p, _dated(base), f"내용판별({kind})"))
            elif _erp_filename(os.path.basename(p)):
                # 내용 판별이 안 되는 ERP 화면(요약표·집계표)도 원본은 남긴다
                # (사용자 지시 2026-08-04 "ERP 모든 데이터 ... 전부 반영").
                # **파일명이 ERP 다운로드 형태일 때만** 가져온다 — Downloads 의 개인 파일을
                # Z: 로 쓸어 담지 않기 위해서다.
                moves.append((p, _dated(S.ERP_DIR), "ERP 파일명 판별(내용 미상)"))
    return moves


def apply(moves, dry=False):
    done, failed = [], []
    for src, dst_dir, why in moves or []:
        name = os.path.basename(src)
        try:
            os.makedirs(dst_dir, exist_ok=True)
            dst = os.path.join(dst_dir, name)
            if os.path.exists(dst):                      # 같은 이름은 덮지 않는다
                stem, ext = os.path.splitext(name)
                dst = os.path.join(dst_dir, f"{stem}__dl{datetime.now():%H%M%S}{ext}")
            if not dry:
                shutil.move(src, dst)
            done.append({"파일": name, "이유": why, "목적지": dst})
        except OSError as exc:
            failed.append({"파일": name, "오류": str(exc)[:120]})
    return done, failed


def main():
    dry = "--apply" not in sys.argv
    moves = plan_moves()
    if moves is None:
        print("원본 자료 폴더(Z:)에 닿지 않습니다 — 이번 회차는 건너뜁니다.")
        return 0
    done, failed = apply(moves, dry=dry)
    os.makedirs(os.path.dirname(REPORT), exist_ok=True)
    payload = {"time": datetime.now().isoformat(timespec="seconds"),
               "dry": dry, "이동": done, "실패": failed}
    tmp = REPORT + f".{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)
    os.replace(tmp, REPORT)
    mode = "미리보기" if dry else "이동"
    print(f"다운로드 흡수({mode}): {len(done)}건" + (f" · 실패 {len(failed)}" if failed else ""))
    for d in done[:8]:
        print(f"  [{d['이유']}] {d['파일']}")
    if dry and done:
        print("  실제 이동: python download_intake.py --apply")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
