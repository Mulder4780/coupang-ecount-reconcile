# -*- coding: utf-8 -*-
"""
export_watch.py — 카톡 [대화 내보내기]가 얼마나 오래됐나 감시 (읽기 전용)

왜(2026-08-04 사용자 지시): "내보내기가 오래되면(1일 이상) 대시보드와 아침 리포트에
'카톡 내보내기 오래됨 — 갱신 필요' 경고를 띄워라."
카톡은 읽기 API가 없고 후킹은 약관위반 금지라 내보내기만이 공식 경로다(AGENTS.md).
그 한 번의 사람 손을 사람이 기억하게 하지 않고, 오래되면 시스템이 먼저 말하게 한다.

- 기준: 카톡 정본 폴더(Z: 3. 카카오톡 내보내기 + kakao/inbox)의 최신 .txt mtime.
- 결과: reports/카톡_내보내기_경과.json
  { newest_file, newest_at, age_hours, stale }  · stale = 24시간 초과
- 소비처: webapp/app_server.py(대표보고 리스크 섹션), daily_brief.py(아침 브리핑 머리).
"""
import glob
import json
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

STALE_HOURS = 24  # 사용자 지시: 1일 이상이면 경고
OUT = os.path.join(ROOT, "reports", "카톡_내보내기_경과.json")


def check():
    import source_dirs as S
    newest, newest_ts = None, 0.0
    for d in S.kakao_dirs():
        if not os.path.isdir(d):
            continue
        for p in glob.glob(os.path.join(d, "**", "*.txt"), recursive=True):
            try:
                ts = os.path.getmtime(p)
            except OSError:
                continue
            if ts > newest_ts:
                newest, newest_ts = p, ts
    age_h = (time.time() - newest_ts) / 3600 if newest_ts else None
    doc = {
        "newest_file": os.path.basename(newest) if newest else None,
        "newest_at": time.strftime("%Y-%m-%d %H:%M", time.localtime(newest_ts)) if newest_ts else None,
        "age_hours": round(age_h, 1) if age_h is not None else None,
        "stale": bool(age_h is None or age_h > STALE_HOURS),
        "checked_at": time.strftime("%Y-%m-%d %H:%M"),
    }
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump(doc, open(OUT, "w", encoding="utf-8"), ensure_ascii=False)
    return doc


def main():
    d = check()
    if d["age_hours"] is None:
        print("카톡 내보내기: 파일 없음 — 갱신 필요")
    else:
        mark = "★ 오래됨(갱신 필요)" if d["stale"] else "최신"
        print(f"카톡 내보내기: {d['newest_at']} ({d['age_hours']}시간 전) — {mark}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
