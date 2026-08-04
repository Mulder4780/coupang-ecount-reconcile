#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""auto_scout.py — GitHub/오픈소스 자동화 후보 일일 검색 (매일 10:00 스케줄러)

왜(2026-08-04 사용자 지시): "깃허브나 오픈 소스를 활용한 자동화 방안을 항상 검색해서
자동으로 매일 오전 10시부터 실행하도록 해."
사람이/AI가 기억해서 찾게 하지 않는다 — 프로젝트의 반복 업무(밴드 수집·ERP 연동·
엑셀 zip 패치·카톡 정리·크롬 자동화·세금계산서)에 맞는 도구가 새로 나오면
reports/자동화_후보.md 에 ★신규 로 표시되어 다음 세션 AI·사람이 본다.

- 비인증 GitHub Search API(분당 10회 제한)라 쿼리 사이 7초 쉼 — 비밀키 불필요.
- 상태(db/auto_scout_seen.json)에 본 저장소를 남겨 '신규'만 도드라지게 한다.
- 실패(오프라인·레이트리밋)해도 exit 0 — 스케줄러가 오류로 남지 않게 하고
  보고서에 실패 사실을 쓴다.
콘솔은 한 줄 요약만(토큰 절약 규칙), 상세는 reports/자동화_후보.md.
"""
import json
import os
import sys
import time
import urllib.parse
import urllib.request

ROOT = os.path.dirname(os.path.abspath(__file__))
SEEN = os.path.join(ROOT, "db", "auto_scout_seen.json")
REPORT = os.path.join(ROOT, "reports", "자동화_후보.md")

# 검색어는 "지금 손이 가는 반복 업무"에서 나온다 — 새 반복 업무가 생기면 여기 추가.
QUERIES = [
    ("밴드 수집(무한스크롤·API)", "naver band api"),
    ("이카운트 ERP 연동", "ecount erp"),
    ("엑셀 서식 보존 편집(zip 패치)", "excel edit preserve charts python"),
    ("카톡 내보내기 파싱", "kakaotalk export parser"),
    ("로그인된 크롬 자동화(CDP)", "chrome devtools protocol attach existing session python"),
    ("세금계산서·홈택스 자동화", "hometax tax invoice automation"),
    ("폴더 감시·자동 흡수", "python watchdog folder intake pipeline"),
    ("윈도우 작업 스케줄러 관리", "windows scheduled task python manage"),
]


def gh_search(q, per_page=5):
    url = ("https://api.github.com/search/repositories?"
           + urllib.parse.urlencode({"q": q, "sort": "updated", "per_page": per_page}))
    req = urllib.request.Request(url, headers={
        "User-Agent": "coupang-work-agent-scout",
        "Accept": "application/vnd.github+json",
    })
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.load(r).get("items", [])


def main():
    seen = {}
    if os.path.exists(SEEN):
        try:
            seen = json.load(open(SEEN, encoding="utf-8"))
        except Exception:
            seen = {}
    today = time.strftime("%Y-%m-%d")
    lines = [f"# 자동화 후보 (GitHub/오픈소스 일일 검색) — {today}", ""]
    new_cnt = total = 0
    errors = []
    for i, (pain, q) in enumerate(QUERIES):
        if i:
            time.sleep(7)  # 비인증 검색 분당 10회 제한
        try:
            items = gh_search(q)
        except Exception as e:
            errors.append(f"{q}: {e}")
            continue
        lines.append(f"## {pain}  (`{q}`)")
        for it in items:
            name = it.get("full_name", "?")
            total += 1
            is_new = name not in seen
            if is_new:
                new_cnt += 1
                seen[name] = {"first_seen": today, "query": q}
            star = "★신규 " if is_new else ""
            desc = (it.get("description") or "").strip()[:120]
            lines.append(f"- {star}[{name}]({it.get('html_url','')}) "
                         f"☆{it.get('stargazers_count',0)} — {desc}")
        lines.append("")
    if errors:
        lines += ["## 검색 실패", *[f"- {e}" for e in errors], ""]
    lines.append("> 적용 판단은 세션 AI가 한다 — 후보를 코드에 넣기 전에 합성검증·"
                 "절대규칙(무차별 API 탐침 금지 등)을 그대로 따른다.")
    os.makedirs(os.path.dirname(REPORT), exist_ok=True)
    open(REPORT, "w", encoding="utf-8").write("\n".join(lines))
    os.makedirs(os.path.dirname(SEEN), exist_ok=True)
    json.dump(seen, open(SEEN, "w", encoding="utf-8"), ensure_ascii=False)
    print(f"자동화 후보: 신규 {new_cnt}건 / 조회 {total}건 · 실패 {len(errors)}쿼리 "
          f"→ reports/자동화_후보.md")
    return 0


if __name__ == "__main__":
    sys.exit(main())
