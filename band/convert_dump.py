# -*- coding: utf-8 -*-
"""
convert_dump.py — 브라우저 수집 덤프(dump_*.json) → 대조 캐시(<band>.json) 변환
게시일 파싱 우선순위: 본문 1행 절대시각 → timeText 절대시각 → 상대시각(수집시각 기준).
변환 후 덤프는 raw_*.json 으로 개명 보존.
"""
import sys, os, re, json, glob
from datetime import datetime, timedelta

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache")
ABS = re.compile(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일\s*(오전|오후)?\s*(\d{1,2}):(\d{2})")


def parse_dt(text, captured_ms):
    m = ABS.search(text or "")
    if m:
        y, mo, d, ap, h, mi = m.groups()
        h = int(h) % 12 + (12 if ap == "오후" else 0)
        return datetime(int(y), int(mo), int(d), h, int(mi))
    base = datetime.fromtimestamp((captured_ms or 0) / 1000) if captured_ms else datetime.now()
    t = text or ""
    m = re.search(r"(\d+)분 전", t)
    if m: return base - timedelta(minutes=int(m.group(1)))
    m = re.search(r"(\d+)시간 전", t)
    if m: return base - timedelta(hours=int(m.group(1)))
    if "어제" in t: return base - timedelta(days=1)
    m = re.search(r"(\d+)일 전", t)
    if m: return base - timedelta(days=int(m.group(1)))
    return None


def main():
    for f in glob.glob(os.path.join(CACHE, "dump_*.json")):
        d = json.load(open(f, encoding="utf-8"))
        band = re.sub(r"\D", "", os.path.basename(f))
        cap = d.get("capturedAt")
        posts = {}
        for no, p in d.get("posts", {}).items():
            dt = parse_dt((p.get("content") or "").split("\n")[0], cap) \
                 or parse_dt(p.get("timeText"), cap)
            posts[no] = {"created_at": int(dt.timestamp() * 1000) if dt else None,
                         "author": p.get("author", ""),
                         "content": (p.get("content") or "")[:2000],
                         "photo_count": p.get("photo_count", 0),
                         "comment_count": p.get("comment_count", 0)}
        out = {"band_name": d.get("name", band), "posts": posts}
        json.dump(out, open(os.path.join(CACHE, f"{band}.json"), "w", encoding="utf-8"), ensure_ascii=False)
        os.replace(f, os.path.join(CACHE, f"raw_{band}.json"))
        dated = sum(1 for p in posts.values() if p["created_at"])
        print(f"{d.get('name', band)}: {len(posts)}건 변환 (날짜 파싱 {dated}건)")


if __name__ == "__main__":
    main()
