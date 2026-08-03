# -*- coding: utf-8 -*-
"""
convert_dump.py — 브라우저 수집 덤프(dump_*.json) → 대조 캐시(<band>.json) 변환
게시일 파싱 우선순위: 본문 1행 절대시각 → timeText 절대시각 → 상대시각(수집시각 기준).
변환 후 덤프는 raw_*.json 으로 개명 보존.
"""
import sys, os, re, json, glob
import hashlib
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


def dump_files():
    """로컬 처리함과 0. 원본 자료의 밴드 JSON 정본을 함께 읽는다."""
    paths = list(glob.glob(os.path.join(CACHE, "dump_*.json")))
    try:
        sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from source_dirs import band_dump_dirs
        for folder in band_dump_dirs():
            paths.extend(glob.glob(os.path.join(folder, "**", "*.json"), recursive=True))
    except Exception:
        pass
    out, seen = [], set()
    for path in paths:
        key = os.path.normcase(os.path.abspath(path))
        if key not in seen:
            seen.add(key)
            out.append(path)
    return sorted(out)


def main():
    for f in dump_files():
        d = json.load(open(f, encoding="utf-8"))
        if not isinstance(d, dict) or not isinstance(d.get("posts"), (dict, list)):
            continue
        # 밴드번호는 **파일명 맨 뒤 숫자 덩어리**다. 전체에서 숫자만 뽑으면
        # dump_api2_90610953 → "290610953" 처럼 앞의 버전 숫자가 섞여 다른 밴드가 된다.
        nums = re.findall(r"(\d{6,})", os.path.basename(f))
        band = str(d.get("band") or (nums[-1] if nums else hashlib.sha256(
            os.path.basename(f).encode("utf-8")).hexdigest()[:10]))
        cap = d.get("capturedAt")
        posts = {}
        source_posts = d.get("posts") or {}
        iterator = source_posts.items() if isinstance(source_posts, dict) else enumerate(source_posts)
        for no, p in iterator:
            if not isinstance(p, dict):
                continue
            # 밴드 API로 받은 덤프는 created_at(ms)을 이미 갖고 있다 — 본문에서 다시 캐낼 필요가 없다.
            # (화면 긁기 덤프만 본문·timeText에서 시각을 파싱한다)
            ms = p.get("created_at")
            dt = None if ms else (parse_dt((p.get("content") or "").split("\n")[0], cap)
                                  or parse_dt(p.get("timeText"), cap))
            posts[no] = {"created_at": int(ms) if ms else (int(dt.timestamp() * 1000) if dt else None),
                         "author": p.get("author", ""),
                         # ★ 2000자로 자르면 **목록형 글이 통째로 잘린다**. 실제로 "미실시 및 AS 진행건 공유"
                    #   글(4,288자)에서 프로젝트NO 36개 중 18개가 사라졌다(2026-07-27).
                    #   한도는 폭주 방지용으로만 남긴다 — 실제 글은 1만 자를 넘지 않는다.
                    "content": (p.get("content") or "")[:20000],
                         "photo_count": p.get("photo_count", 0),
                         "comment_count": p.get("comment_count", 0)}
        # ★ 기존 캐시에 **덮어쓰지 않고 합친다**.
        #   수집 방식마다 커버하는 기간이 달라(화면 긁기=과거, API=최근) 덮어쓰면
        #   한쪽 기간이 통째로 사라진다(2026-07-26에 12~4월이 날아갔다).
        dst = os.path.join(CACHE, f"{band}.json")
        merged, before = {}, 0
        if os.path.exists(dst):
            try:
                old = json.load(open(dst, encoding="utf-8"))
                merged = old.get("posts") or {}
                before = len(merged)
            except Exception:
                merged = {}
        for no, rec in posts.items():
            cur = merged.get(no)
            # 날짜가 있는 쪽·본문이 긴 쪽을 남긴다
            if not cur or (rec["created_at"] and not cur.get("created_at"))                or len(rec["content"]) > len(cur.get("content") or ""):
                merged[no] = rec
        out = {"band_name": d.get("name", band), "posts": merged}
        json.dump(out, open(dst, "w", encoding="utf-8"), ensure_ascii=False)
        # 원본 자료 정본은 이름·내용 그대로 둔다. 로컬 처리함의 dump만 raw로 바꿔
        # 다음 실행에서 반복 변환되지 않게 한다.
        try:
            in_cache = os.path.commonpath([os.path.abspath(f), os.path.abspath(CACHE)]) == os.path.abspath(CACHE)
        except ValueError:  # C: 처리함과 Z: 원본처럼 드라이브가 다르면 공통경로가 없다.
            in_cache = False
        if in_cache:
            raw = os.path.join(CACHE, f"raw_{os.path.basename(f)[5:-5]}.json")
            os.replace(f, raw)
        dated = sum(1 for p in merged.values() if p["created_at"])
        print(f"{d.get('name', band)}: {len(posts)}건 반영 → 캐시 {before}→{len(merged)}건 "
              f"(날짜 있는 글 {dated}건)")


if __name__ == "__main__":
    main()
