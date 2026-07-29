# -*- coding: utf-8 -*-
"""
fetch_images.py — 밴드 게시글의 거래명세서·세금계산서 **사진을 내려받는다**
================================================================================
수집 덤프(dump_*.json)에는 사진 URL만 담겨 있다. 브라우저 안에서는 CORS 때문에
이미지를 못 가져오지만(네이버 CDN), **서버에서 직접 요청하면 받아진다**(확인 완료).

  · 대상: 본문에 명세서·계산서·견적·청구·세금 이 들어간 글의 사진
  · 저장: band/docs_inbox/  → doc_ocr.py 가 Windows 내장 OCR로 금액·번호를 읽는다
  · 썸네일(type=s75)은 건너뛴다 — 75px라 글자가 안 읽힌다

실행
  python band/fetch_images.py            # 받기
  python band/fetch_images.py --dry      # 대상만 세어 보기
"""
import sys, os, re, json, glob, time, hashlib, urllib.request

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, "cache")
OUT = os.path.join(HERE, "docs_inbox")
DOC = re.compile(r"명세서|계산서|견적|청구|세금")
UA = {"User-Agent": "Mozilla/5.0", "Referer": "https://www.band.us/"}


def post_year(p, body):
    """글의 연도 — created_at 우선, 없으면 본문의 'YYYY년'."""
    t = p.get("created_at") or 0
    try:
        t = int(t)
        if t > 1e12:
            t //= 1000
        if t > 0:
            import datetime
            return datetime.datetime.fromtimestamp(t).year
    except Exception:
        pass
    m = re.search(r"(20\d{2})년", body[:200])
    return int(m.group(1)) if m else 0


def targets(all_photos=False, year=0):
    """[(밴드, 게시글ID, 날짜, URL)]

    기본은 **문서로 보이는 글**의 사진만(OCR 대상). `--all` 이면 그 글의 사진을 전부 받는다
    (사용자 지시 2026-07-29: 빼먹지 말고 다 가져와). `--year 2026` 이면 그 해 글만.
    ★ 캐시에 URL 이 있는 것만 받을 수 있다 — 밴드가 말하는 장수보다 적다.
      나머지는 공식 API(심사 대기) 가 있어야 한다. 브라우저로는 못 긁는다(AGENTS.md 참고)."""
    out = []
    for f in sorted(glob.glob(os.path.join(CACHE, "dump_*.json")) +
                    glob.glob(os.path.join(CACHE, "raw_*.json"))):
        try:
            d = json.load(open(f, encoding="utf-8"))
        except Exception:
            continue
        band = str(d.get("band") or os.path.basename(f))
        posts = d.get("posts") or {}
        for pid, p in (posts.items() if isinstance(posts, dict) else enumerate(posts)):
            body = p.get("content") or ""
            if not all_photos and not DOC.search(body):
                continue
            if year and post_year(p, body) != year:
                continue
            m = re.search(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일", body)
            day = "%s%02d%02d" % (m.group(1)[2:], int(m.group(2)), int(m.group(3))) if m else "000000"
            for u in (p.get("images") or []):
                if "type=s75" in u:            # 75px 썸네일 — OCR 불가
                    continue
                out.append((band, str(pid), day, u))
    return out


def out_dir():
    """저장 위치 — 원본은 서버('0. 원본 자료')가 정본이다(2026-07-28 이전 완료).
    서버가 안 붙어 있으면 로컬 inbox 로 떨어뜨린다."""
    try:
        sys.path.insert(0, os.path.dirname(HERE))
        from source_dirs import DOC_PHOTO_DIRS
        d = DOC_PHOTO_DIRS[0]
        os.makedirs(d, exist_ok=True)
        return d
    except Exception:
        return OUT


def main():
    dry = "--dry" in sys.argv
    all_photos = "--all" in sys.argv
    year = 0
    if "--year" in sys.argv:
        try:
            year = int(sys.argv[sys.argv.index("--year") + 1])
        except (IndexError, ValueError):
            year = 0
    global OUT
    OUT = out_dir()
    os.makedirs(OUT, exist_ok=True)
    t = targets(all_photos, year)
    print("대상 사진 %d장 (%s%s)" % (
        len(t), "글 전체" if all_photos else "문서로 보이는 글",
        (" · %d년" % year) if year else ""))
    print("  저장 위치:", OUT)
    if dry:
        return
    got = skip = fail = 0
    for i, (band, pid, day, url) in enumerate(t, 1):
        key = hashlib.md5(url.encode()).hexdigest()[:8]
        ext = ".png" if ".png" in url.lower() else ".jpg"
        dst = os.path.join(OUT, f"band{band}_{day}_{pid}_{key}{ext}")
        if os.path.exists(dst):
            skip += 1
            continue
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=30) as r:
                b = r.read()
            if len(b) < 8000:                  # 너무 작으면 문서가 아니다
                skip += 1
                continue
            open(dst, "wb").write(b)
            got += 1
        except Exception:
            fail += 1
        if i % 50 == 0:
            print(f"  {i}/{len(t)} — 받음 {got} · 건너뜀 {skip} · 실패 {fail}")
        time.sleep(0.12)                       # 서버 예의
    print(f"완료: 받음 {got} · 건너뜀 {skip} · 실패 {fail} → {OUT}")


if __name__ == "__main__":
    main()
