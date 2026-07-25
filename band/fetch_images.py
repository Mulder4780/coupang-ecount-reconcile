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


def targets():
    """[(밴드, 게시글ID, 날짜, URL)] — 문서로 보이는 글의 큰 사진만"""
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
            if not DOC.search(body):
                continue
            m = re.search(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일", body)
            day = "%s%02d%02d" % (m.group(1)[2:], int(m.group(2)), int(m.group(3))) if m else "000000"
            for u in (p.get("images") or []):
                if "type=s75" in u:            # 75px 썸네일 — OCR 불가
                    continue
                out.append((band, str(pid), day, u))
    return out


def main():
    dry = "--dry" in sys.argv
    os.makedirs(OUT, exist_ok=True)
    t = targets()
    print(f"대상 사진 {len(t)}장 (문서로 보이는 글)")
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
