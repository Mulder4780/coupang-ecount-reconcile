# -*- coding: utf-8 -*-
"""
archive_posts.py — 밴드 게시글을 **글 단위로** 보관한다 (PDF·텍스트·사진 한 세트)

사용자 지시(2026-08-05): "밴드도 각 게시글을 pdf출력 보관 텍스트 보관 사진 및 첨부파일
보관 하는 알고리즘 추가."

무엇을 만드나 — 글 하나당 파일 세트(이름만 봐도 무슨 글인지 알게):
    <밴드이름>/2026/07/[5321]_2026-07-14_UJ2600895_돌발AS.pdf     ← 사람이 보는 고정본
                                        …_UJ2600895_돌발AS.txt    ← 본문 전문(검색용)
                                        …_UJ2600895_돌발AS_01.jpg ← 사진(원본 화질)
  · 대괄호 안이 **밴드 글번호**(원본 추적 키), 그 뒤가 날짜·프로젝트NO·업무유형.
  · 프로젝트NO 는 본문에서 UJ 패턴을 찾아 넣는다. 없으면 자리를 비운다(지어내지 않는다).

원본은 어디서 오나: `band/cache/<밴드>.json` (수집기가 만든 정본 캐시).
사진은 캐시의 images URL 을 서버에서 직접 받는다(브라우저는 CORS 로 막힌다).

  python band/archive_posts.py                 # 새 글만 (이미 있으면 건너뜀)
  python band/archive_posts.py --band 90610953 # 한 밴드만
  python band/archive_posts.py --limit 50      # 이번 회차 최대 건수(기본 200)
  python band/archive_posts.py --force         # 이미 있어도 다시 만든다
"""
import argparse
import glob
import json
import os
import re
import sys
import time
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

CACHE = os.path.join(HERE, "cache")
UA = {"User-Agent": "Mozilla/5.0", "Referer": "https://www.band.us/"}
UJ = re.compile(r"UJ\d{7}")
KIND = [("정기점검", "정기점검"), ("돌발", "돌발AS"), ("납품", "신규납품"),
        ("설치", "설치"), ("철거", "철거"), ("계단", "계단"), ("점검", "점검")]


def out_root():
    import source_dirs as S
    return os.path.join(S.BAND_DIR, "게시글보관")


def safe(s, n=40):
    s = re.sub(r"[\\/:*?\"<>|\r\n\t]", " ", str(s or "")).strip()
    return re.sub(r"\s+", " ", s)[:n]


def kind_of(text):
    for key, label in KIND:
        if key in (text or ""):
            return label
    return ""


def base_name(no, post):
    ms = post.get("created_at")
    day = time.strftime("%Y-%m-%d", time.localtime(ms / 1000)) if ms else "날짜미상"
    content = post.get("content") or ""
    uj = (UJ.search(content) or [None])
    uj = uj.group(0) if hasattr(uj, "group") else ""
    parts = [f"[{no}]", day]
    if uj:
        parts.append(uj)
    k = kind_of(content)
    if k:
        parts.append(k)
    return "_".join(parts), day


def fetch_photo(url, path):
    if os.path.exists(path):
        return "skip"
    try:
        req = urllib.request.Request(url, headers=UA)
        with urllib.request.urlopen(req, timeout=30) as r:
            data = r.read()
        if len(data) < 2000:            # 썸네일·에러 페이지
            return "small"
        with open(path, "wb") as f:
            f.write(data)
        return "ok"
    except Exception:
        return "fail"


def render_pdf(no, post, band_name, photos, pdf_path):
    from archive_render import html_to_pdf, esc
    ms = post.get("created_at")
    when = time.strftime("%Y-%m-%d %H:%M", time.localtime(ms / 1000)) if ms else "(시각 미상)"
    imgs = "".join(f'<img src="file:///{p.replace(chr(92), "/")}">' for p in photos)
    html = f"""
<h1>{esc(band_name)} — 글 {no}</h1>
<div class="meta">
  <b>작성</b> {esc(when)} &nbsp;|&nbsp; <b>글쓴이</b> {esc(post.get('author') or '(미상)')}
  &nbsp;|&nbsp; <b>사진</b> {len(photos)}장 &nbsp;|&nbsp; <b>댓글</b> {esc(post.get('comment_count') or 0)}
</div>
<pre>{esc(post.get('content') or '(본문 없음)')}</pre>
<div class="photos">{imgs}</div>
<div class="foot">밴드 원본: https://www.band.us/band/{esc(post.get('_band'))}/post/{no}
 · 보관 생성 {time.strftime('%Y-%m-%d %H:%M')} · 이 파일은 자동 생성본이며 원본을 수정하지 않는다.</div>
"""
    return html_to_pdf(html, pdf_path)


def archive_band(band, posts, limit, force, stat):
    name = safe(posts.get("_band_name") or band, 30)
    root = os.path.join(out_root(), name)
    nos = sorted((k for k in posts if str(k).isdigit()), key=lambda x: -int(x))
    for no in nos:
        if stat["made"] >= limit:
            return
        post = posts[no]
        if not isinstance(post, dict):
            continue
        post["_band"] = band
        base, day = base_name(no, post)
        ym = day[:7].replace("-", os.sep) if day != "날짜미상" else "날짜미상"
        d = os.path.join(root, ym)
        os.makedirs(d, exist_ok=True)
        txt_p = os.path.join(d, base + ".txt")
        pdf_p = os.path.join(d, base + ".pdf")
        if not force and os.path.exists(pdf_p) and os.path.exists(txt_p):
            stat["skip"] += 1
            continue
        # 1) 텍스트 — 검색·대조용 정본
        with open(txt_p, "w", encoding="utf-8") as f:
            f.write(f"밴드: {name} ({band})\n글번호: {no}\n작성: {day}\n"
                    f"글쓴이: {post.get('author') or ''}\n"
                    f"사진: {post.get('photo_count') or 0} · 댓글: {post.get('comment_count') or 0}\n"
                    f"원본: https://www.band.us/band/{band}/post/{no}\n"
                    + "-" * 60 + "\n" + (post.get("content") or ""))
        # 2) 사진 — 원본 화질로 글 옆에 둔다
        photos = []
        for i, url in enumerate(post.get("images") or [], 1):
            p = os.path.join(d, f"{base}_{i:02d}.jpg")
            r = fetch_photo(url, p)
            if r in ("ok", "skip"):
                photos.append(p)
                if r == "ok":
                    stat["photo"] += 1
        # 3) PDF — 사람이 보던 모습 그대로 고정
        if render_pdf(no, post, name, photos, pdf_p):
            stat["made"] += 1
        else:
            stat["pdf_fail"] += 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--band")
    ap.add_argument("--limit", type=int, default=200)
    ap.add_argument("--force", action="store_true")
    a = ap.parse_args()

    stat = {"made": 0, "skip": 0, "photo": 0, "pdf_fail": 0}
    files = [os.path.join(CACHE, f"{a.band}.json")] if a.band else \
        [f for f in glob.glob(os.path.join(CACHE, "*.json"))
         if os.path.basename(f)[:-5].isdigit()]
    for f in sorted(files):
        if not os.path.exists(f):
            continue
        band = os.path.basename(f)[:-5]
        doc = json.load(open(f, encoding="utf-8"))
        posts = doc.get("posts") or {}
        posts["_band_name"] = doc.get("band_name") or band
        archive_band(band, posts, a.limit, a.force, stat)
    print(f"밴드 게시글 보관: 새로 {stat['made']}건 · 건너뜀 {stat['skip']} · "
          f"사진 {stat['photo']}장 · PDF실패 {stat['pdf_fail']} → {out_root()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
