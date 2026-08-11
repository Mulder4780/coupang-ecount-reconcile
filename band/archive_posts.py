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
import concurrent.futures as cf
import threading
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
POST_WORKERS = 5       # 동시에 띄우는 크롬 수 — 글당 크롬 하나가 PDF 를 만든다
PHOTO_WORKERS = 8      # 한 글 안에서만 겹친다 — 밴드에 한꺼번에 몰지 않는다
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
    tmp = path + f".part-{os.getpid()}-{threading.get_ident()}"
    try:
        req = urllib.request.Request(url, headers=UA)
        with urllib.request.urlopen(req, timeout=30) as r:
            data = r.read()
        if len(data) < 2000:            # 썸네일·에러 페이지
            return "small"
        with open(tmp, "wb") as f:
            f.write(data)
        os.replace(tmp, path)
        return "ok"
    except Exception:
        return "fail"
    finally:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass


def archive_paths(root, no, post):
    """현재 글 내용 기준 보관 세트 경로를 한 곳에서 만든다.

    수정된 글은 같은 번호라도 파일명이 달라질 수 있다. 따라서 폴더 전체 파일 수나
    예전 개정본이 아니라 **현재 캐시가 가리키는 이 세트**가 완성됐는지를 본다.
    """
    base, day = base_name(no, post)
    ym = day[:7].replace("-", os.sep) if day != "날짜미상" else "날짜미상"
    folder = os.path.join(root, ym)
    photos = [os.path.join(folder, f"{base}_{i:02d}.jpg")
              for i, _url in enumerate(post.get("images") or [], 1)]
    return {
        "folder": folder,
        "txt": os.path.join(folder, base + ".txt"),
        "pdf": os.path.join(folder, base + ".pdf"),
        "photos": photos,
    }


def archive_inventory(root):
    """보관 경로를 SMB 왕복 한 번으로 읽는다."""
    present = set()
    if root and os.path.isdir(root):
        for current, _dirs, files in os.walk(root):
            for name in files:
                present.add(os.path.normcase(os.path.abspath(os.path.join(current, name))))
    return present


def archive_complete(paths, present=None):
    def has(path):
        if present is None:
            return os.path.exists(path)
        return os.path.normcase(os.path.abspath(path)) in present

    return has(paths["pdf"]) and has(paths["txt"]) and all(has(p) for p in paths["photos"])


def _atomic_text(path, text):
    tmp = path + f".part-{os.getpid()}-{threading.get_ident()}"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(text)
        os.replace(tmp, path)
    finally:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass


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
    tmp = pdf_path + f".part-{os.getpid()}-{threading.get_ident()}"
    try:
        made = html_to_pdf(html, tmp)
        if made and os.path.exists(tmp):
            os.replace(tmp, pdf_path)
            return pdf_path
        return None
    finally:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass


def archive_band(band, posts, limit, force, stat):
    """★ 글 단위로 **겹쳐서** 돌린다 (2026-08-08 실측).

    글 하나마다 헤드리스 크롬을 새로 띄워 PDF 를 만든다(프로필 폴더까지 새로).
    글당 20~30초라 순차로는 남은 3,400글에 하루가 걸린다. 대부분이 크롬을
    기다리는 시간이라 스레드로 겹치면 그대로 줄어든다 — 파이썬은 일을 안 한다.
    크롬을 몇 개까지 띄울지는 `POST_WORKERS` 가 정한다(메모리를 보고 잡은 값).
    """
    name = safe(posts.get("_band_name") or band, 30)
    root = os.path.join(out_root(), name)
    present = archive_inventory(root)
    nos = sorted((k for k in posts if str(k).isdigit()), key=lambda x: -int(x))
    todo = []
    for no in nos:
        post = posts[no]
        if not isinstance(post, dict):
            continue
        # ★ **시각 없는 글은 보관하지 않는다** (2026-08-08 실측 873건).
        #   밴드는 아직 없는 번호에도 앱 껍데기를 준다 — 그 화면을 뜯으면 직전 글
        #   본문이 그대로 잡히고 시각만 빈다(2026-08-07 사고). 수집기는 이미 그런
        #   수확을 버리는데(검증 [130]) 보관기가 다시 주워 담고 있었다.
        #   '날짜미상' 폴더에 873개가 쌓였고, 글 하나에 크롬을 한 번 띄우므로
        #   **없는 글에 세 시간을 썼다.** 게다가 모수에는 안 들어가서 아무리 만들어도
        #   '남음'이 줄지 않는다 — 영영 안 끝나는 일이 된다.
        if not post.get("created_at") or post.get("deleted"):
            continue
        paths = archive_paths(root, no, post)
        if force or not archive_complete(paths, present):
            todo.append(no)
    lock = threading.Lock()

    def one(no):
        post = posts[no]
        post = dict(post)
        post["_band"] = band
        paths = archive_paths(root, no, post)
        d = paths["folder"]
        os.makedirs(d, exist_ok=True)
        txt_p, pdf_p = paths["txt"], paths["pdf"]
        if not force and archive_complete(paths):
            with lock:
                stat["skip"] += 1
            return False
        # 1) 텍스트 — 검색·대조용 정본
        if force or not os.path.exists(txt_p):
            _atomic_text(
                txt_p,
                f"밴드: {name} ({band})\n글번호: {no}\n작성: {base_name(no, post)[1]}\n"
                f"글쓴이: {post.get('author') or ''}\n"
                f"사진: {post.get('photo_count') or 0} · 댓글: {post.get('comment_count') or 0}\n"
                f"원본: https://www.band.us/band/{band}/post/{no}\n"
                + "-" * 60 + "\n" + (post.get("content") or ""),
            )
        # 2) 사진 — 원본 화질로 글 옆에 둔다
        #    ★ **한 장씩 받으면 안 된다** (2026-08-08 실측). 글 하나에 사진이 열댓
        #      장이고, 네이버 CDN 왕복이 대부분 대기 시간이다. 순차로 돌리니
        #      40분에 16글밖에 못 갔다 — 남은 3,400글이면 며칠이다.
        #      기다리는 일이라 스레드로 겹치면 그대로 줄어든다(CPU 를 안 쓴다).
        #      한 글 안에서만 겹친다 — 밴드 서버에 한꺼번에 몰지 않기 위해서다.
        urls = list(enumerate(post.get("images") or [], 1))
        photos, got = [None] * len(urls), {}
        if urls:
            with cf.ThreadPoolExecutor(max_workers=PHOTO_WORKERS) as ex:
                futs = {}
                for i, url in urls:
                    p = paths["photos"][i - 1]
                    futs[ex.submit(fetch_photo, url, p)] = (i, p)
                for fu in cf.as_completed(futs):
                    i, p = futs[fu]
                    try:
                        got[i] = (fu.result(), p)
                    except Exception:
                        got[i] = ("fail", p)
        for i, p in urls:                      # 순서는 원본 순서 그대로 지킨다
            r, path = got.get(i, ("fail", ""))
            if r in ("ok", "skip"):
                photos[i - 1] = path
                if r == "ok":
                    with lock:
                        stat["photo"] += 1
        photos = [x for x in photos if x]
        # 3) PDF — 사람이 보던 모습 그대로 고정
        # PDF·텍스트가 이미 있어도 빠진 사진은 다시 받는다. 새 사진이 생겼으면 PDF도
        # 다시 굳혀야 사람이 보는 고정본과 사진 세트가 같은 상태가 된다.
        should_render = force or not os.path.exists(pdf_p) or any(
            result == "ok" for result, _path in got.values())
        ok = bool(render_pdf(no, post, name, photos, pdf_p)) if should_render else True
        complete_now = ok and archive_complete(paths)
        with lock:
            stat.setdefault("incomplete", 0)
            stat["made" if complete_now else ("pdf_fail" if not ok else "incomplete")] += 1
        return True

    # ★ 상한은 **새로 만든 글** 기준이다. 이미 있는 것은 세지 않는다 —
    #   그러면 매 회차가 앞부분만 다시 훑고 끝나 영영 뒤로 못 간다.
    with cf.ThreadPoolExecutor(max_workers=POST_WORKERS) as ex:
        futs, it = set(), iter(todo)
        submitted = 0
        while True:
            while len(futs) < POST_WORKERS and submitted < limit:
                try:
                    futs.add(ex.submit(one, next(it)))
                    submitted += 1
                except StopIteration:
                    break
            if not futs:
                break
            done, futs = cf.wait(futs, return_when=cf.FIRST_COMPLETED)
            # `limit`은 완성 성공 수가 아니라 이번 회차가 손댄 미완성 글 수다. CDN에서
            # 영구 실패하는 사진 몇 장이 앞에 있어도 그 글들만 무한 재시도하지 않는다.
            if submitted >= limit and not futs:
                break


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--band")
    ap.add_argument("--limit", type=int, default=200)
    ap.add_argument("--force", action="store_true")
    a = ap.parse_args()

    stat = {"made": 0, "skip": 0, "photo": 0, "pdf_fail": 0, "incomplete": 0}
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
          f"사진 {stat['photo']}장 · 미완 {stat['incomplete']} · "
          f"PDF실패 {stat['pdf_fail']} → {out_root()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
