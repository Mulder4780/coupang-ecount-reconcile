# -*- coding: utf-8 -*-
"""
band_sync.py — 밴드 게시글 자동 수집 (증분 동기화)
====================================================
- /v2.1/bands 로 가입 밴드 목록 조회 → config band.targets(이름 부분일치)로 대상 선정(비면 전체)
- /v2/band/posts 를 paging.next_params 로 역페이징, 캐시(cache/{band_key}.json)에 증분 저장
- 트래픽 예의: 페이지당 0.7초 간격, 실행당 밴드별 최대 max_pages(기본 30)페이지
- 이미 수집된 게시글(post_key)에 도달하면 그 페이지에서 중단 → 재실행 비용 최소화

실행:  python band_sync.py            # 증분 수집
       python band_sync.py --list     # 밴드 목록만 표시(수집 없음)
"""
import sys, os, json, time, ssl, urllib.request, urllib.parse

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(os.path.dirname(BASE_DIR), "config", "ecount_config.json")
TOKEN_PATH = os.path.join(BASE_DIR, ".band_token.json")
CACHE_DIR = os.path.join(BASE_DIR, "cache")
API = "https://openapi.band.us"

cfg = json.load(open(CONFIG_PATH, encoding="utf-8"))
band_cfg = cfg.get("band", {})
DELAY = float(band_cfg.get("page_delay_sec", 0.7))
MAX_PAGES = int(band_cfg.get("max_pages_per_run", 30))
TARGETS = [t.strip() for t in band_cfg.get("targets", []) if t.strip()]

try:
    TOKEN = json.load(open(TOKEN_PATH, encoding="utf-8"))["access_token"]
except Exception:
    sys.exit("토큰이 없습니다. 먼저  python band_auth.py  로 1회 인증하세요.")

_ctx = ssl.create_default_context()

def get(path, params):
    url = API + path + "?" + urllib.parse.urlencode(params)
    with urllib.request.urlopen(urllib.request.Request(url), timeout=30, context=_ctx) as r:
        d = json.loads(r.read().decode("utf-8"))
    if d.get("result_code") != 1:
        raise RuntimeError(f"BAND API 오류 {d.get('result_code')}: {json.dumps(d, ensure_ascii=False)[:200]}")
    return d["result_data"]

def api_comments(item):
    """목록 API 가 실어 주는 최근 댓글을 캐시 모양으로 옮긴다 (2026-08-08).

    캐시 모양은 화면 긁기(convert_dump)와 **한 글자도 다르면 안 된다** —
    읽는 쪽(band_extract.comment_text → cancel_watch)이 한 곳이기 때문이다.
    시각 없는 댓글은 여기서도 버린다: 순서를 못 세우면 취소 판정의 근거가 없다.
    """
    out = []
    for c in (item.get("latest_comments") or item.get("comments") or []):
        if not isinstance(c, dict):
            continue
        body = str(c.get("content") or c.get("body") or "").strip()
        ms = c.get("created_at")
        if not body or not ms:
            continue
        out.append({"author": str((c.get("author") or {}).get("name") or "").strip(),
                    "created_at": int(ms), "content": body[:2000]})
    return out


def merge_comments(old, new):
    """댓글은 쌓이는 것이라 **합친다**(같은 사람·같은 시각·같은 말은 한 번만)."""
    seen, out = set(), []
    for c in list(old or []) + list(new or []):
        if not isinstance(c, dict) or not c.get("created_at") or not c.get("content"):
            continue
        key = (c.get("author") or "", int(c["created_at"]), c["content"])
        if key in seen:
            continue
        seen.add(key)
        out.append({"author": key[0], "created_at": key[1], "content": key[2]})
    out.sort(key=lambda c: c["created_at"])
    return out


def main():
    bands = get("/v2.1/bands", {"access_token": TOKEN}).get("bands", [])
    print(f"가입 밴드 {len(bands)}개:")
    for b in bands:
        mark = "★" if (not TARGETS or any(t in b["name"] for t in TARGETS)) else " "
        print(f"  {mark} {b['name']}  (band_key={b['band_key'][:12]}...)")
    if "--list" in sys.argv:
        return

    os.makedirs(CACHE_DIR, exist_ok=True)
    sel = [b for b in bands if not TARGETS or any(t in b["name"] for t in TARGETS)]
    for b in sel:
        key = b["band_key"]
        cpath = os.path.join(CACHE_DIR, f"{key}.json")
        try:
            cache = json.load(open(cpath, encoding="utf-8"))
        except Exception:
            cache = {"band_name": b["name"], "posts": {}}
        known = set(cache["posts"].keys())
        params = {"access_token": TOKEN, "band_key": key, "locale": "ko-KR"}
        new_cnt, page = 0, 0
        while page < MAX_PAGES:
            page += 1
            data = get("/v2/band/posts", params)
            items = data.get("items", [])
            hit_known = False
            for it in items:
                pk = it.get("post_key")
                if not pk:
                    continue
                if pk in known:
                    hit_known = True
                    # ★ 아는 글이라고 그냥 넘기면 **댓글이 영영 안 들어온다** (2026-08-08).
                    #   댓글은 글보다 나중에 달린다 — 접수 취소 통보가 바로 그 모양이다.
                    #   그래서 글 자체는 그대로 두고 댓글만 합친다(있을 때만).
                    cm = api_comments(it)
                    if cm:
                        old = cache["posts"].get(pk) or {}
                        old["comments"] = merge_comments(old.get("comments"), cm)
                        old["comment_count"] = it.get("comment_count", old.get("comment_count", 0))
                        cache["posts"][pk] = old
                    continue
                cache["posts"][pk] = {
                    "created_at": it.get("created_at"),
                    "author": (it.get("author") or {}).get("name", ""),
                    # ★ 2000자로 자르면 **목록형 글이 통째로 잘린다**. 실제로 "미실시 및 AS 진행건 공유"
                    #   글(4,288자)에서 프로젝트NO 36개 중 18개가 사라졌다(2026-07-27).
                    #   한도는 폭주 방지용으로만 남긴다 — 실제 글은 1만 자를 넘지 않는다.
                    "content": (it.get("content") or "")[:20000],
                    "photo_count": len(it.get("photos") or []),
                    "comment_count": it.get("comment_count", 0),
                    # ★ 댓글 본문 (2026-08-08). 취소 통보는 대부분 댓글로 온다.
                    "comments": api_comments(it),
                    # 목록 API 는 최근 댓글만 준다 — 적힌 수만큼 왔을 때만 '다 읽었다'.
                    "comments_full": len(api_comments(it)) >= int(it.get("comment_count", 0) or 0),
                }
                new_cnt += 1
            nxt = (data.get("paging") or {}).get("next_params")
            if hit_known or not nxt or not items:
                break
            params = dict(nxt); params["access_token"] = TOKEN
            time.sleep(DELAY)
        json.dump(cache, open(cpath, "w", encoding="utf-8"), ensure_ascii=False)
        print(f"  → '{b['name']}': 신규 {new_cnt}건 (누적 {len(cache['posts'])}건, {page}페이지 조회)")
    print("수집 완료. 다음: python band_reconcile.py")

if __name__ == "__main__":
    main()
