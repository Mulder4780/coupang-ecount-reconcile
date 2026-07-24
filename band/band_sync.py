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
                    continue
                cache["posts"][pk] = {
                    "created_at": it.get("created_at"),
                    "author": (it.get("author") or {}).get("name", ""),
                    "content": (it.get("content") or "")[:2000],
                    "photo_count": len(it.get("photos") or []),
                    "comment_count": it.get("comment_count", 0),
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
