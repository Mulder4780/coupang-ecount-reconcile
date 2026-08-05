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


CHANGED_LOG = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                           "reports", "밴드_수정글.json")

# 밴드 화면 문구 — 수집 방식(피드 긁기 vs 상세 페이지)에 따라 붙었다 떨어졌다 한다.
# 이걸 걸러내지 않으면 **수집 방식만 바뀌어도 전부 '수정됨'으로 잡힌다**(실제로 571건
# 오탐이 났다). 글쓴이가 고친 것만 남기려면 화면 장식은 비교에서 빼야 한다.
_UI_NOISE = re.compile(
    r"(글 옵션|표정짓기|댓글쓰기|공동리더|\d+명이 읽었습니다|더보기|"
    r"메인 콘텐츠로 바로가기|BAND|밴드, 페이지, 게시글 검색|새글 피드|"
    r"새로운 새소식이[^\n]*|새로운 채팅 메시지[^\n]*|내 정보, 설정, 로그아웃|"
    r"게시글|사진첩|일정|첨부|멤버 \d+|초대|글쓰기|미션 인증 설정|밴드 설정|"
    r"밴드와 게시글이 공개되지 않습니다[^\n]*|검색|발견|\d+)")


def _norm(text):
    """비교용 정규화 — 화면 문구·공백·숫자 장식을 걷어낸 '사람이 쓴 내용'만 남긴다."""
    t = _UI_NOISE.sub(" ", str(text or ""))
    return re.sub(r"\s+", " ", t).strip()


def _mark_changed(band, nos):
    """수정된 글 번호를 남긴다 — 대조·사진수집이 이 목록을 다시 훑는다.

    사람이 글을 고치면 그 글에 딸린 판정(완료 여부·금액·사진)도 다시 봐야 한다.
    조용히 덮어쓰기만 하면 "언제 무엇이 바뀌었는지" 아무도 모른다.
    """
    doc = {"갱신": datetime.now().isoformat(timespec="seconds"), "밴드": {}}
    try:
        old = json.load(open(CHANGED_LOG, encoding="utf-8"))
        if isinstance(old.get("밴드"), dict):
            doc["밴드"] = old["밴드"]
    except Exception:
        pass
    cur = set(doc["밴드"].get(str(band)) or [])
    cur.update(str(n) for n in nos)
    doc["밴드"][str(band)] = sorted(cur, key=lambda x: (len(x), x))
    doc["합계"] = sum(len(v) for v in doc["밴드"].values())
    try:
        os.makedirs(os.path.dirname(CHANGED_LOG), exist_ok=True)
        json.dump(doc, open(CHANGED_LOG, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    except Exception:
        pass
    print(f"  ★ 수정된 글 {len(nos)}건 감지({band}) → reports/밴드_수정글.json")


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
                         "comment_count": p.get("comment_count", 0),
                         # ★ 사진 URL 보존(2026-08-05). 예전에는 여기서 images 를 버려
                         #   캐시에 URL 이 남지 않았고, 게시글 보관이 사진을 **0장** 받았다
                         #   (본문·사진수는 있는데 주소가 없어 내려받을 수가 없었다).
                         "images": [u for u in (p.get("images") or []) if u]}
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
        # ★ 밴드 글은 **수정된다**(2026-08-04 확인): 상태가 바뀌면 같은 글의 본문·사진을
        #   고쳐 다시 올린다. 예전 규칙("본문이 긴 쪽을 남긴다")은 **짧아지는 수정과
        #   같은 길이의 내용 변경을 통째로 놓쳤다.** 이제 수집 시각이 더 최신이면
        #   내용이 달라진 것을 교체하고, 무엇이 바뀌었는지 기록을 남긴다.
        changed = []
        cap_ms = int(cap or 0)
        for no, rec in posts.items():
            cur = merged.get(no)
            rec["captured_at"] = cap_ms or rec.get("captured_at") or 0
            if not cur:
                merged[no] = rec
                continue
            new_txt, old_txt = rec["content"] or "", cur.get("content") or ""
            # '…더보기'로 잘린 피드 수집분이 상세 전문을 덮어쓰지 않게 한다.
            truncated = len(new_txt) < len(old_txt) * 0.9 and old_txt.startswith(new_txt[:200])
            newer = rec["captured_at"] >= int(cur.get("captured_at") or 0)
            if rec["created_at"] and not cur.get("created_at"):
                merged[no] = rec
            elif len(new_txt) > len(old_txt) and not newer:
                merged[no] = rec                      # 예전 규칙(같은 회차 품질 차이)
            elif newer and not truncated and _norm(new_txt) != _norm(old_txt):
                rec["updated_at"] = cap_ms
                rec["prev_len"] = len(old_txt)
                merged[no] = rec
                changed.append(no)                    # 수정된 글 — 다시 대조해야 한다
            elif len(new_txt) > len(old_txt):
                merged[no] = rec
            # ★ 재수집 시각은 어느 분기가 이기든 **단조증가**로 남긴다(2026-08-04).
            #   덤프는 매 실행 전부 재처리되므로, 본문이 긴 옛 덤프가 나중에 이기면
            #   위 분기만으로는 스탬프가 0으로 되돌아가 recheck_plan 이 영원히
            #   '재수집 전'으로 보고 같은 글을 무한 반복한다.
            if not rec.get("images") and cur.get("images"):
                rec["images"] = cur["images"]          # 옛 수집분의 사진 주소를 잃지 않는다
            if merged.get(no) is rec and not rec.get("images") and cur.get("images"):
                merged[no]["images"] = cur["images"]
            keep = max(int(cur.get("captured_at") or 0), int(rec.get("captured_at") or 0))
            if keep:
                merged[no]["captured_at"] = keep
        if changed:
            _mark_changed(band, changed)
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
