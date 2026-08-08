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


def conv_comments(raw, captured_ms):
    """덤프의 댓글을 캐시 모양 `{author, created_at, content}` 로 옮긴다 (2026-08-08).

    ★ **시각이 안 잡히면 버린다.** 글 본문과 같은 규칙이다 — 밴드는 아직 안 그려진
      자리에도 껍데기를 주므로 시각 없는 수확은 직전 화면이 묻어 온 것일 수 있다.
      취소 판정은 '댓글이 글보다 나중'이라는 순서를 근거로 삼는데, 시각이 없으면
      그 순서 자체를 세울 수 없다.
    ★ 같은 사람이 같은 말을 같은 시각에 두 번 남길 수는 없다 — 회차가 겹쳐 들어와도
      중복은 여기서 접는다(캐시 합치기가 댓글을 더하기만 하면 계속 불어난다).
    """
    out, seen = [], set()
    for c in (raw or []):
        if not isinstance(c, dict):
            continue
        body = str(c.get("content") or c.get("body") or "").strip()
        if not body:
            continue
        ms = c.get("created_at")
        if not ms:
            dt = parse_dt(c.get("timeText"), captured_ms)
            ms = int(dt.timestamp() * 1000) if dt else None
        if not ms:
            continue
        author = str(c.get("author") or "").strip()
        key = (author, int(ms), body)
        if key in seen:
            continue
        seen.add(key)
        out.append({"author": author, "created_at": int(ms), "content": body[:2000]})
    out.sort(key=lambda c: c["created_at"])
    return out


def known_bands():
    """캐시에 이미 있는 밴드번호 — 파일명이 애매할 때의 가장 좋은 근거다."""
    try:
        return {f[:-5] for f in os.listdir(CACHE)
                if f.endswith(".json") and f[:-5].isdigit()}
    except OSError:
        return set()


def band_from_name(basename, known=None):
    """파일명에서 밴드번호를 고른다.

    ★ **맨 뒤 숫자를 집으면 안 된다** (2026-08-08 실사고). 수집본 파일명에 날짜
      꼬리표가 붙는다 — `84789192_260807.json` 의 `260807` 은 2026-08-07 이다.
      맨 뒤를 집은 탓에 **두 밴드가 `260807` 이라는 없는 밴드 하나로 합쳐졌고**,
      캐시에 5,453글짜리 유령 밴드가 생겼다. 아무도 이상하다 하지 않았다 —
      글도 있고 날짜도 있고 개수도 그럴듯했기 때문이다. 재수집 회차가 그 밴드에
      붙여넣기 파일까지 만들어 놓고 나서야 드러났다(있지도 않은 밴드를 긁으라고).
    ★ 그렇다고 맨 앞도 아니다. 예전 사고는 반대 방향이었다 —
      `dump_api2_90610953` 에서 앞의 버전 숫자가 섞이면 다른 밴드가 된다.
    그래서 자리(앞/뒤)가 아니라 **무엇처럼 생겼는가**로 고른다:
      ① 이미 캐시에 있는 밴드번호가 후보에 있으면 그것 — 가장 확실한 근거
      ② 없으면 **가장 긴** 숫자 덩어리 — 밴드번호는 8자리, 날짜 꼬리표는 6자리다
    ★ ②의 '가장 긴' 은 **위로도 막아야 한다** (2026-08-08 두 번째 실사고).
      `dump_202608082047_null.json` 의 `202608082047` 은 12자리 **시각 도장**인데
      6자리 날짜보다 길어서 ②가 그것을 골랐다 — 유령 밴드 `202608082047` 이
      캐시에 생겼고, 그 빈 캐시가 `make_oneclick` 을 첫 밴드에서 죽여 **모든 밴드의
      붙여넣기 파일이 하나도 안 만들어졌다.** 앞 사고와 방향만 반대일 뿐 같은 일이다.
      밴드번호는 8자리다. 그러니 후보를 **8자리에 가까운 것**으로 좁힌다.
    """
    nums = re.findall(r"(\d{6,})", basename)
    if not nums:
        return None
    known = known_bands() if known is None else known
    for n in nums:
        if n in known:
            return n
    # 밴드번호로 있을 수 있는 길이만 남긴다(관측된 밴드는 전부 8자리 — 7~10 만 허용).
    #   날짜 꼬리표(6자리)도 시각 도장(12·14자리)도 여기서 함께 떨어진다.
    plausible = [n for n in nums if 7 <= len(n) <= 10]
    if not plausible:
        return None          # 모르면 **모른다고 한다** — 없는 밴드를 만드는 것보다 낫다
    longest = max(len(n) for n in plausible)
    return [n for n in plausible if len(n) == longest][-1]


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
# ★ "밴드가 조용한 것"과 "수집이 막힌 것"을 가르는 근거 (2026-08-07 지시).
#   신선도 판정은 '날짜 있는 최신 글'만 봐서, 밴드에 새 글이 없는 날에도 "★밀림"이라고
#   외쳤다. 그 경보를 믿고 없는 번호를 긁다가 40건이 전부 같은 글로 들어온 것이 오늘 사고다.
#   근거로 쓸 수 있는 것은 **missing 뿐**이다 — 밴드가 "삭제되었거나 찾을 수 없습니다"라고
#   명시한 번호다. failed/no-time 은 화면이 안 그려졌을 때도 나오므로(오늘이 그랬다)
#   '없음'의 증거가 되지 못한다. 증거를 좁게 잡는 편이 거짓 안심보다 낫다.
PROBE_LOG = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                         "reports", "밴드_확인시각.json")


def _absent_above(top, missing, notime):
    """수집 최대 번호(top) 위로 '아직 없는 글'로 확인된 번호들.

    두 가지가 증거가 된다.
      · `missing` — 밴드가 '삭제되었거나 찾을 수 없습니다'라고 명시한 번호.
      · `notime` 중 **지문이 서로 같은 것이 2개 이상** — 아직 없는 번호를 열면 밴드가
        200 과 앱 껍데기를 주고 그 자리에 **직전 화면 본문**이 그대로 남는다. 그래서
        없는 번호끼리는 본문 지문이 똑같다(2026-08-07 실측: 3539~3578 마흔 건이 전부
        같은 글이었다). 하나만 있으면 '화면이 늦게 그려진 것'과 구분되지 않으므로
        **2개 이상 같은 지문일 때만** 증거로 친다.
    """
    out = set(int(n) for n in missing if str(n).isdigit() and int(n) > top)
    sigs = {}
    for no, sig in (notime or {}).items():
        if str(no).isdigit() and int(no) > top and sig:
            sigs.setdefault(sig, []).append(int(no))
    for sig, nos in sigs.items():
        if len(nos) >= 2:
            out.update(nos)
    return sorted(out)


# ★ 리다이렉트 실패를 '삭제된 글'로 판정하는 데 필요한 **서로 다른 회차** 수 (2026-08-07).
#   1 이면 안 된다. 로그인이 풀렸거나 네트워크가 끊긴 회차는 **모든 번호**가 리다이렉트로
#   실패하는데, 그 한 번을 근거로 묘비를 세우면 멀쩡한 글을 통째로 지운 것으로 적는다.
#   되돌릴 수 없는 판정이므로 증거는 좁게 잡는다 — CLAUDE.md "실패는 삭제의 증거가 아니다".
REDIRECT_ROUNDS_FOR_DELETED = 2


def _feed_sigs(notime):
    """한 회차에서 **피드 껍데기**로 확인된 본문 지문들.

    없는 번호(또는 지워진 번호)를 열면 밴드는 200 과 앱 껍데기를 주고 그 자리에
    직전 화면 — 즉 **피드 맨 위 글** — 이 그대로 남는다. 그래서 그런 번호끼리는
    본문 지문이 서로 똑같다. 같은 지문이 2개 이상이면 그것이 '피드 껍데기'다.

    지문이 저 혼자면 증거로 치지 않는다. 그건 '화면이 늦게 그려진 진짜 글'과
    구분되지 않는다.
    """
    sigs = {}
    for no, sig in (notime or {}).items():
        if str(no).isdigit() and sig:
            sigs.setdefault(sig, []).append(int(no))
    return {s for s, nos in sigs.items() if len(nos) >= 2}


def _redirect_hits(notime, ok_count):
    """이 회차가 '리다이렉트로 확인'한 번호들 → set.

    `ok_count` 는 이 회차에서 **실제로 수확된 글 수**다. 0 이면 아무것도 돌려주지
    않는다 — 한 건도 못 받은 회차는 밴드가 아니라 **이쪽이 고장난 회차**이고,
    그런 회차의 실패는 무엇의 증거도 되지 못한다. 이 한 줄이 없으면 로그인이 풀린
    밤 한 번으로 수천 건이 묘비를 쓴다.
    """
    if not ok_count:
        return set()
    feed = _feed_sigs(notime)
    return {int(no) for no, sig in (notime or {}).items()
            if str(no).isdigit() and sig in feed}


def _mark_redirect_deleted(band, merged, rounds):
    """여러 회차가 같은 번호를 리다이렉트로 확인했으면 **묘비를 세운다** → 세운 수.

    왜 필요한가 (2026-08-07, 분담판 [13])
      밴드 구멍 9건(3525·3397·3378·3374·3373·2598·2597·2595·2573)이 매 회차 9/9 로
      실패하는데 아무 데도 안 적혀서, 다음 회차 계획이 **또 같은 9건을 뽑았다.**
      `missing`(밴드가 '삭제되었거나 찾을 수 없습니다'라고 명시)만 묘비를 세우고
      리다이렉트 실패는 세우지 않았기 때문이다. 그 조심성 자체는 옳았다 —
      실패는 삭제의 증거가 아니다. 다만 **한 번의 실패**가 증거가 아닐 뿐,
      서로 다른 날 · 서로 다른 회차가 같은 번호에서 같은 모양으로 실패하고
      그 회차들이 다른 글은 멀쩡히 받아 왔다면, 그것은 이야기가 다르다.

    `rounds` = {번호: {회차 캡처시각, ...}}
    """
    if not rounds:
        return 0
    n = 0
    for no, whens in rounds.items():
        if len(whens) < REDIRECT_ROUNDS_FOR_DELETED:
            continue
        key = str(no)
        cur = merged.get(key)
        # 본문을 받아 둔 진짜 글은 절대 건드리지 않는다. 리다이렉트가 여러 번 났어도
        # 손에 본문이 있으면 그건 있는 글이다 — 여기서 지우면 되돌릴 수 없다.
        if isinstance(cur, dict) and (cur.get("created_at") or cur.get("content")):
            continue
        if isinstance(cur, dict) and cur.get("deleted"):
            continue
        last = max(whens)
        merged[key] = {"deleted": True, "deleted_at": last,
                       "captured_at": max(int((cur or {}).get("captured_at") or 0), last),
                       "deleted_by": "redirect",
                       "why": f"서로 다른 회차 {len(whens)}번이 피드 리다이렉트로 확인 "
                              f"— 지워진 글(분담판 [13], 2026-08-07)"}
        n += 1
    return n


def _record_probe(band, name, merged, missing, cap_ms, notime=None):
    """이 회차가 '번호 N 위로는 글이 없다'를 증명했으면 적어 둔다.

    성립 조건은 하나뿐이다: **수집 최대 번호 바로 다음 번호가 '없음'으로 확인**될 것.
    중간에 건너뛴 채 위쪽만 없음이면 그 사이를 모르므로 증거가 아니다.
    """
    real = [int(k) for k, v in merged.items()
            if str(k).isdigit() and isinstance(v, dict) and not v.get("deleted")]
    if not real:
        return
    top = max(real)
    absent = _absent_above(top, missing, notime)
    if not absent or absent[0] != top + 1:
        return
    when = (datetime.fromtimestamp(cap_ms / 1000).strftime("%Y-%m-%d %H:%M")
            if cap_ms else "")
    if not when:
        return
    try:
        doc = json.load(open(PROBE_LOG, encoding="utf-8"))
    except Exception:
        doc = {}
    prev = doc.get(str(band)) or {}
    if str(prev.get("확인시각") or "") > when:
        return                        # 더 최근 확인이 이미 있으면 옛 회차로 덮지 않는다
    doc[str(band)] = {"이름": name, "확인시각": when, "수집최대": top,
                      "없음확인": absent[0], "연속없음": len(absent)}
    os.makedirs(os.path.dirname(PROBE_LOG), exist_ok=True)
    tmp = PROBE_LOG + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, PROBE_LOG)

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
    # ★ 리다이렉트 실패는 **회차를 가로질러** 세야 뜻이 생긴다 (분담판 [13]).
    #   한 덤프만 보면 "이번에 실패했다"밖에 모른다. 덤프는 매 실행 전부 재처리되므로
    #   여기 담아 두면 이 한 번의 실행이 곧 '여러 회차를 본 것'이 된다.
    #   {밴드: {번호: {회차 캡처시각, ...}}}
    redirect_rounds = {}
    for f in dump_files():
        d = json.load(open(f, encoding="utf-8"))
        if not isinstance(d, dict) or not isinstance(d.get("posts"), (dict, list)):
            continue
        band = str(d.get("band") or band_from_name(os.path.basename(f)) or
                   hashlib.sha256(os.path.basename(f).encode("utf-8")).hexdigest()[:10])
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
                         # ★ 댓글 본문 (2026-08-08). 취소 통보는 대부분 댓글로 온다.
                         #   시각 없는 댓글은 버린다 — 본문과 같은 규칙이다.
                         "comments": conv_comments(p.get("comments"), cap),
                         # 적힌 개수만큼 못 읽었으면 그 사실을 남긴다. 이것이 없으면
                         # '댓글 없음'과 '못 읽음'이 캐시에서 똑같아 보인다(조용한 사고).
                         "comments_full": bool(p.get("comments_full")),
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
            # ★ 오염 표시(clean_contaminated)는 **날짜 없는 재병합이 못 덮는다** (2026-08-07).
            #   덤프는 매 실행 전부 재처리된다. 가짜(피드 리다이렉트) 본문을 담은 옛 덤프가
            #   Z: 에 그대로 있어서, 표시를 해 두어도 다음 회차가 도로 가짜를 살려냈다
            #   (실측: 표시 621건이 한 회차 만에 0건). 작성일을 **가진** 기록만 표시를
            #   뚫을 수 있다 — 그건 진짜 글을 제대로 다시 모았다는 뜻이므로 복구가 맞다.
            if cur and cur.get("contaminated") and not rec.get("created_at"):
                continue
            if not cur:
                merged[no] = rec
                continue
            # ★ 댓글은 **합친다** — 어느 분기가 이기든 잃지 않는다 (2026-08-08).
            #   덤프는 매 실행 전부 재처리되므로, 댓글을 못 담던 시절의 옛 덤프가
            #   나중에 이기면 애써 모은 댓글이 통째로 사라진다. 본문과 달리 댓글은
            #   '고쳐지는' 것이 아니라 **쌓이는** 것이라 합치는 쪽이 언제나 옳다.
            #   (conv_comments 가 같은 사람·같은 시각·같은 말을 접는다)
            rec["comments"] = conv_comments(
                (cur.get("comments") or []) + (rec.get("comments") or []), cap_ms)
            rec["comments_full"] = bool(rec.get("comments_full") or cur.get("comments_full"))
            new_txt, old_txt = rec["content"] or "", cur.get("content") or ""
            # '…더보기'로 잘린 피드 수집분이 상세 전문을 덮어쓰지 않게 한다.
            truncated = len(new_txt) < len(old_txt) * 0.9 and old_txt.startswith(new_txt[:200])
            newer = rec["captured_at"] >= int(cur.get("captured_at") or 0)
            if rec.get("created_at") and not cur.get("created_at"):
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
        # ★ 삭제된 글에 묘비를 세운다 (2026-08-05).
        #   밴드는 지운 글을 열면 '삭제됨' 안내 대신 **밴드 홈 화면**을 돌려준다.
        #   그래서 수집기가 본문을 못 찾고, recheck_plan 은 캐시의 옛 기록만 보고
        #   "아직 재수집 안 됐다"며 **영원히 같은 번호를 다시 뽑았다**(실제로 4건이
        #   모든 회차에서 반복 실패했다). 없는 글은 없다고 적어야 목록이 줄어든다.
        for no in (d.get("missing") or []):
            no = str(no)
            rec = merged.get(no) or {}
            rec["deleted"] = True
            rec["deleted_at"] = cap_ms
            rec["captured_at"] = max(int(rec.get("captured_at") or 0), cap_ms)
            merged[no] = rec
        # 이 회차가 리다이렉트로 확인한 번호를 회차 시각과 함께 쌓아 둔다.
        # 판정은 모든 덤프를 다 본 **뒤에** 한다 — 한 회차만 보고 묘비를 세우지 않는다.
        try:
            ok_here = sum(1 for p in posts.values() if p.get("created_at"))
            for no in _redirect_hits(d.get("notime") or {}, ok_here):
                redirect_rounds.setdefault(band, {}).setdefault(no, set()).add(cap_ms)
        except Exception:
            pass
        if changed:
            _mark_changed(band, changed)
        gone = len(d.get("missing") or [])
        if gone:
            print(f"  · 삭제된 글 {gone}건 기록({band}) — 다음 회차부터 다시 훑지 않는다")
        # 이 회차가 '수집 최대 번호 위로는 글이 없다'를 증명했으면 남긴다.
        # 신선도 판정이 이것을 보고 '밀림'과 '조용함'을 가른다.
        try:
            _record_probe(band, d.get("name", band), merged, d.get("missing") or [],
                          cap_ms, d.get("notime") or {})
        except Exception:
            pass
        # ★ 오염 표시는 **병합 때마다** 다시 매긴다 (2026-08-07 두 번째 실사고).
        #   위 226행 가드는 '이미 표시된 것'을 지켜 줄 뿐, **새로 들어온 가짜**는 못 막는다.
        #   그래서 아침에 621건을 손으로 표시해 두었는데 15:32 회차에 새 덤프가 유령 22건을
        #   들여왔고, 업무추출에서 정기점검 UJ2601407 이 1건 → 23건으로 부풀었다.
        #   화면은 멀쩡해 보였다 — 아무도 몰랐다. 한 번 치우는 것으로는 끝나지 않는다.
        try:
            # 다른 폴더에서 이 파일을 모듈로 불러도 옆 파일을 찾도록 제 폴더를 먼저 넣는다.
            # (여기서 조용히 실패하면 보호가 통째로 꺼진다 — 그게 이 사고의 재발 경로다)
            _here = os.path.dirname(os.path.abspath(__file__))
            if _here not in sys.path:
                sys.path.insert(0, _here)
            import clean_contaminated
            for no in clean_contaminated.find(merged):
                merged[no] = {"contaminated": True,
                              "captured_at": (merged.get(no) or {}).get("captured_at") or cap_ms,
                              "why": "iframe 리다이렉트로 피드 본문이 잡힌 가짜 기록(병합 시 자동 판정)"}
        except Exception as e:
            print(f"  · 오염 자동판정 건너뜀({band}): {e}")
        out = {"band_name": d.get("name", band), "posts": merged}
        # ★ 임시파일에 다 쓴 뒤 **한 번에 갈아끼운다** (2026-08-07 실사고).
        #   예전에는 `open(dst,"w")` 로 정본을 **먼저 비우고** 19MB 를 흘려 넣었다.
        #   그 몇 초 동안 파일을 읽는 쪽은 **반쪽짜리 JSON** 을 본다 — 이날 합성검증이
        #   두 번 죽었고, 죽은 자리가 매번 달라서(char 2,581,022 → 9,738,084) 한동안
        #   "캐시가 깨졌다"고 오해했다. 실제로는 쓰는 중이었을 뿐이다.
        #   더 나쁜 경우는 쓰다가 프로세스가 죽는 것이다 — 그때는 **정말로** 깨지고,
        #   8,500 글을 다시 긁어야 한다(밤샘 한 번 분량). os.replace 는 원자적이라
        #   읽는 쪽은 옛 파일이나 새 파일만 보고 그 중간은 못 본다.
        tmp = dst + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(out, fh, ensure_ascii=False)
        os.replace(tmp, dst)
        # 원본 자료 정본은 이름·내용 그대로 둔다. 로컬 처리함의 dump만 raw로 바꿔
        # 다음 실행에서 반복 변환되지 않게 한다.
        try:
            in_cache = os.path.commonpath([os.path.abspath(f), os.path.abspath(CACHE)]) == os.path.abspath(CACHE)
        except ValueError:  # C: 처리함과 Z: 원본처럼 드라이브가 다르면 공통경로가 없다.
            in_cache = False
        if in_cache:
            raw = os.path.join(CACHE, f"raw_{os.path.basename(f)[5:-5]}.json")
            os.replace(f, raw)
        # ★ `.get` 이어야 한다 (2026-08-07 실사고). 지운 글의 묘비 기록에는 본문이 없어
        #   `created_at` 키 자체가 없다. 과거글 구간에는 지운 글이 수백 건씩 섞여 있어서,
        #   밤새 모은 6천여 건이 **전부 캐시에 못 들어가고** "덤프 → 캐시 [FAIL]" 한 줄만
        #   남았다. 수집은 멀쩡히 됐는데 쓰이지 않는, 이 프로젝트가 제일 무서워하는 모양이다.
        dated = sum(1 for p in merged.values() if p.get("created_at"))
        print(f"{d.get('name', band)}: {len(posts)}건 반영 → 캐시 {before}→{len(merged)}건 "
              f"(날짜 있는 글 {dated}건)")

    # ── 모든 덤프를 본 뒤에야 리다이렉트 묘비를 세운다 (분담판 [13]) ──────────────
    # 캐시를 다시 열어 고치는 이유는, 판정에 필요한 '서로 다른 회차'가 위 반복문을
    # 다 돌아야 비로소 모이기 때문이다. 회차 하나로는 판정할 수 없다.
    for band, rounds in redirect_rounds.items():
        ripe = {no: w for no, w in rounds.items()
                if len(w) >= REDIRECT_ROUNDS_FOR_DELETED}
        if not ripe:
            continue
        dst = os.path.join(CACHE, f"{band}.json")
        try:
            doc = json.load(open(dst, encoding="utf-8"))
        except Exception:
            continue
        merged = doc.get("posts") or {}
        n = _mark_redirect_deleted(band, merged, ripe)
        if not n:
            continue
        doc["posts"] = merged
        tmp = dst + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(doc, fh, ensure_ascii=False)
        os.replace(tmp, dst)
        print(f"  · 리다이렉트로 확인된 삭제 글 {n}건 기록({band}) "
              f"— 서로 다른 회차 {REDIRECT_ROUNDS_FOR_DELETED}번 이상. 다시 훑지 않는다")


if __name__ == "__main__":
    main()
