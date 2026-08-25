# -*- coding: utf-8 -*-
"""
convert_dump.py — 브라우저 수집 덤프(dump_*.json) → 대조 캐시(<band>.json) 변환
게시일 파싱 우선순위: 본문 1행 절대시각 → timeText 절대시각 → 상대시각(수집시각 기준).
변환 후 덤프는 raw_*.json 으로 개명 보존.
"""
import sys, os, re, json, time
import hashlib
from datetime import datetime, timedelta

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache")
DEFAULT_CACHE = CACHE
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATE = os.path.join(ROOT, "reports", "밴드덤프_변환상태.json")
LOCK = os.path.join(CACHE, ".convert_dump.lock")
DEFAULT_STATE, DEFAULT_LOCK = STATE, LOCK
STATE_SCHEMA = 1
# 자동 파이프라인이 900초에 자른다. 신규·변경은 먼저 전부 반영하고, 코드가 바뀌어
# 다시 보는 과거 덤프만 이 예산 안에서 끊어 다음 회차가 이어받는다.
REPLAY_BUDGET_SEC = int(os.environ.get("BAND_CONVERT_REPLAY_BUDGET_SEC", "600"))
ABS = re.compile(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일\s*(오전|오후)?\s*(\d{1,2}):(\d{2})")
# ★ 밴드는 시각을 **네 가지 모양**으로 적는다 (2026-08-12 실측). ABS 하나만 보던
#   동안 댓글은 한 건도 캐시에 못 들어왔다 — 수집기가 6건을 멀쩡히 읽어 덤프에
#   담았는데 여기서 전부 버려졌고, 글은 `comments_full=True` 로 닫혀 '확인된 0개'가
#   되어 **다시 뽑히지도 않았다.** 오류는 한 줄도 안 났다.
#   나이별로 이렇게 줄여 적는다: 오늘 `오후 3:57` · 올해 `3월 31일 오전 8:14` ·
#   지난해부터 `2026년 1월 26일`(시각 없음).
MD_TIME = re.compile(r"(\d{1,2})월\s*(\d{1,2})일\s*(오전|오후)?\s*(\d{1,2}):(\d{2})")
ABS_DAY = re.compile(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일")
HM = re.compile(r"(오전|오후)\s*(\d{1,2}):(\d{2})")


def tmp_path(dst):
    """갈아끼우기용 임시 경로 — **pid 를 박아 회차끼리 안 겹치게** 한다.

    ★ 2026-08-11 실측: 5분 파이프라인·09:50 회차·dump_watch 가 같은 밴드를 동시에
      흡수하면 고정 이름 `dst + ".tmp"` 를 서로 가져가 한쪽 os.replace 가
      FileNotFoundError 로 죽었다 — 제 tmp 를 상대가 이미 정본으로 옮긴 뒤다.
      이름에 pid 가 있으면 경쟁 자체가 없다. 크래시가 남긴 옛 pid 의 tmp 는
      한 시간 뒤 다음 회차가 조용히 치운다(내용은 이미 다음 흡수가 다시 만든다 —
      덤프는 매 실행 전부 재처리되므로 잃는 것이 없다).
    """
    import time
    try:
        base = os.path.basename(dst)
        folder = os.path.dirname(dst) or "."
        for n in os.listdir(folder):
            if n.startswith(base + ".") and n.endswith(".tmp"):
                p = os.path.join(folder, n)
                try:
                    if time.time() - os.path.getmtime(p) > 3600:
                        os.remove(p)
                except OSError:
                    pass
    except OSError:
        pass
    return dst + f".{os.getpid()}.tmp"


def swap_in(tmp, dst, tries=6, wait=0.5):
    """`os.replace(tmp, dst)` — 단, 윈도우에서 **읽는 쪽이 물고 있으면 실패한다.**

    ★ 2026-08-08 실사고: 흡수가 `PermissionError [WinError 5]` 로 한 번 죽었다.
      앱 서버가 마침 `band/cache/84789192.json`(5MB)을 읽는 중이었던 것으로 보인다.
      리눅스와 달리 윈도우는 **열려 있는 파일을 갈아끼우지 못한다.**
      그때 남는 것은 5MB `.tmp` 하나와 **옛 캐시** 뿐이다 — 새 글을 다 긁어 놓고도
      캐시는 어제 것이고, 다음 회차는 그 옛 캐시를 보고 "바뀐 것 없음"을 내놓는다.
      **아무 일도 안 일어났는데 안심시키는 결과다.** 조용한 사고의 전형이다.

    읽는 쪽은 곧 파일을 놓으므로 **잠깐 기다렸다 다시** 걸면 대개 풀린다(실측 재시도 1회).
    끝내 안 되면 `.tmp` 를 지우지 않고 **예외를 올린다** — 애써 만든 새 캐시를 버리는
    것보다, 실패를 실패라고 말하고 사람이 그 `.tmp` 를 쓰게 두는 편이 낫다.
    """
    import time
    for i in range(tries):
        try:
            os.replace(tmp, dst)
            if i:
                print(f"  · 캐시 갈아끼우기 {i + 1}번째에 성공({os.path.basename(dst)}) "
                      f"— 누가 읽는 중이었습니다")
            return
        except PermissionError:
            if i == tries - 1:
                raise
            time.sleep(wait * (i + 1))       # 0.5s → 1.0s → … 물러서며 기다린다


def _norm_path(path):
    return os.path.normcase(os.path.abspath(path))


def _runtime_state_path():
    # 합성검증·도구가 CACHE만 임시폴더로 바꾸던 기존 계약을 지킨다. 그때 실제
    # reports 진행표를 섞으면 테스트가 실데이터의 완료 상태를 바꿔 버린다.
    if _norm_path(CACHE) != _norm_path(DEFAULT_CACHE) and STATE == DEFAULT_STATE:
        return os.path.join(CACHE, ".convert_dump.state.json")
    return STATE


def _runtime_lock_path():
    if _norm_path(CACHE) != _norm_path(DEFAULT_CACHE) and LOCK == DEFAULT_LOCK:
        return os.path.join(CACHE, ".convert_dump.lock")
    return LOCK


def _runtime_wasted_path():
    """헛수확 기록도 **CACHE 를 옮기면 같이 따라간다**.

    위 둘과 같은 계약이다([247]) — 검증이 CACHE 만 임시폴더로 바꾸는데 이 파일만
    진짜 reports 를 가리키면, 합성 자료가 실측 증거에 섞인다. 2026-08-19 에
    `t194` 가 조율표를 그렇게 오염시켰고 그때는 40줄 중 38줄이 가짜였다.
    """
    if _norm_path(CACHE) != _norm_path(DEFAULT_CACHE) and WASTED_LOG == DEFAULT_WASTED:
        return os.path.join(CACHE, "밴드_헛수확.json")
    return WASTED_LOG


def converter_version():
    """덤프 해석 규칙의 지문.

    변환기 본문과 오염 판정기가 바뀌면 과거 덤프도 새 규칙으로 다시 봐야 한다.
    수동 버전 숫자는 올리는 것을 잊을 수 있어 실제 두 파일의 바이트를 해싱한다.
    """
    h = hashlib.sha256()
    for path in (__file__, os.path.join(os.path.dirname(__file__), "clean_contaminated.py")):
        try:
            with open(path, "rb") as fh:
                for chunk in iter(lambda: fh.read(1024 * 1024), b""):
                    h.update(chunk)
        except OSError:
            h.update(("<missing>" + os.path.basename(path)).encode("utf-8"))
    return h.hexdigest()[:16]


def _atomic_json(path, doc):
    """JSON을 정본 옆 임시파일에 완성한 뒤 한 번에 갈아끼운다."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = tmp_path(path)
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=1)
        fh.flush()
        try:
            os.fsync(fh.fileno())
        except OSError:
            pass
    swap_in(tmp, path)


def load_state(path=None):
    path = path or _runtime_state_path()
    try:
        with open(path, encoding="utf-8") as fh:
            doc = json.load(fh)
        if not isinstance(doc, dict) or doc.get("schema") != STATE_SCHEMA:
            raise ValueError("state schema")
        if not isinstance(doc.get("files"), dict):
            raise ValueError("state files")
        return doc
    except (OSError, ValueError, json.JSONDecodeError):
        return {"schema": STATE_SCHEMA, "target_version": "", "completed_version": "",
                "files": {}, "updated_at": ""}


def save_state(state, path=None):
    state["schema"] = STATE_SCHEMA
    state["updated_at"] = datetime.now().isoformat(timespec="seconds")
    _atomic_json(path or _runtime_state_path(), state)


def _scan_json_root(root, recursive=True, local_dump_only=False):
    """한 루트의 JSON과 stat을 한 번에 받는다 → (행, 완전스캔 여부)."""
    rows, complete = {}, True
    stack = [root]
    while stack:
        folder = stack.pop()
        try:
            with os.scandir(folder) as it:
                entries = list(it)
        except OSError:
            complete = False
            continue
        for entry in entries:
            try:
                if entry.is_dir(follow_symlinks=False):
                    if recursive:
                        stack.append(entry.path)
                    continue
                if not entry.is_file(follow_symlinks=False):
                    continue
                name = entry.name
                if local_dump_only:
                    if not (name.startswith("dump_") and name.endswith(".json")):
                        continue
                elif not name.endswith(".json"):
                    continue
                st = entry.stat()
            except OSError:
                complete = False
                continue
            key = _norm_path(entry.path)
            rows[key] = {"path": entry.path, "root": _norm_path(root),
                         "size": int(st.st_size),
                         "mtime_ns": int(getattr(st, "st_mtime_ns", int(st.st_mtime * 1e9))),
                         "source_kind": "local_dump" if local_dump_only else "z_dump"}
    return rows, complete


def dump_inventory(state=None, explicit_paths=None):
    """덤프 목록과 **루트별 완전스캔 여부**를 함께 돌려준다.

    Z:가 끊겼는데 빈 목록을 '전부 없어짐'으로 읽으면 파일별 redirect 근거가 통째로
    사라진다. 그래서 못 본 루트는 따로 남기고 그 루트의 상태는 가지치지 않는다.
    """
    files, complete_roots, unavailable_roots = {}, set(), set()
    if explicit_paths is not None:
        roots = {}
        for path in explicit_paths:
            try:
                st = os.stat(path)
            except OSError:
                continue
            root = _norm_path(os.path.dirname(path))
            key = _norm_path(path)
            files[key] = {"path": path, "root": root, "size": int(st.st_size),
                          "mtime_ns": int(getattr(st, "st_mtime_ns", int(st.st_mtime * 1e9))),
                          "source_kind": "explicit"}
            roots[root] = True
        complete_roots.update(roots)
        return {"files": files, "complete_roots": complete_roots,
                "unavailable_roots": unavailable_roots}

    os.makedirs(CACHE, exist_ok=True)
    local, ok = _scan_json_root(CACHE, recursive=False, local_dump_only=True)
    files.update(local)
    (complete_roots if ok else unavailable_roots).add(_norm_path(CACHE))

    configured = []
    try:
        sys.path.insert(0, ROOT)
        import source_dirs
        canonical = [os.path.join(source_dirs.BAND_DIR, "수집본"),
                     os.path.join(source_dirs.BAND_DIR, "브라우저덤프")]
        reported = list(source_dirs.band_dump_dirs())
        # 기존 호출자와 합성검증은 band_dump_dirs를 임시 원본으로 바꾼다. 정본 아래
        # 경로라면 빠진 갈래까지 확인하고, 정본 밖의 명시적 대체면 그 목록만 따른다.
        canonical_keys = {_norm_path(p) for p in canonical}
        configured = (reported if reported and
                      any(_norm_path(p) not in canonical_keys for p in reported)
                      else canonical)
        band_parent_visible = bool(reported) or os.path.isdir(source_dirs.BAND_DIR)
    except Exception:
        band_parent_visible = False
    for root in configured:
        key = _norm_path(root)
        if os.path.isdir(root):
            rows, ok = _scan_json_root(root, recursive=True)
            files.update(rows)
            (complete_roots if ok else unavailable_roots).add(key)
        elif band_parent_visible:
            # 상위 밴드 폴더가 보이는데 이 갈래만 없으면 '빈 루트'를 끝까지 본 것이다.
            complete_roots.add(key)
        else:
            unavailable_roots.add(key)

    # 로컬 dump는 처리 뒤 raw_로 이름을 바꾼다. 그 파일에서 얻은 redirect 근거가
    # 다음 회차에 사라지지 않도록 **이 상태가 이미 아는 raw 별칭만** 다시 목록에 넣는다.
    for key, entry in ((state or {}).get("files") or {}).items():
        if entry.get("source_kind") != "local_raw" or key in files:
            continue
        path = entry.get("path") or key
        try:
            st = os.stat(path)
        except OSError:
            continue
        files[key] = {"path": path, "root": _norm_path(CACHE), "size": int(st.st_size),
                      "mtime_ns": int(getattr(st, "st_mtime_ns", int(st.st_mtime * 1e9))),
                      "source_kind": "local_raw"}
    return {"files": files, "complete_roots": complete_roots,
            "unavailable_roots": unavailable_roots}


def _same_fingerprint(entry, row):
    return (int(entry.get("size") or -1) == int(row.get("size") or -2)
            and int(entry.get("mtime_ns") or -1) == int(row.get("mtime_ns") or -2))


def _entry_current(entry, row, version):
    return bool(entry and _same_fingerprint(entry, row)
                and entry.get("converter_version") == version)


def _prune_state_files(state, inventory):
    """끝까지 본 루트에서 실제로 사라진 파일만 상태에서도 뺀다."""
    current = inventory["files"]
    complete = set(inventory["complete_roots"])
    for key, entry in list((state.get("files") or {}).items()):
        if key in current:
            continue
        if entry.get("source_kind") == "local_raw" and os.path.isfile(entry.get("path") or key):
            continue
        if entry.get("root") in complete:
            state["files"].pop(key, None)


def _state_entry(row, version, status, band="", captured_at=0, redirect_hits=()):
    return {"path": row["path"], "root": row["root"], "size": int(row["size"]),
            "mtime_ns": int(row["mtime_ns"]), "source_kind": row.get("source_kind") or "",
            "converter_version": version, "status": status, "band": str(band or ""),
            "captured_at": int(captured_at or 0),
            "redirect_hits": sorted({str(n) for n in (redirect_hits or ())},
                                    key=lambda x: (len(x), x))}


def redirect_rounds_from_state(state, inventory, version):
    """현재 실제 파일·현재 코드에 해당하는 redirect 근거만 다시 조립한다."""
    rounds = {}
    for key, row in inventory["files"].items():
        entry = (state.get("files") or {}).get(key)
        if not _entry_current(entry, row, version) or entry.get("status") != "merged":
            continue
        band, cap = str(entry.get("band") or ""), int(entry.get("captured_at") or 0)
        if not band:
            continue
        for no in entry.get("redirect_hits") or []:
            rounds.setdefault(band, {}).setdefault(int(no), set()).add(cap)
    return rounds


def _lock_acquire(wait_sec=None):
    """캐시 read-modify-write와 상태 체크포인트를 한 변환기만 수행하게 한다."""
    wait_sec = float(os.environ.get("BAND_CONVERT_LOCK_WAIT_SEC", "120")
                     if wait_sec is None else wait_sec)
    lock_path = _runtime_lock_path()
    os.makedirs(os.path.dirname(lock_path), exist_ok=True)
    end = time.monotonic() + max(0.0, wait_sec)
    while True:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            dead = False
            try:
                import pid_alive
                with open(lock_path, encoding="utf-8", errors="replace") as fh:
                    words = fh.read().split()
                pid, fingerprint, born = pid_alive.owner_from_words(words)
                dead = pid_alive.owner_alive(pid, pid_started_at=fingerprint,
                                             born_before=born) is False
            except (OSError, ValueError):
                dead = False
            if dead:
                try:
                    os.unlink(lock_path)
                    continue
                except OSError:
                    pass
            if time.monotonic() >= end:
                return False
            time.sleep(0.25)
            continue
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            import pid_alive
            fh.write(f"{os.getpid()} {pid_alive.stamp()} {datetime.now():%Y-%m-%dT%H:%M:%S}\n")
        return True


def _lock_release():
    try:
        os.unlink(_runtime_lock_path())
    except OSError:
        pass


def _hour24(ap, h):
    return int(h) % 12 + (12 if ap == "오후" else 0)


def parse_dt(text, captured_ms):
    t = text or ""
    m = ABS.search(t)
    if m:
        y, mo, d, ap, h, mi = m.groups()
        return datetime(int(y), int(mo), int(d), _hour24(ap, h), int(mi))
    base = datetime.fromtimestamp((captured_ms or 0) / 1000) if captured_ms else datetime.now()

    # 연도가 빠진 모양. 밴드는 미래를 보여 주지 않으므로, 수확 시각보다 뒤로 나오면
    # 그것은 작년 것이다 — 연도를 짐작하는 것이 아니라 **불가능한 쪽을 지우는** 것이다.
    m = MD_TIME.search(t)
    if m:
        mo, d, ap, h, mi = m.groups()
        try:
            dt = datetime(base.year, int(mo), int(d), _hour24(ap, h), int(mi))
        except ValueError:
            dt = None
        if dt:
            if dt > base + timedelta(days=1):
                try:
                    dt = dt.replace(year=base.year - 1)
                except ValueError:
                    return None
            return dt

    # 오늘 것은 시각만 적는다.
    if "어제" not in t:
        m = HM.search(t)
        if m:
            ap, h, mi = m.groups()
            return base.replace(hour=_hour24(ap, h), minute=int(mi),
                                second=0, microsecond=0)

    m = re.search(r"(\d+)분 전", t)
    if m: return base - timedelta(minutes=int(m.group(1)))
    m = re.search(r"(\d+)시간 전", t)
    if m: return base - timedelta(hours=int(m.group(1)))
    if "어제" in t:
        y = base - timedelta(days=1)
        m = HM.search(t)
        if m:
            ap, h, mi = m.groups()
            return y.replace(hour=_hour24(ap, h), minute=int(mi),
                             second=0, microsecond=0)
        return y
    m = re.search(r"(\d+)일 전", t)
    if m: return base - timedelta(days=int(m.group(1)))

    # 날짜만 있고 시각이 없는 모양. **하루의 끝으로 놓는다.**
    # 아는 것은 날짜뿐이고 시각은 모른다 — 그런데 이 값이 쓰이는 자리는
    # '댓글이 글보다 나중인가'라는 순서 하나다([155]). 하루의 끝은 그 순서를
    # 절대 깨지 않는 유일한 지점이다(00:00 으로 놓으면 같은 날 올라온 글보다
    # 댓글이 앞서서, 실제로 달린 취소 댓글이 순서에서 밀려 안 보이게 된다).
    m = ABS_DAY.search(t)
    if m:
        y, mo, d = m.groups()
        try:
            return datetime(int(y), int(mo), int(d), 23, 59, 59)
        except ValueError:
            return None
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


def looks_like_cache(d):
    """이 JSON 이 **덤프가 아니라 캐시 사본**인가 (2026-08-19 · 분담판 [154]).

    Z: 수집본 폴더에 캐시 통사본이 날짜별로 쌓여 있어 흡수기가 **제 출력을 입력으로
    다시 먹고 있었다**(실측 272개 중 132개). 그 안의 옛 값이 새 수확을 덮어, 그날 정상
    수확한 글의 `created_at` 이 흡수 뒤 다시 사라졌다.

    가르는 근거는 **스키마 하나**다 — 캐시는 `band_name` 을 갖고 `band` 가 없다.
    실측 오탐 0 · 미탐 0(진짜 덤프 133 · 캐시 사본 132 · API 수집본 6은 둘 다 가짐).
    파일명으로 가르면 `raw_api_90610953.json` 같은 진짜 수집본까지 걸린다 —
    **못 읽는 것보다 잘못 거르는 것이 나쁘다**(`[172]`).
    """
    return isinstance(d, dict) and "band_name" in d and not d.get("band")


# ── 이 앱이 다루는 밴드는 둘뿐이다 (2026-08-22 형님 지시) ──────────────────
#   "밴드는 앞으로 이 두개만 긁어와 이 앱은 이 두개만 필요해"
#     · 84789192 = (주)유니버셜리프트 쿠팡AS
#     · 90610953 = 매출처업무
#   ★ **모양 검사보다 훨씬 센 근거다.** 예전에는 '7~10자리 숫자'라는 생김새로만
#     걸렀는데, 그 문은 날짜 꼬리표(`260807`)·시각 도장(`202608082047`)·해시 앞토막
#     (`8518730`)을 전부 통과시켰다. 그렇게 생긴 유령 캐시는 다음 실행에서 '아는
#     번호'가 되어 스스로를 되살렸고(2026-08-12·2026-08-19 실사고) 사람이 파일을
#     치워야만 나았다. 목록이 확정된 지금은 **근원에서 막는 것**이 옳다.
#   ★ 밴드가 늘면 **이 한 줄만** 고친다. 코드는 안 고친다.
BAND_IDS = ("84789192", "90610953")


def plausible_band(n):
    """이 앱이 긁는 밴드인가.

    ★ **판정은 한 곳이다.** `known_bands`·`band_from_name`·`collect_queue`·
      `collect_sources`·`recheck_plan` 이 각자 재면 언젠가 갈리고, 갈린 뒤에는
      어느 쪽이 맞는지 아무도 모른다. 실측 2026-08-22 로 `recheck_plan` 만
      `isdigit()` 로 따로 재고 있었다 — 그래서 그 회차만 유령을 계속 받았다.
    """
    return n in BAND_IDS


def known_bands():
    """캐시에 이미 있는 밴드번호 — 파일명이 애매할 때의 가장 좋은 근거다.

    ★ **유령은 '아는 번호'가 아니다** (2026-08-12 실사고 · 분담판 [45]).
      전에는 캐시에 있는 숫자 파일명을 그대로 다 담았다. 그런데 유령 캐시가 한 번
      생기면 그 번호가 여기 들어와 `band_from_name` 의 규칙① 에 걸리고, ②의 길이
      방어를 **통째로 건너뛴다** — 유령이 스스로를 영구히 되살리는 구조였다.
      실측: `band/cache/202608082047.json`(12자리 시각 도장, 글 0개)이 살아 있어서
      `dump_202608082047_84789192.json` 이 84789192 가 아니라 유령으로 갔고,
      그 파일이 그날 08:08 에 또 새로 써졌다. 유령 파일을 사람이 치워야만 나았다.
      이제 근거로 삼기 전에 **생김새부터 본다** — 캐시에 있다는 것만으로는 모자란다.
    """
    try:
        return {f[:-5] for f in os.listdir(CACHE)
                if f.endswith(".json") and plausible_band(f[:-5])}
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
    plausible = [n for n in nums if plausible_band(n)]
    if not plausible:
        return None          # 모르면 **모른다고 한다** — 없는 밴드를 만드는 것보다 낫다
    longest = max(len(n) for n in plausible)
    return [n for n in plausible if len(n) == longest][-1]


def _ui_junk(txt):
    """밴드 화면 UI 가 통째로 본문으로 딸려 온 수확인가.

    2026-08-20 실측 21건(90610953 14 · 84789192 7). 앞머리가 '글 옵션' 이나
    'N명이 읽었습니다' 이고 끝이 '채팅' 으로 닫힌다 — 밴드 글 본문에는 없는 말이다.
    ★ 이것이 위험한 이유는 길이다. UI 글자가 더해져 **정상 본문보다 길어지므로**
      아래 병합의 '길면 이긴다' 갈래를 그냥 통과해 멀쩡한 본문을 덮는다.
      그래서 다시 긁어도 안 고쳐진다 — 오염이 이기는 구조였다([169]).
    문은 좁게 둔다([172]) — 잘못 지목하면 멀쩡한 본문이 안 들어온다.
    실측 8,605글 중 21건(0.24%)만 걸린다.
    """
    if not txt:
        return False
    head = txt[:200]
    return ('글 옵션' in head) or ('명이 읽었습니다' in head) or txt.rstrip().endswith('채팅')


def dump_files():
    """로컬 처리함과 0. 원본 자료의 밴드 JSON 정본을 함께 읽는다."""
    inventory = dump_inventory(load_state())
    # raw_는 파일별 redirect 근거를 보존하려고 내부 inventory에만 되살린다.
    # 이 공개 함수는 예전 계약대로 실제 입력 덤프만 돌려준다.
    return sorted(row["path"] for row in inventory["files"].values()
                  if row.get("source_kind") != "local_raw")


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
# ★ 다시 긁었는데 **받자마자 오염으로 되돌아간** 번호 (2026-08-25 · 분담판 [225]).
#   그 자리에서 세지 않으면 "3건 반영" 한 줄만 남아 성공처럼 읽힌다([169]).
WASTED_LOG = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                          "reports", "밴드_헛수확.json")
DEFAULT_WASTED = WASTED_LOG


def _twin_index(merged):
    """살아남은 글의 본문 지문 → 번호. 되돌아간 것이 **누구를 베껴 왔는지** 대기 위해서다."""
    idx = {}
    for k, v in merged.items():
        if not isinstance(v, dict) or v.get("contaminated") or v.get("deleted"):
            continue
        sig = (v.get("content") or "")[:120]
        if sig and sig not in idx:
            idx[sig] = k
    return idx


def _note_wasted(band, dup, twins, cap_ms):
    """헛수확을 회차 시각과 함께 쌓아 둔다 → 이 회차에 새로 쌓인 수.

    왜 세는가 (2026-08-25 실사고 · 분담판 [225])
      오염 번호를 다시 긁으면 밴드가 **이웃 글의 본문**을 그대로 돌려준다. 그러면
      `clean_contaminated.find` 이 그것을 정확히 가짜로 잡아 도로 오염으로 표시한다 —
      **그 판정은 맞다.** 잘못은 그다음이다: 아무도 그 사실을 안 적어서 다음 회차가
      같은 번호를 또 뽑고, 브라우저는 번호 하나에 20초를 또 쓴다.
      실측 2026-08-25: 269건을 긁어 **되살아난 것 0건**(84789192 3건은 받자마자
      되돌아갔고 90610953 237건은 시각이 없어 버려졌다).

    ★ 여기서 아무것도 거르지 않는다([172]) — 세어서 말할 뿐이다. 이것을 근거로
      묘비를 세우면 실재하는 글이 유령이 된다(되돌릴 수 없는 쪽 · [217]·[334]).
    ★ 못 적어도 흡수는 그대로 간다 — 기록 하나로 회차를 세우지 않는다.
    """
    if not dup:
        return 0
    try:
        path = _runtime_wasted_path()
        doc = {}
        if os.path.exists(path):
            with open(path, encoding="utf-8") as fh:
                doc = json.load(fh) or {}
        book = doc.setdefault(str(band), {})
        n = 0
        for no in dup:
            row = book.setdefault(str(no), {"회차": [], "베낀번호": ""})
            if cap_ms and cap_ms not in row["회차"]:
                row["회차"].append(cap_ms)
                row["회차"] = sorted(row["회차"])[-10:]
                n += 1
            if twins.get(no):
                row["베낀번호"] = twins[no]
        doc["갱신"] = cap_ms
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(doc, fh, ensure_ascii=False)
        os.replace(tmp, path)
        return n
    except Exception as e:
        print(f"  · 헛수확 기록 실패({band}): {e}")
        return 0


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
        with open(PROBE_LOG, encoding="utf-8") as fh:
            doc = json.load(fh)
    except Exception:
        doc = {}
    prev = doc.get(str(band)) or {}
    if str(prev.get("확인시각") or "") > when:
        return                        # 더 최근 확인이 이미 있으면 옛 회차로 덮지 않는다
    doc[str(band)] = {"이름": name, "확인시각": when, "수집최대": top,
                      "없음확인": absent[0], "연속없음": len(absent)}
    os.makedirs(os.path.dirname(PROBE_LOG), exist_ok=True)
    tmp = tmp_path(PROBE_LOG)                     # pid 별 이름 — 회차 경쟁 방지
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
        with open(CHANGED_LOG, encoding="utf-8") as fh:
            old = json.load(fh)
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
        with open(CHANGED_LOG, "w", encoding="utf-8") as fh:
            json.dump(doc, fh, ensure_ascii=False, indent=1)
    except Exception:
        pass
    print(f"  ★ 수정된 글 {len(nos)}건 감지({band}) → reports/밴드_수정글.json")


def _convert_files(files=None, checkpoint=None, apply_redirect=True):
    """주어진 덤프를 예전 변환 규칙 그대로 병합한다.

    checkpoint는 **캐시 교체와 로컬 raw 개명까지 끝난 뒤** 호출된다. 상태파일이
    캐시보다 먼저 앞서가는 일을 막기 위해 증분 실행기가 이 경계를 사용한다.
    """
    # ★ 리다이렉트 실패는 **회차를 가로질러** 세야 뜻이 생긴다 (분담판 [13]).
    #   한 덤프만 보면 "이번에 실패했다"밖에 모른다. 덤프는 매 실행 전부 재처리되므로
    #   여기 담아 두면 이 한 번의 실행이 곧 '여러 회차를 본 것'이 된다.
    #   {밴드: {번호: {회차 캡처시각, ...}}}
    redirect_rounds = {}
    skipped_cache = skipped_noband = 0
    for f in (dump_files() if files is None else files):
        with open(f, encoding="utf-8") as fh:
            d = json.load(fh)
        if not isinstance(d, dict) or not isinstance(d.get("posts"), (dict, list)):
            if checkpoint:
                checkpoint(f, f, "ignored_schema", "", 0, ())
            continue
        # ★ **캐시 사본은 덤프가 아니다** (2026-08-19 실사고 · 분담판 [154]).
        #   Z: 수집본 폴더에 캐시 통사본이 날짜별로 쌓여 있어(8/14~8/19 매일 두 개)
        #   흡수기가 **제 출력을 입력으로 다시 먹고 있었다.** 실측 덤프 목록 272개 중
        #   **132개가 캐시 사본**(약 1.6GB)이었고, 그 안의 옛 값이 새 수확을 덮어
        #   그날 정상 수확한 90610953/5442 의 `created_at` 이 흡수 뒤 다시 None 이 됐다.
        #   가르는 근거는 **스키마**다 — 캐시는 `band_name` 을 갖고 `band` 가 없다.
        #   실측 오탐 0 · 미탐 0(진짜 덤프 133 · 캐시 사본 132 · API 수집본 6은 둘 다 가짐).
        #   파일명(밴드번호와 같음)으로 가르면 `raw_api_90610953.json` 같은 진짜
        #   수집본까지 걸린다 — **문은 좁은 쪽으로 잡는다**(`[172]`).
        if looks_like_cache(d):
            skipped_cache += 1
            if checkpoint:
                checkpoint(f, f, "ignored_cache_copy", "", int(d.get("capturedAt") or 0), ())
            continue
        # ★ 밴드를 못 읽으면 **해시 이름 캐시를 만들지 않는다** (같은 사고의 곁가지).
        #   예전에는 sha256 앞 10자리로 캐시를 만들었는데, 그 파일이 다음 실행에서
        #   또 덤프로 읽혀 `band_from_name('8518730cc9') -> 8518730`(7자리) 으로
        #   **유령 밴드를 낳았다**(자기 증식). 2026-08-12 의 `202608082047` 사고와
        #   같은 종류인데 7자리라 길이 검사를 통과한다. 못 읽으면 **안 만들고 센다**(`[169]`).
        band = str(d.get("band") or band_from_name(os.path.basename(f)) or "")
        if not band:
            skipped_noband += 1
            if checkpoint:
                checkpoint(f, f, "ignored_no_band", "", int(d.get("capturedAt") or 0), ())
            continue
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
                with open(dst, encoding="utf-8") as fh:
                    old = json.load(fh)
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
            # ★ 본문이 안 바뀐 글은 아래에서 merged[no]=rec 를 **안 한다**. 그러면 위에서
            #   합친 댓글·완독표시가 rec 에만 있고 버려져, 재수집분이 옛 '미확인'을 영영
            #   못 덮는다 — 백필이 같은 95건을 매 회차 다시 뽑으며 수렴하지 못한다
            #   (2026-08-11 실측: 열린 원장 95건이 두 번 재수집·흡수 뒤에도 안 줄었다).
            #   댓글은 '쌓이는' 것이라 본문 교체 여부와 무관하게 늘 cur 에 최신으로 남긴다.
            cur["comments"] = rec["comments"]
            cur["comments_full"] = rec["comments_full"]
            new_txt, old_txt = rec["content"] or "", cur.get("content") or ""
            # '…더보기'로 잘린 피드 수집분이 상세 전문을 덮어쓰지 않게 한다.
            # ★ 앞 조건만으로는 **영영 참이 안 됐다** (2026-08-20 실측 · [167]).
            #   잘린 수확은 끝에 '...더보기' 를 **달고** 오는데 그 표시까지 포함해 prefix 를
            #   대므로 old_txt 는 절대 그것으로 시작하지 않는다. 그래서 90610953/5233 이
            #   4,288자 → 136자로, 5442 가 214자로 덮였다(둘 다 글쓴이도 빈칸이 됐다).
            #   가드가 있는데 한 번도 안 걸린 자리다([169] — 계기가 0 을 내면 아무도 의심 안 한다).
            #   '더보기' 는 밴드가 스스로 '이 글은 접혀 있다'고 적어 준 표시이므로 그 자체가 근거다.
            #   문은 좁다 — **옛 본문보다 짧을 때만** 이라 최악의 결과가 '긴 본문을 지킨다' 이다.
            collapsed = new_txt.rstrip().endswith('더보기') and len(new_txt) < len(old_txt)
            truncated = (len(new_txt) < len(old_txt) * 0.9
                         and old_txt.startswith(new_txt[:200])) or collapsed
            newer = rec["captured_at"] >= int(cur.get("captured_at") or 0)
            # ★ **아는 시각을 모르는 것으로 되돌리지 않는다** (2026-08-19 · [154]).
            #   아래 갈래들은 본문이 길거나 새로우면 `rec` 로 통째로 갈아 끼운다.
            #   그때 `rec` 에 `created_at` 이 없으면 이미 알던 시각을 **잃는다** —
            #   시각 없는 글은 `datalake.band_day()` 가 빈 값을 줘 어떤 기간 질문에도
            #   안 걸린다(`[152]` 와 같은 조용한 사고). 본문·댓글은 새것이 이기되
            #   시각만은 보존한다.
            if not rec.get("created_at") and cur.get("created_at"):
                rec["created_at"] = cur["created_at"]
            # ★ 글쓴이도 같은 규칙이다 — 아는 값을 모르는 것으로 되돌리지 않는다([334]).
            #   밴드 글에 글쓴이가 없는 일은 없다. 빈 글쓴이는 "지워졌다"가 아니라
            #   **"이번 수확이 못 읽었다"** 는 뜻이고, 그것으로 덮으면 누가 올린 글인지를 잃는다.
            if not (rec.get("author") or "").strip() and (cur.get("author") or "").strip():
                rec["author"] = cur["author"]
            # ★ 화면 UI 가 딸려 온 수확은 깨끗한 본문을 못 덮는다(2026-08-20 · _ui_junk).
            #   그 반대(오염을 깨끗한 것으로 고치는 것)는 **길이와 무관하게** 받는다 —
            #   오염본이 더 길기 때문에 길이 규칙만 두면 영영 안 고쳐진다.
            junk_new, junk_old = _ui_junk(new_txt), _ui_junk(old_txt)
            if junk_new and not junk_old and old_txt.strip():
                pass                                  # cur 을 그대로 둔다
            elif junk_old and not junk_new and new_txt.strip():
                rec["updated_at"] = cap_ms
                rec["prev_len"] = len(old_txt)
                rec["ui_junk_fixed"] = True
                merged[no] = rec
                changed.append(no)
            elif rec.get("created_at") and not cur.get("created_at"):
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
        file_redirect_hits = set()
        try:
            ok_here = sum(1 for p in posts.values() if p.get("created_at"))
            for no in _redirect_hits(d.get("notime") or {}, ok_here):
                redirect_rounds.setdefault(band, {}).setdefault(no, set()).add(cap_ms)
                file_redirect_hits.add(no)
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
        dup_now = {}
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
            _found = clean_contaminated.find(merged)
            # ★ 방금 받아 온 것이 **그 자리에서** 되돌아갔는지 먼저 붙잡는다([169]).
            #   표시한 뒤에 물으면 옛 오염과 구별할 수 없다 — 그러면 "3건 반영" 한 줄만
            #   남아 성공처럼 읽힌다(분담판 [225] · 2026-08-25).
            _twins = _twin_index(merged)
            dup_now = {no: _twins.get(_found[no], "") for no in _found if no in posts}
            for no in _found:
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
        tmp = tmp_path(dst)                       # pid 별 이름 — 회차 경쟁 방지
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(out, fh, ensure_ascii=False)
        swap_in(tmp, dst)
        # 원본 자료 정본은 이름·내용 그대로 둔다. 로컬 처리함의 dump만 raw로 바꿔
        # 다음 실행에서 반복 변환되지 않게 한다.
        try:
            in_cache = os.path.commonpath([os.path.abspath(f), os.path.abspath(CACHE)]) == os.path.abspath(CACHE)
        except ValueError:  # C: 처리함과 Z: 원본처럼 드라이브가 다르면 공통경로가 없다.
            in_cache = False
        final_path = f
        if in_cache and os.path.basename(f).startswith("dump_"):
            raw = os.path.join(CACHE, f"raw_{os.path.basename(f)[5:-5]}.json")
            os.replace(f, raw)
            final_path = raw
        if checkpoint:
            checkpoint(f, final_path, "merged", band, cap_ms, file_redirect_hits)
        # ★ `.get` 이어야 한다 (2026-08-07 실사고). 지운 글의 묘비 기록에는 본문이 없어
        #   `created_at` 키 자체가 없다. 과거글 구간에는 지운 글이 수백 건씩 섞여 있어서,
        #   밤새 모은 6천여 건이 **전부 캐시에 못 들어가고** "덤프 → 캐시 [FAIL]" 한 줄만
        #   남았다. 수집은 멀쩡히 됐는데 쓰이지 않는, 이 프로젝트가 제일 무서워하는 모양이다.
        dated = sum(1 for p in merged.values() if p.get("created_at"))
        print(f"{d.get('name', band)}: {len(posts)}건 반영 → 캐시 {before}→{len(merged)}건 "
              f"(날짜 있는 글 {dated}건)")
        # ★ 받자마자 되돌아간 것을 **숫자로 말한다**(분담판 [225] · 2026-08-25).
        #   안 적으면 "N건 반영" 이 성공으로 읽힌다 — 이 저장소가 되풀이해 당한 모양이다([169]).
        if dup_now:
            _n = _note_wasted(band, list(dup_now), dup_now, cap_ms)
            _ex = ", ".join((f"{k}(={v}와 같은 글)" if v else k)
                            for k, v in list(dup_now.items())[:4])
            print(f"  · 그중 {len(dup_now)}건은 **받자마자 오염으로 되돌아갔다** "
                  f"— 다른 번호와 같은 글이다({_ex}). "
                  f"다시 긁어도 같다 → reports/밴드_헛수확.json (이번 회차 새로 {_n}건)")

    # ── 모든 덤프를 본 뒤에야 리다이렉트 묘비를 세운다 (분담판 [13]) ──────────────
    # 캐시를 다시 열어 고치는 이유는, 판정에 필요한 '서로 다른 회차'가 위 반복문을
    # 다 돌아야 비로소 모이기 때문이다. 회차 하나로는 판정할 수 없다.
    if apply_redirect:
        _apply_redirect_rounds(redirect_rounds)

    # ★ **뺀 것은 조용히 빼지 않는다**(`[169]`) — 0건이 '다 봤다'로 읽히면 안 된다.
    if skipped_cache or skipped_noband:
        print(f"  · 덤프가 아닌 파일 건너뜀 — 캐시 사본 {skipped_cache}개"
              f" · 밴드를 못 읽은 파일 {skipped_noband}개 (분담판 [154])")
    return redirect_rounds


def _apply_redirect_rounds(redirect_rounds):
    """완전 재생이 확인된 때에만 redirect 묘비를 캐시에 적용한다."""
    for band, rounds in redirect_rounds.items():
        ripe = {no: w for no, w in rounds.items()
                if len(w) >= REDIRECT_ROUNDS_FOR_DELETED}
        if not ripe:
            continue
        dst = os.path.join(CACHE, f"{band}.json")
        try:
            with open(dst, encoding="utf-8") as fh:
                doc = json.load(fh)
        except Exception:
            continue
        merged = doc.get("posts") or {}
        n = _mark_redirect_deleted(band, merged, ripe)
        if not n:
            continue
        doc["posts"] = merged
        tmp = tmp_path(dst)                       # pid 별 이름 — 회차 경쟁 방지
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(doc, fh, ensure_ascii=False)
        swap_in(tmp, dst)
        print(f"  · 리다이렉트로 확인된 삭제 글 {n}건 기록({band}) "
              f"— 서로 다른 회차 {REDIRECT_ROUNDS_FOR_DELETED}번 이상. 다시 훑지 않는다")


def main(explicit_paths=None, budget_sec=None, lock_wait_sec=None):
    """신규·변경 우선 증분 변환기.

    첫 도입 또는 변환 규칙 변경 때의 과거 파일은 시간예산만큼 처리하고 파일마다
    이어받기 상태를 남긴다. 신규·변경 파일은 그 대기열보다 먼저 전부 처리한다.
    """
    if not _lock_acquire(lock_wait_sec):
        print("  · 밴드 덤프 변환기가 이미 실행 중입니다 — 이번 호출은 겹쳐 쓰지 않습니다")
        return 0
    try:
        version = converter_version()
        state = load_state()
        inventory = dump_inventory(state, explicit_paths=explicit_paths)
        _prune_state_files(state, inventory)
        state["target_version"] = version

        # 첫 도입에서 458개를 '신규'로 오해해 시간예산을 무시하지 않도록, 현재
        # 목록을 pending으로 먼저 원자 저장한다. 캐시 완료 표시는 아직 하나도 아니다.
        bootstrap = not state.get("files") and not state.get("completed_version")
        if bootstrap:
            for key, row in inventory["files"].items():
                state["files"][key] = _state_entry(row, "", "pending")
            state["bootstrap_pending"] = True
            save_state(state)

        fresh, replay = [], []
        for key, row in inventory["files"].items():
            entry = state["files"].get(key)
            if _entry_current(entry, row, version):
                continue
            # 지문이 없거나 바뀐 파일은 방금 들어온 자료다. 같은 파일인데 변환기
            # 버전만 옛것이면 재생 대기열로 보내 시간예산 안에서 이어간다.
            target = fresh if not entry or not _same_fingerprint(entry, row) else replay
            target.append((key, row))
        newest_first = lambda item: (-int(item[1].get("mtime_ns") or 0), item[0])
        fresh.sort(key=newest_first)
        replay.sort(key=newest_first)

        current = {"key": None, "row": None}

        def checkpoint(old_path, final_path, status, band, captured_at, redirect_hits):
            old_key = _norm_path(old_path)
            row = inventory["files"].get(old_key) or current["row"]
            if row is None:
                raise RuntimeError(f"체크포인트 입력을 찾을 수 없습니다: {old_path}")
            final_key = _norm_path(final_path)
            final_row = dict(row)
            if final_key != old_key:
                st = os.stat(final_path)
                final_row.update({"path": final_path, "root": _norm_path(CACHE),
                                  "size": int(st.st_size),
                                  "mtime_ns": int(getattr(st, "st_mtime_ns",
                                                          int(st.st_mtime * 1e9))),
                                  "source_kind": "local_raw"})
                inventory["files"].pop(old_key, None)
                state["files"].pop(old_key, None)
                inventory["files"][final_key] = final_row
            state["files"][final_key] = _state_entry(
                final_row, version, status, band, captured_at, redirect_hits)
            # merged 상태는 캐시 swap 성공 뒤에만 여기 온다. 여기서 상태까지 swap하면
            # 중간에 죽어도 다음 회차는 정확히 이 다음 파일부터 이어간다.
            save_state(state)

        processed_fresh = processed_replay = 0
        for key, row in fresh:
            current.update(key=key, row=row)
            _convert_files([row["path"]], checkpoint=checkpoint, apply_redirect=False)
            processed_fresh += 1

        replay_started = time.monotonic()
        replay_budget = REPLAY_BUDGET_SEC if budget_sec is None else max(0.0, float(budget_sec))
        for key, row in replay:
            if time.monotonic() - replay_started >= replay_budget:
                break
            current.update(key=key, row=row)
            _convert_files([row["path"]], checkpoint=checkpoint, apply_redirect=False)
            processed_replay += 1

        all_current = all(
            _entry_current(state["files"].get(key), row, version)
            for key, row in inventory["files"].items())
        full_replay = all_current and not inventory["unavailable_roots"]
        if full_replay:
            # redirect는 파일별 근거를 상태에서 완전히 다시 조립한다. Z:를 못 본 회차나
            # 구버전 백로그가 남은 회차는 묘비를 절대 세우지 않는다.
            _apply_redirect_rounds(redirect_rounds_from_state(state, inventory, version))
            state["completed_version"] = version
            state["bootstrap_pending"] = False
        else:
            state["bootstrap_pending"] = True
            why = []
            if not all_current:
                why.append("과거 덤프 재생 대기")
            if inventory["unavailable_roots"]:
                why.append("원본 루트 확인 불가")
            print("  · redirect 삭제 판정 보류 — " + " · ".join(why))
        save_state(state)
        remaining = sum(
            not _entry_current(state["files"].get(key), row, version)
            for key, row in inventory["files"].items())
        print(f"  · 증분 변환: 신규·변경 {processed_fresh}개 · 과거 재생 {processed_replay}개"
              f" · 남음 {remaining}개")
        return 0
    finally:
        _lock_release()


if __name__ == "__main__":
    # ★ 수집 문 — 남의 차선 일을 하는 **사람 창**만 막는다(2026-08-22 형님 지시).
    #   무인 회차(워치독·09:50·증분)는 그대로 통과한다 — 막으면 자동 수집이
    #   통째로 멈추면서 회차는 '성공'으로 적힌다([169]).
    import os as _os, sys as _sys
    _sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
    import collect_gate as _gate
    _gate.guard("밴드 덤프 흡수(캐시 갱신)")
    main()
