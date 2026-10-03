# -*- coding: utf-8 -*-
"""표준 업무 절차서를 공유 폴더에 **번호 순서대로** 게시한다 (2026-10-02 형님 지시).

형님 지시: "여기도 번호 매겨서 모든 절차서 관리해 규칙대로" → "자동으로 해둬".

자리: `source_dirs.SOP_DIR` (`2. CSOS DATA\\11. 표준 업무 절차서`)
  1. 돌발AS 업무 절차서        ← 워드·PPT (묶음 하나에 번호 폴더 하나)
  2. 정기점검 업무 절차서
  3. 유니웍스 넘김 자료(JSON)  ← 모든 묶음의 JSON
  4. 아직 못 채운 칸           ← 사람이 채울 목록(여기서 안 만든다)
  old                          ← 바로 앞 판들
  새 묶음이 생기면 **5. 부터** 이어 붙인다(1~4 는 이미 뜻이 정해진 자리다).

★ **바뀌었을 때만 게시한다.** 같은 내용을 날짜만 바꿔 넣으면 어느 것이 최신인지
  헷갈리기만 한다(2026-10-01 실측: 9/30 판과 내용이 같았다). 내용 지문은
  `스키마판 + 절차서 칸` 으로 잰다 — `만든때` 는 뺀다(그것만 매번 바뀐다).
★ **바뀐 게 없으면 Z: 를 한 번도 안 만진다**([168]) — 지문은 로컬 표지
  `reports/.절차서_게시.json` 과 대 본다. 워치독 30분마다 불려도 값이 0 이다.
★ **옛 판은 지우지 않고 `old` 로 옮긴다**(형님 지시). 이름이 겹치면 번호를
  붙여 둘 다 남긴다 — 덮지 않는다.
★ **새 판을 먼저 놓고 옛 판을 나중에 옮긴다.** 중간에 끊겨도 폴더가 비지 않는다
  (옛 판이 하나 더 남을 뿐이고 다음 회차가 옮긴다).
★ 처음 한 번: 표지가 없으면 **이미 게시된 JSON 의 지문**을 읽어 같으면 표지만
  적는다 — 내용이 같은 판을 다시 올리지 않는다.
★ 못 하면 **못 했다고 말한다**([169]) — 폴더가 없으면 만들지 않고 이유를 돌려준다
  (Z: 가 끊긴 날 로컬에 엉뚱한 폴더가 생기면 안 된다).

사람: `python sop_publish.py` (바뀐 묶음만) · `--force` (내용이 같아도 새 판)
"""
import datetime
import glob
import hashlib
import io
import json
import os
import re
import shutil
import sys
import tempfile

if hasattr(sys.stdout, "reconfigure"):          # 무인 회차는 sys.stdout 이 None 이다([235])
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

BASE = os.path.dirname(os.path.abspath(__file__))
if BASE not in sys.path:
    sys.path.insert(0, BASE)

import sop_export as X                            # noqa: E402
import sop_store as S                             # noqa: E402

MARK = os.path.join(BASE, "reports", ".절차서_게시.json")
JSON_DIR = "3. 유니웍스 넘김 자료(JSON)"
OLD_DIR = "old"
RESERVED = {1, 2, 3, 4}                           # 묶음 아닌 뜻이 정해진 자리 포함


def _stem(book):
    return re.sub(r'[\\/:*?"<>|\s]+', "_", book)


def content_key(schema, rows):
    """게시 판단에 쓰는 지문 — `만든때` 는 뺀다(그것만 매번 바뀐다)."""
    raw = json.dumps({"스키마판": schema, "절차서": rows}, ensure_ascii=False,
                     sort_keys=True, default=str)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def store_key(book, store=None):
    """store 는 절차서 파일 경로다(없으면 정본)."""
    rows = X._rows(None, None, None, store, book)
    return content_key(X.UNIWORKS_SCHEMA, [{k: r.get(k) for k in S.칸} for r in rows]), len(rows)


def published_key(path):
    """이미 게시된 JSON 의 지문 — 못 읽으면 None(모름)."""
    try:
        with io.open(path, encoding="utf-8") as fh:
            d = json.load(fh)
        return content_key(d.get("스키마판"), d.get("절차서") or [])
    except Exception:
        return None


def _read_mark():
    try:
        with io.open(MARK, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _write_mark(m):
    os.makedirs(os.path.dirname(MARK), exist_ok=True)
    tmp = MARK + ".tmp"
    with io.open(tmp, "w", encoding="utf-8") as fh:
        json.dump(m, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, MARK)


def book_folder(root, book):
    """그 묶음의 번호 폴더 — 있으면 그것, 없으면 다음 번호로 새로 정한다(만들지는 않는다)."""
    used = set()
    for name in os.listdir(root):
        m = re.match(r"^(\d+)\.\s*(.*)$", name)
        if not m:
            continue
        used.add(int(m.group(1)))
        if m.group(2).strip() == book:
            return os.path.join(root, name)
    n = max(used | RESERVED) + 1
    return os.path.join(root, "%d. %s" % (n, book))


def _to_old(root, path):
    old = os.path.join(root, OLD_DIR)
    os.makedirs(old, exist_ok=True)
    dst = os.path.join(old, os.path.basename(path))
    i = 2
    while os.path.exists(dst):                     # 덮지 않는다 — 둘 다 남긴다
        b, e = os.path.splitext(os.path.basename(path))
        dst = os.path.join(old, "%s(%d)%s" % (b, i, e))
        i += 1
    shutil.move(path, dst)
    return dst


def publish(root=None, force=False, today=None, store=None):
    """바뀐 묶음만 게시한다. 돌려주는 것: 사람이 읽을 한 줄(할 일 없으면 빈 글자)."""
    root = root or _root()
    d = S.load(store)                              # store 는 경로다(시험용) · 없으면 정본
    books = [b.get("이름") for b in (S.books(d) or []) if b.get("이름")]
    if not books:
        return "절차서 게시: 묶음이 없다 — 게시할 것이 없다"
    mark = _read_mark()
    todo = []
    for book in books:
        key, n = store_key(book, store)
        if n and (force or mark.get(book) != key):
            todo.append((book, key, n))
    if not todo:
        return ""                                  # Z: 를 안 만진다([168])
    if not os.path.isdir(root):
        return "절차서 게시: 폴더에 못 닿았다(%s) — 다음 회차에 다시 본다" % root
    stamp = today or datetime.date.today().strftime("%Y%m%d")
    done, same, blocked = [], [], []
    for book, key, n in todo:
        stem = _stem(book)
        jdir = os.path.join(root, JSON_DIR)
        bdir = book_folder(root, book)
        olds_json = sorted(glob.glob(os.path.join(jdir, stem + "_*.json")))
        # 처음 한 번 — 이미 같은 내용이 게시돼 있으면 표지만 적는다.
        if not force and mark.get(book) is None and olds_json \
                and published_key(olds_json[-1]) == key:
            mark[book] = key
            same.append(book)
            continue
        olds = [p for p in glob.glob(os.path.join(bdir, stem + "_*.*"))] + olds_json
        tmpd = tempfile.mkdtemp(prefix="sop_pub_")
        try:
            made = []
            for fmt, dst_dir in (("docx", bdir), ("pptx", bdir), ("json", jdir)):
                p, _k = X.export(fmt, os.path.join(tmpd, "%s_%s.%s" % (stem, stamp, fmt)),
                                 묶음=book, store=store)
                made.append((p, dst_dir))
            # ★ 회사 공유 폴더로 나가기 **전에** 개인 사업 이름이 섞였는지 거른다
            #   (2026-10-03 형님 지시 · 판정은 outbound_guard 한 곳).  걸리면 그 묶음은
            #   안 올리고 표지도 안 올린다 — 고친 뒤 다음 회차가 다시 본다.
            import outbound_guard
            걸림 = [os.path.basename(p) for p, _d in made if _has_banned(p, outbound_guard)]
            if 걸림:
                blocked.append("%s(%s)" % (book, ", ".join(걸림)))
                continue
            os.makedirs(bdir, exist_ok=True)
            os.makedirs(jdir, exist_ok=True)
            new = []
            for p, dst_dir in made:
                dst = os.path.join(dst_dir, os.path.basename(p))
                shutil.copy2(p, dst + ".part")
                os.replace(dst + ".part", dst)
                new.append(os.path.normcase(dst))
        finally:
            shutil.rmtree(tmpd, ignore_errors=True)
        moved = 0
        for p in olds:                             # 새 판을 놓은 **뒤에** 옮긴다
            if os.path.normcase(p) not in new and os.path.exists(p):
                _to_old(root, p)
                moved += 1
        mark[book] = key
        done.append("%s %d건(옛 판 %d개 → old)" % (book, n, moved))
    _write_mark(mark)
    parts = []
    if done:
        parts.append("절차서 새 판 게시: " + " · ".join(done))
    if blocked:
        parts.append("★ 개인 사업 이름이 섞여 안 올림: " + " · ".join(blocked))
    if same:
        parts.append("이미 같은 판이 게시돼 있어 표지만 적음: " + " · ".join(same))
    return " / ".join(parts)


def _has_banned(path, guard):
    """나가는 파일에 금지어가 있나 — 낱말 판정은 `outbound_guard.scan_text` 한 곳이다([162]).

    ★ 워드·PPT 는 zip 이라 글자로 열면 '못읽음'이 된다 — 그것을 '깨끗'으로 치면
      검사가 있으나 마나다([169]). 그래서 zip 안 xml 을 열어 같은 판정에 넘긴다.
    ★ 못 읽으면 **걸림으로 친다** — 회사 공유 폴더로 나가는 길이라 모를 때는 안 내보낸다.
    """
    import zipfile
    try:
        if path.lower().endswith((".docx", ".pptx", ".xlsx")):
            with zipfile.ZipFile(path) as z:
                return any(guard.scan_text(z.read(n).decode("utf-8", "replace"))
                           for n in z.namelist() if n.endswith(".xml"))
        with io.open(path, encoding="utf-8") as fh:
            return bool(guard.scan_text(fh.read()))
    except Exception:
        return True


def _root():
    import source_dirs as SD
    return SD.SOP_DIR


def main(argv):
    force = "--force" in argv
    try:
        msg = publish(force=force)
    except Exception as e:                         # 못 했으면 못 했다고 말한다([169])
        print("절차서 게시 실패: %s: %s" % (type(e).__name__, e))
        return 1
    print(msg or "절차서 게시: 바뀐 묶음 없음(공유 폴더는 안 건드렸다)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
