# -*- coding: utf-8 -*-
"""
ingest.py — 밴드 수집 덤프를 받아 관리대장까지 한 번에 반영
================================================================================
브라우저 수집기(collect_band.js)가 만든 dump_*.json 을 cache/ 에 넣고 이걸 실행하면:
  1) dump → 캐시 변환(convert_dump)  — 게시일 파싱, raw_*.json 보존
  2) 캐시 월별 현황 출력             — 어느 달이 새로 들어왔는지
  3) 24_밴드업무추출 시트 반영        — --sheet 지정 시
  4) 각 월 백필 가능 여부 안내        — 행 용량 부족하면 확장 명령까지 알려준다

실행
  python band/ingest.py                       # 변환 + 현황만
  python band/ingest.py --sheet               # + 24시트 반영
  python band/ingest.py --sheet --backfill    # + 월별 백필까지(02/04 시트 등록)
"""
import sys, os, re, json, glob, subprocess, collections
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CACHE = os.path.join(HERE, "cache")
PY = sys.executable


def run(args, label):
    r = subprocess.run([PY] + args, cwd=ROOT, capture_output=True, text=True,
                       encoding="utf-8", errors="replace",
                       env={**os.environ, "PYTHONIOENCODING": "utf-8"}, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    out = "\n".join(l for l in (r.stdout or "").splitlines()
                    if "Warning" not in l and "warn(" not in l)
    print(f"— {label}: {(out.strip().splitlines() or ['(출력 없음)'])[-1]}")
    return r.returncode == 0, out


def extract_images():
    """수집 덤프에 담겨 온 문서 사진(base64)을 docs_inbox/ 로 풀어 놓는다.
    밴드 이미지 URL은 로그인 쿠키가 있어야 열려서, 브라우저에서 받아 온 본문을 그대로 쓴다."""
    import base64
    out = os.path.join(HERE, "docs_inbox")
    os.makedirs(out, exist_ok=True)
    n = 0
    files = glob.glob(os.path.join(CACHE, "raw_*.json")) + glob.glob(os.path.join(CACHE, "dump_*.json"))
    try:
        from source_dirs import band_dump_dirs
        for folder in band_dump_dirs():
            files += glob.glob(os.path.join(folder, "**", "*.json"), recursive=True)
    except Exception:
        pass
    for f in dict.fromkeys(files):
        try:
            d = json.load(open(f, encoding="utf-8"))
        except Exception:
            continue
        posts = d.get("posts") or {}
        for pid, v in (posts.items() if isinstance(posts, dict) else enumerate(posts)):
            for i, data in enumerate(v.get("imageData") or []):
                m = re.match(r"data:image/(\w+);base64,(.+)$", str(data), re.S)
                if not m:
                    continue
                dst = os.path.join(out, f"band_{d.get('band','x')}_{pid}_{i}.{m.group(1)}")
                if os.path.exists(dst):
                    continue
                try:
                    open(dst, "wb").write(base64.b64decode(m.group(2)))
                    n += 1
                except Exception:
                    pass
    return n


def cache_months():
    """캐시에 들어 있는 게시글의 월별 분포"""
    out = {}
    for p in sorted(glob.glob(os.path.join(CACHE, "*.json"))):
        b = os.path.basename(p)
        if b.startswith("raw_") or b.startswith("dump_"):
            continue
        try:
            d = json.load(open(p, encoding="utf-8"))
        except Exception:
            continue
        posts = d.get("posts") or {}
        c = collections.Counter()
        for v in (posts.values() if isinstance(posts, dict) else posts):
            ts = v.get("created_at")
            if ts:
                c[datetime.fromtimestamp(ts / 1000).strftime("%Y-%m")] += 1
        out[d.get("band_name") or b] = c
    return out


def main():
    args = sys.argv[1:]
    from convert_dump import dump_files
    dumps = dump_files()
    if dumps:
        print(f"수집 덤프 {len(dumps)}개 발견 → 변환")
        run([os.path.join(HERE, "convert_dump.py")], "dump → 캐시 변환")
    else:
        print("새 덤프 없음 — 기존 캐시로 진행 "
              "(밴드에서 새로 수집하려면 band/collect_band.js 를 브라우저 콘솔에 붙여넣으세요)")

    got = extract_images()
    if got:
        print(f"문서 사진 {got}장 → band/docs_inbox/ (다음 단계에서 OCR로 읽습니다)")

    months = cache_months()
    allm = collections.Counter()
    print("\n■ 캐시 월별 게시글")
    for name, c in months.items():
        print(f"  {name}: {dict(sorted(c.items()))}")
        allm.update(c)
    if not allm:
        sys.exit("캐시가 비어 있습니다.")
    lo, hi = min(allm), max(allm)
    print(f"  전체 {sum(allm.values())}건 · {lo} ~ {hi}")

    todo = sorted(allm)
    if "--sheet" in args:
        # ★ **--month 를 붙여 월별로 돌리면 안 된다.** band_extract 의 --sheet 는 24시트를
        #   그 달 것으로 **통째로 덮어쓴다**. 월별 반복이면 마지막 달만 남아 나머지가 사라진다
        #   (2026-07-27 실사고: 4,223행 → 506행. 2025-12~2026-05가 통째로 날아갔다).
        #   한 번만, 월 지정 없이 부른다 — 캐시 전체가 그대로 들어간다.
        run([os.path.join(ROOT, "band_extract.py"), "--sheet"], "24시트 반영(전체)")

    if "--backfill" in args:
        print("\n■ 월별 백필")
        for mo in todo:
            ok, out = run([os.path.join(ROOT, "backfill_rows.py"), "--month", mo, "--apply"], f"백필 {mo}")
            for line in out.splitlines():
                if "용량 초과" in line or "expand_rows" in line:
                    print(f"    ★ {line.strip()}")
    else:
        print("\n다음 단계:  python band/ingest.py --sheet --backfill   (24시트 반영 + 02/04 등록)")


if __name__ == "__main__":
    main()
