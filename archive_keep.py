# -*- coding: utf-8 -*-
"""
archive_keep.py — 나중에 **복구하거나 이어서 코딩할 때** 필요한 것만 골라 서버에 보관한다
===============================================================================
사용자 지시(2026-07-30): "앞으로 로그 기록 등 나중에 복구하거나 코딩을 진행할 때 필요한
자료는 별도로 보관해줘(다른 좋은 방법이 있으면 좋은 방법으로)"

## 더 좋은 방법을 골랐다: 파일을 쌓는 게 아니라 **git bundle**
  reports/ 를 날짜별로 복사해 쌓으면 용량만 늘고 정작 코드는 못 되살린다.
  `git bundle` 은 **저장소 전체(모든 커밋·브랜치·이력)를 한 파일**로 만든다.
  그 파일 하나만 있으면 어디서든 `git clone <파일>` 로 완전히 되살아난다 —
  GitHub 이 막히거나 PC가 죽어도. 그래서 코드는 bundle 로, 나머지는 아래 원칙으로 나눈다.

## 무엇을 보관하고 무엇을 버리는가 (핵심은 '되살릴 수 있는가')
  보관: · git bundle(코드+이력 전부)            ← 코딩을 이어서 하려면 이것만 있으면 된다
        · db/ledger_queue.db SQLite 일관 백업   ← 아직 엑셀에 안 들어간 확정 입력·UX 근거
        · reports/*.md|csv|json 중 **사실 기록**(대조 결과·종합리포트·세션인계·자료현황)
        · updates/applied_*.json               ← 원장에 무엇을 왜 썼는지의 증거
        · 19시트 인수인계 텍스트 사본 · INCIDENTS.md · AGENTS.md
        · 관리대장 최신본 1개(원장 자체는 이미 vN 으로 서버에 남지만, 짝을 맞춰 둔다)
  버림: · node_modules·__pycache__·.tmp — 다시 만들면 된다
        · band/cache·ocr_cache 원본 — 크고, 없어도 재수집 가능하며 사실은 리포트에 남는다
        · **비밀키(config/*.json)** — 절대 보관하지 않는다(규칙 1). 무엇이 필요한지 목록만 남긴다.

## 어디에 두는가
  PC가 죽으면 PC 안의 백업은 같이 죽는다. 그래서 **회사 서버(Z:)** 에 둔다.
  보관 위치: `<0. 원본 자료>/_보관/YYYY-MM-DD/`

## 얼마나 남기는가
  최근 14일은 매일, 그 앞은 **각 달 1일치만** 남긴다(무한 증식 방지).
  bundle 은 내용이 같으면(직전과 같은 커밋) 다시 만들지 않는다.

사용
  python archive_keep.py            # 오늘치 보관 + 정리
  python archive_keep.py --dry      # 무엇을 담을지만 보기
  python archive_keep.py --self-test
"""
import sys, os, re, json, glob, shutil, subprocess, hashlib
from datetime import datetime, date, timedelta

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

KEEP_DAILY_DAYS = 14        # 최근 며칠은 매일 남기나
SKIP_DIRS = {"node_modules", "__pycache__", ".git", "ocr_cache", "cache", "docs_inbox"}
# 리포트 중 '사실 기록'만. 이름이 바뀌어도 확장자와 위치로 걸러진다.
REPORT_EXT = (".md", ".csv", ".json", ".txt")
SECRET_RE = re.compile(r"(API_CERT_KEY|client_secret|worker_token|enqueue_token|password|pin)",
                       re.I)


def archive_root():
    from source_dirs import ERP_DIR
    return os.path.join(os.path.dirname(ERP_DIR), "_보관")


# ── 순수 판정 (합성 검증 대상) ─────────────────────────────────
def keep_days(days, today=None, keep_daily=KEEP_DAILY_DAYS):
    """보관 폴더 날짜 목록 → 남길 날짜 집합.

    최근 keep_daily 일은 전부, 그 앞은 **각 달의 가장 이른 하루**만 남긴다.
    ★ 매일 쌓기만 하면 1년에 365벌이 된다. 오래된 것은 '그 달에 이런 상태였다' 만 있으면 된다."""
    today = today or date.today()
    recent, older = set(), {}
    for d in days:
        if not isinstance(d, date):
            continue
        if (today - d).days <= keep_daily:
            recent.add(d)
        else:
            key = (d.year, d.month)
            if key not in older or d < older[key]:
                older[key] = d
    return recent | set(older.values())


def parse_day(name):
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", name or "")
    if not m:
        return None
    try:
        return date(*(int(x) for x in m.groups()))
    except ValueError:
        return None


def has_secret(text):
    """비밀키가 섞였는지 본다 — 보관물은 사람 손을 여러 번 타므로 여기서 한 번 더 막는다."""
    return bool(SECRET_RE.search(text or ""))


def wanted(path, root):
    """이 파일을 보관할까. 되살릴 수 있는 것(캐시·빌드산출물)은 담지 않는다."""
    rel = os.path.relpath(path, root).replace("\\", "/")
    parts = rel.split("/")
    if any(p in SKIP_DIRS for p in parts):
        return False
    if parts[0] == "config":                  # 비밀키는 어떤 경우에도 담지 않는다
        return False
    if rel.endswith((".tmp", ".tmp.xlsx", ".pyc", ".lock")):
        return False
    if parts[0] == "reports":
        return rel.endswith(REPORT_EXT)
    if parts[0] == "updates":
        return os.path.basename(rel).startswith("applied_")
    return rel in ("AGENTS.md", "INCIDENTS.md", "CLAUDE.md")


def bundle_action(head, current_marker, dst_exists, previous_heads):
    """오늘 bundle을 유지/교체/생성/생략할지 정한다(순수 함수)."""
    if dst_exists:
        return "keep" if current_marker == head else "replace"
    if head in set(previous_heads or []):
        return "skip"
    return "create"


# ── 실행 ─────────────────────────────────────────────────────
def git_bundle(dst, dry=False):
    """저장소 전체를 한 파일로. 같은 날 HEAD가 바뀌면 원자적으로 최신 bundle로 교체한다."""
    try:
        head = subprocess.run(["git", "rev-parse", "HEAD"], cwd=ROOT, capture_output=True,
                              text=True, timeout=60).stdout.strip()
    except Exception:
        return None, "git 없음"
    if not head:
        return None, "커밋 없음"
    stamp = os.path.join(os.path.dirname(dst), ".bundle_head")
    current = ""
    try:
        current = open(stamp, encoding="utf-8").read().strip()
    except OSError:
        pass
    previous = []
    for cand in glob.glob(os.path.join(archive_root(), "*", ".bundle_head")):
        if os.path.normcase(os.path.abspath(cand)) == os.path.normcase(os.path.abspath(stamp)):
            continue
        try:
            value = open(cand, encoding="utf-8").read().strip()
            if value:
                previous.append(value)
        except OSError:
            pass
    action = bundle_action(head, current, os.path.isfile(dst), previous)
    if dry:
        return None, f"bundle {action} 예정 (HEAD {head[:8]})"
    if action == "keep":
        return dst, f"bundle 최신 유지 (HEAD {head[:8]})"
    if action == "skip":
        return None, f"bundle 생략 — 이전 보관과 같은 HEAD {head[:8]}"

    os.makedirs(os.path.dirname(dst), exist_ok=True)
    temp = dst + f".tmp.{os.getpid()}"
    marker_temp = stamp + f".tmp.{os.getpid()}"
    try:
        if os.path.exists(temp):
            os.unlink(temp)
        r = subprocess.run(["git", "bundle", "create", temp, "--all"], cwd=ROOT,
                           capture_output=True, text=True, timeout=900)
        if r.returncode != 0 or not os.path.exists(temp):
            return None, f"bundle 실패: {(r.stderr or '')[:80]}"
        # 기존 bundle은 새 파일 생성이 완전히 끝난 뒤 한 번에 교체한다.
        os.replace(temp, dst)
        with open(marker_temp, "w", encoding="utf-8") as fh:
            fh.write(head)
        os.replace(marker_temp, stamp)
    finally:
        for leftover in (temp, marker_temp):
            try:
                if os.path.exists(leftover):
                    os.unlink(leftover)
            except OSError:
                pass
    # ★ bundle 은 **커밋된 것만** 담는다. 미커밋 변경은 들어가지 않으므로 반드시 알린다.
    #   (실측 확인: bundle 복구 시 커밋 전 새 파일 2개가 없었다 — 조용하면 복구 때 사고가 된다)
    dirty = uncommitted()
    note = f"bundle {os.path.getsize(dst) // 1024}KB (HEAD {head[:8]}, 커밋 {commit_count()}개)"
    if dirty:
        note += f" · ★ 미커밋 {dirty}개는 담기지 않았다 — 커밋 후 다시 보관할 것"
    return dst, note


def uncommitted():
    try:
        r = subprocess.run(["git", "status", "--porcelain"], cwd=ROOT, capture_output=True,
                           text=True, timeout=120)
        return len([l for l in (r.stdout or "").splitlines()
                    if l.strip() and "outputs/" not in l and "/tmp/" not in l])
    except Exception:
        return 0


def commit_count():
    try:
        r = subprocess.run(["git", "rev-list", "--count", "HEAD"], cwd=ROOT,
                           capture_output=True, text=True, timeout=120)
        return (r.stdout or "0").strip()
    except Exception:
        return "?"


def collect(day_dir, dry=False):
    n = size = skipped = 0
    for path in glob.glob(os.path.join(ROOT, "**", "*"), recursive=True):
        if not os.path.isfile(path) or not wanted(path, ROOT):
            continue
        rel = os.path.relpath(path, ROOT)
        if rel.endswith((".md", ".json", ".txt", ".csv")) and os.path.getsize(path) < 2_000_000:
            try:
                if has_secret(open(path, encoding="utf-8", errors="replace").read(200_000)):
                    skipped += 1
                    continue                  # 비밀키 형태가 보이면 담지 않는다
            except OSError:
                pass
        n += 1
        size += os.path.getsize(path)
        if not dry:
            dst = os.path.join(day_dir, rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(path, dst)
    return n, size, skipped


def ledger_copy(day_dir, dry=False):
    try:
        from workbook_patch import latest_master
        path, ver = latest_master()
    except Exception:
        return "관리대장 없음"
    if dry:
        return f"관리대장 v{ver} 예정"
    dst = os.path.join(day_dir, os.path.basename(path))
    if not os.path.exists(dst):
        shutil.copy2(path, dst)
    return f"관리대장 v{ver}"


def ledger_db_copy(day_dir, dry=False):
    """WAL 사용 중인 입력 DB를 파일복사가 아닌 SQLite backup API로 보관한다."""
    dst = os.path.join(day_dir, "db", "ledger_queue.db")
    if dry:
        return "입력·UX DB 예정"
    try:
        import ledger_db
        ledger_db.backup_to(dst)
        return f"입력·UX DB {os.path.getsize(dst) // 1024}KB"
    except Exception as exc:
        return f"입력·UX DB 보관 실패: {str(exc)[:100]}"


def prune(dry=False):
    base = archive_root()
    if not os.path.isdir(base):
        return "정리할 것 없음"
    days = {}
    for name in os.listdir(base):
        d = parse_day(name)
        if d:
            days[d] = os.path.join(base, name)
    keep = keep_days(list(days))
    gone = [p for d, p in days.items() if d not in keep]
    if not dry:
        for p in gone:
            shutil.rmtree(p, ignore_errors=True)
    return f"보관 {len(days)}일 중 {len(gone)}일 정리{'(dry)' if dry else ''} · 유지 {len(keep)}일"


def manifest(day_dir, lines):
    with open(os.path.join(day_dir, "복구방법.md"), "w", encoding="utf-8") as fh:
        fh.write(f"# 복구 방법 ({datetime.now():%Y-%m-%d %H:%M})\n\n")
        fh.write("## 코드를 되살리려면 — bundle 하나로 끝난다\n\n```bash\n")
        fh.write("git clone coupang_repo.bundle 복구폴더\n```\n\n")
        fh.write("모든 커밋·브랜치·이력이 그대로 살아난다. GitHub 이 막혀도 된다.\n\n")
        fh.write("## 여기에 **없는** 것 (일부러 안 담았다)\n\n")
        fh.write("- `config/*.json` — 비밀키(이카운트 API키·PIN·밴드 토큰). 규칙 1로 절대 보관하지 않는다.\n")
        fh.write("  복구 시 새로 발급하거나 담당자에게 받아 다시 만든다. 필요한 파일:\n")
        fh.write("  `ecount_config.json`, `webapp.json`, `gcal.json`, `erp_allowed_ips.json`,\n")
        fh.write("  `cloud_queue.local.json` (예시는 `*.example.json` 참고)\n")
        fh.write("- `node_modules`·`__pycache__`·OCR/밴드 캐시 — 다시 만들면 된다.\n\n")
        fh.write("## 담긴 것\n\n")
        for ln in lines:
            fh.write(f"- {ln}\n")


def self_test():
    bad = 0
    today = date(2026, 7, 30)
    days = [date(2026, 7, 30), date(2026, 7, 20), date(2026, 7, 5), date(2026, 7, 3),
            date(2026, 6, 28), date(2026, 6, 2), date(2026, 5, 15)]
    keep = keep_days(days, today)
    want = {date(2026, 7, 30), date(2026, 7, 20),      # 최근 14일
            date(2026, 7, 3), date(2026, 6, 2), date(2026, 5, 15)}   # 달별 가장 이른 하루
    if keep != want:
        print("  [FAIL] 보관정책", sorted(keep)); bad += 1
    if parse_day("2026-07-30") != date(2026, 7, 30) or parse_day("reports") is not None:
        print("  [FAIL] 날짜 파싱"); bad += 1
    # 비밀키는 어떤 경로로도 담기지 않는다
    if wanted(os.path.join(ROOT, "config", "ecount_config.json"), ROOT):
        print("  [FAIL] config 를 담으려 한다"); bad += 1
    if not has_secret('{"API_CERT_KEY": "abc"}') or not has_secret("client_secret=x"):
        print("  [FAIL] 비밀키 탐지"); bad += 1
    if has_secret("정산 703행 4.6억"):
        print("  [FAIL] 정상 문서를 비밀로 오판"); bad += 1
    for got, want in (
        (bundle_action("new", "new", True, []), "keep"),
        (bundle_action("new", "old", True, []), "replace"),
        (bundle_action("new", "", False, ["new"]), "skip"),
        (bundle_action("new", "", False, ["old"]), "create"),
    ):
        if got != want:
            print(f"  [FAIL] bundle 갱신 판정 {got} != {want}"); bad += 1
    # 되살릴 수 있는 것은 담지 않는다
    for rel in ("outputs/x/node_modules/a.js", "band/cache/90610953.json",
                "reports/a.tmp.xlsx", "webapp/__pycache__/x.pyc"):
        if wanted(os.path.join(ROOT, *rel.split("/")), ROOT):
            print(f"  [FAIL] {rel} 를 담으려 한다"); bad += 1
    # 담아야 하는 것은 담는다
    for rel in ("reports/종합리포트_x.md", "updates/applied_1.json", "AGENTS.md",
                "INCIDENTS.md"):
        if not wanted(os.path.join(ROOT, *rel.split("/")), ROOT):
            print(f"  [FAIL] {rel} 를 빠뜨린다"); bad += 1
    print("archive_keep self-test:", "OK" if not bad else f"{bad}건 실패")
    return bad == 0


def main():
    if "--self-test" in sys.argv:
        sys.exit(0 if self_test() else 1)
    dry = "--dry" in sys.argv
    day_dir = os.path.join(archive_root(), f"{date.today():%Y-%m-%d}")
    lines = []
    b, msg = git_bundle(os.path.join(day_dir, "coupang_repo.bundle"), dry)
    lines.append(msg)
    n, size, skipped = collect(day_dir, dry)
    lines.append(f"기록 파일 {n}개 {size // 1024}KB" + (f" · 비밀키 의심 {skipped}개 제외" if skipped else ""))
    lines.append(ledger_copy(day_dir, dry))
    lines.append(ledger_db_copy(day_dir, dry))
    if not dry:
        manifest(day_dir, lines)
    lines.append(prune(dry))
    print(f"보관 위치: {day_dir}{' (dry — 실제 복사 없음)' if dry else ''}")
    for ln in lines:
        print("  ·", ln)


if __name__ == "__main__":
    main()
