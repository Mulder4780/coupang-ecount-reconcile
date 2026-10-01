# -*- coding: utf-8 -*-
"""
ops_manual.py — 매뉴얼·로그·오류를 **형님이 지정한 폴더**에 올린다
===============================================================================
형님 지시(2026-10-02): *"모든건 메뉴얼로 만들고 메뉴얼대로 진행해 내가 지정한 폴더에
만들고 로그관리나 오류 등 전부 문서화해서 관리하고 …"*
형님이 고르신 자리: `<2. CSOS DATA>/12. 운영 매뉴얼`

## 무엇을 하나 — 셋
① **매뉴얼 게시**: `reports/_운영매뉴얼_초안/` 의 문서를 지정 폴더로 올린다.
   **바뀐 것만** 올리고, 바뀌면 옛 판을 `05_과거기록/` 으로 옮긴다(지우지 않는다).
② **하루 한 장 로그**: 원본 로그(watchdog·daily_run·realtime)를 읽어
   `02_로그/YYYY/MM/YYYY-MM-DD.md` 한 장으로 줄인다.
③ **오류 내보내기**: 이미 있는 오류 사전·기록(`error_book`)을 `03_오류/` 로 내보낸다.

## 왜 '내보내기'인가 — **새로 만들지 않았다**([162]·[172])
오류 사전과 기록은 `error_book.py` 가 이미 갖고 있다. 두 곳에서 각자 관리하면
언젠가 갈리고, 갈린 뒤엔 어느 쪽이 맞는지 아무도 모른다. 그래서 **정본은 그대로
두고 사본만** 올린다.

## 못 닿으면 — **"올렸다"고 적지 않는다**([169])
Z: 가 끊기면 아무것도 안 하고 *"공유 폴더에 못 닿았다"* 를 돌려준다.
다음 회차가 다시 본다. 조용히 넘어가면 '올렸는데 왜 없지'가 된다.

## 사람이 쓰는 법
  python ops_manual.py            # 매뉴얼 게시 + 어제 로그 한 장 + 오류 내보내기
  python ops_manual.py --print    # 무엇이 있나만 본다
  python ops_manual.py --dry      # 할 일만 말하고 안 쓴다
  python ops_manual.py --day 2026-10-01   # 그 날짜 로그 한 장만
"""
import argparse
import datetime
import glob
import hashlib
import io
import os
import re
import shutil
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
if BASE not in sys.path:
    sys.path.insert(0, BASE)

초안 = os.path.join(BASE, "reports", "_운영매뉴얼_초안")
LOG_SRC = {
    "워치독(30분)": os.path.join(BASE, "reports", "watchdog_log.txt"),
    "일일대조(09:50)": os.path.join(BASE, "reports", "daily_run_log.txt"),
    "문제감시(4시간)": os.path.join(BASE, "reports", "realtime_monitor_log.txt"),
}


def manual_root():
    try:
        import ops_backup
        return ops_backup.manual_root()
    except Exception:
        return None


def _sha(p):
    h = hashlib.sha256()
    with io.open(p, "rb") as fh:
        for b in iter(lambda: fh.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


# ── ① 매뉴얼 게시 ──────────────────────────────────────────────
def publish_manual(root=None, dry=False, today=None):
    root = root if root is not None else manual_root()
    if not root:
        return "매뉴얼 게시: 지정 폴더를 못 찾았다"
    if not os.path.isdir(os.path.dirname(root)):
        return "매뉴얼 게시: 공유 폴더에 못 닿았다 — 다음 회차에 다시 본다"
    if not os.path.isdir(초안):
        return "매뉴얼 게시: 올릴 초안이 없다(%s)" % 초안
    today = today or datetime.date.today().isoformat()
    올림, 그대로 = [], 0
    for src in sorted(glob.glob(os.path.join(초안, "**", "*.md"), recursive=True)):
        rel = os.path.relpath(src, 초안)
        dst = os.path.join(root, rel)
        if os.path.isfile(dst) and _sha(dst) == _sha(src):
            그대로 += 1
            continue
        if dry:
            올림.append(rel)
            continue
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        if os.path.isfile(dst):
            # ★ 옛 판을 지우지 않는다 — 무엇이 바뀌었는지 되짚을 근거다.
            옛 = os.path.join(root, "05_과거기록", "매뉴얼", today,
                             rel.replace(os.sep, "__"))
            os.makedirs(os.path.dirname(옛), exist_ok=True)
            shutil.copy2(dst, 옛)
        shutil.copy2(src, dst + ".part")
        os.replace(dst + ".part", dst)
        올림.append(rel)
    if not 올림:
        return "매뉴얼 게시: 바뀐 문서 없음(%d장 그대로)" % 그대로
    return "매뉴얼 게시: %d장 %s(그대로 %d장) — %s" % (
        len(올림), "올릴 것" if dry else "올림", 그대로,
        ", ".join(os.path.basename(x) for x in 올림[:4]))


# ── ② 하루 한 장 로그 ──────────────────────────────────────────
_날짜 = re.compile(r"^\[(\d{2})-(\d{2})\s+(\d{2}):(\d{2})\]")


def _그날줄(path, day):
    """그 날짜 줄만 뽑는다. 파일이 없으면 **빈 목록이 아니라 None**([169])."""
    if not os.path.isfile(path):
        return None
    y, m, d = day.split("-")
    want = "[%s-%s " % (m, d)
    out = []
    try:
        with io.open(path, encoding="utf-8", errors="replace") as fh:
            for line in fh:
                if line.startswith(want):
                    out.append(line.rstrip("\n"))
    except Exception:
        return None
    return out


def _경보뽑기(lines):
    """한 줄에 여러 단계가 ' | ' 로 붙어 온다 — 경보처럼 보이는 조각만 센다."""
    표 = {}
    for line in lines:
        몸 = line.split("] ", 1)[-1]
        for 조각 in 몸.split(" | "):
            t = 조각.strip()
            if not t:
                continue
            if re.search(r"정상|없음|그대로|최신|준비됨|생략|0개|0건", t):
                continue
            표[t[:110]] = 표.get(t[:110], 0) + 1
    return 표


def day_log(day, root=None, dry=False):
    root = root if root is not None else manual_root()
    조각 = ["# %s 하루 기록" % day, "",
           "*기계가 자동으로 만든 요약이다. 자세한 것은 `ecount/reports/` 의 원본 로그에 있다.*",
           ""]
    못읽음 = []
    어제표 = {}
    try:
        어제 = (datetime.date.fromisoformat(day) - datetime.timedelta(days=1)).isoformat()
        ls = _그날줄(LOG_SRC["워치독(30분)"], 어제)
        어제표 = _경보뽑기(ls or [])
    except Exception:
        pass

    for 이름, p in LOG_SRC.items():
        lines = _그날줄(p, day)
        if lines is None:
            못읽음.append(이름)
            continue
        조각.append("## %s — %d회" % (이름, len(lines)))
        조각.append("")
        표 = _경보뽑기(lines)
        새것 = [k for k in 표 if k not in 어제표]
        if 이름.startswith("워치독") and 표:
            if 새것:
                조각.append("**어제 없던 것**")
                조각.append("")
                for k in sorted(새것, key=lambda x: -표[x])[:12]:
                    조각.append("- %s *(%d회)*" % (k, 표[k]))
                조각.append("")
            계속 = [k for k in 표 if k in 어제표]
            if 계속:
                조각.append("계속되는 것 %d가지 — %s" %
                            (len(계속), ", ".join(sorted(계속, key=lambda x: -표[x])[:3])))
                조각.append("")
        elif 표:
            for k in sorted(표, key=lambda x: -표[x])[:10]:
                조각.append("- %s *(%d회)*" % (k, 표[k]))
            조각.append("")
        if not 표:
            조각.append("특별한 것 없음")
            조각.append("")

    # DB 백업이 그날 떴나
    try:
        import ops_backup
        base = os.path.join(root, ops_backup.BACKUP_SUB) if root else None
        있나 = bool(base and os.path.isdir(os.path.join(base, day)))
        대기 = os.path.isdir(os.path.join(ops_backup.WAIT_DIR, day))
        조각.append("## DB 백업")
        조각.append("")
        조각.append("- 공유 폴더: %s" % ("떴다" if 있나 else "없다"))
        if 대기:
            조각.append("- **이 PC 대기함에 있다** — 공유 폴더에 못 닿아 아직 못 올렸다")
        조각.append("")
    except Exception:
        pass

    if 못읽음:
        조각.append("## 못 읽은 것")
        조각.append("")
        조각.append("- %s — 로그 파일을 못 읽었다. **'문제 없음'이 아니라 '모름'이다.**"
                    % ", ".join(못읽음))
        조각.append("")

    글 = "\n".join(조각) + "\n"
    if not root or not os.path.isdir(os.path.dirname(root)):
        return "하루 로그(%s): 공유 폴더에 못 닿아 못 올렸다" % day, 글
    dst = os.path.join(root, "02_로그", day[:4], day[5:7], day + ".md")
    if dry:
        return "하루 로그(%s): 올릴 것 → %s" % (day, dst), 글
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    io.open(dst + ".part", "w", encoding="utf-8").write(글)
    os.replace(dst + ".part", dst)
    return "하루 로그(%s): 올림" % day, 글


# ── ③ 오류 내보내기 ────────────────────────────────────────────
def publish_errors(root=None, dry=False, days=30):
    root = root if root is not None else manual_root()
    if not root or not os.path.isdir(os.path.dirname(root)):
        return "오류 내보내기: 공유 폴더에 못 닿았다"
    dstdir = os.path.join(root, "03_오류")
    한 = []
    # 사전 — 정본은 reports/오류_사전.md 다. 사본만 올린다.
    사전 = os.path.join(BASE, "reports", "오류_사전.md")
    if os.path.isfile(사전):
        if not dry:
            os.makedirs(dstdir, exist_ok=True)
            shutil.copy2(사전, os.path.join(dstdir, "오류_사전.md"))
        한.append("사전")
    else:
        한.append("사전 없음(모름)")
    # 최근 기록 — error_book 에게 물어본다([162])
    try:
        import error_book as E
        res = E.rollup(days=days)
        # ★ 칸 이름을 짐작하지 않는다([165]). 2026-10-02 에 '목록'·'items' 로 물었다가
        #   **921건이 있는데 "0건"** 으로 적었다 — 오류도 안 나고 그럴듯해 보였다.
        #   error_book.rollup 이 실제로 주는 칸은 이것들이다.
        칸 = ("회귀", "새오류", "아는것", "못본것")
        없는칸 = [k for k in 칸 if k not in res]
        글 = ["# 최근 %d일 오류 — %s 기준" % (days, datetime.date.today().isoformat()), "",
             "*정본은 `ecount/reports/오류기록/YYYY-MM.jsonl` 이다. 이것은 읽기용 사본이다.*", "",
             "**모두 %s건 · 갈래 %s가지**" % (res.get("합계", "?"), res.get("갈래", "?")), ""]
        if 없는칸:
            글.append("> ⚠ 기대한 칸(%s)이 없다 — **덜 세었을 수 있다**(모름)."
                      % ", ".join(없는칸))
            글.append("")
        모두비었나 = True
        for k in 칸:
            v = res.get(k) or []
            if not v:
                continue
            모두비었나 = False
            글.append("## %s — %d가지" % (k, len(v)))
            글.append("")
            글.append("| 무엇 | 몇 번 | 사전 설명 | 마지막 |")
            글.append("|---|---:|---|---|")
            for it in v[:40]:
                글.append("| %s | %s | %s | %s |" % (
                    str(it.get("지문") or "?")[:70],
                    it.get("건수", "?"),
                    str(it.get("사전") or "(사전에 없음)")[:40],
                    str(it.get("마지막") or "")[:19]))
            글.append("")
        if 모두비었나:
            글.append("갈래별 목록이 비어 있다.")
            글.append("")
            글.append("> ⚠ '없다'와 '못 읽었다'는 다른 말이다. 합계가 0이 아닌데 목록이"
                      " 비어 있으면 **읽는 쪽이 틀린 것**이다.")
        if not dry:
            os.makedirs(dstdir, exist_ok=True)
            p = os.path.join(dstdir, "최근오류_%s.md" % datetime.date.today().isoformat())
            io.open(p, "w", encoding="utf-8").write("\n".join(글) + "\n")
        한.append("최근기록")
    except Exception as e:
        한.append("최근기록 못냄(%s)" % str(e)[:50])
    return "오류 내보내기: %s" % " · ".join(한)


def _print():
    root = manual_root()
    print("지정 폴더:", root or "(못 찾음)")
    print("  닿나:", os.path.isdir(root) if root else False)
    n = len(glob.glob(os.path.join(초안, "**", "*.md"), recursive=True))
    print("  올릴 초안:", n, "장  (%s)" % 초안)
    for 이름, p in LOG_SRC.items():
        print("  %-16s %s" % (이름, "있음 %.1f KB" % (os.path.getsize(p) / 1024.0)
                              if os.path.isfile(p) else "없음"))


def main(argv=None):
    ap = argparse.ArgumentParser(description="매뉴얼·로그·오류를 지정 폴더에 올린다")
    ap.add_argument("--print", dest="show", action="store_true")
    ap.add_argument("--dry", action="store_true")
    ap.add_argument("--root")
    ap.add_argument("--day", help="하루 로그를 만들 날짜(기본: 어제)")
    a = ap.parse_args(argv)
    if a.show:
        _print()
        return 0
    day = a.day or (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
    print(publish_manual(root=a.root, dry=a.dry))
    말, _ = day_log(day, root=a.root, dry=a.dry)
    print(말)
    print(publish_errors(root=a.root, dry=a.dry))
    return 0


if __name__ == "__main__":
    sys.exit(main())
