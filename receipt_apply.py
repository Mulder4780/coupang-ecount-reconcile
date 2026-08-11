# -*- coding: utf-8 -*-
"""receipt_apply.py — 입금내역 한 파일을 **끝까지** 반영한다 (2026-08-11 지시)

사용자 지시: "김미영 입금내역 엑셀은 내가 주기적으로 세션에 업로드 할거야"
(16.Share 공유폴더 대기에서 **세션 업로드**로 창구 변경 — 엑셀 손입력 종료의 일부)

`kakao_apply.py` 와 같은 골격이다 — 조각은 이미 다 있고, 빠진 것은
사람이 파일을 건넸을 때 **한 줄로 끝까지 가는 길**이다:

  python receipt_apply.py <파일...> [--now]

  ① 내용 판별 — `inbox_scan.classify` 가 receipt 라고 해야만 간다.
     아니면 **exit 2 로 멈춘다** — 엉뚱한 파일을 '0건 반영 성공'으로
     끝내면 숫자도 나오고 오류도 안 나서 아무도 안 본다([169]).
  ② 정본 복사 — `7. 입금내역` 날짜폴더로 **복사**(원본 보존).
     같은 내용이면 이름이 달라도 다시 안 넣는다(SHA256).
  ③ 대조 — `receipt_fill --queue` (입금↔원장 양방향 유일 일치만 자동).
  ④ `--now`(사람 명령)일 때만 즉시 보관본 — 무인(`COUPANG_UNATTENDED=1`)은 거부.

★ csv 는 xlsx 로 바꿔 넣는다 — `receipt_fill.load_deposits` 는 `*.xlsx` 만
  읽는다(실측). csv 를 그대로 두면 **저장은 되는데 영영 안 읽힌다** —
  읽지 않는 파일은 빈칸과 같다([165]). xls(구형)는 변환기가 없으면
  거부하고 이유를 말한다 — 조용히 넣고 안 읽히는 것보다 낫다.

단계마다 `reports/입금_반영회차.json` 에 자국을 남긴다(죽어도 남는다).
"""
import json
import os
import subprocess
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

TRACE = os.path.join(ROOT, "reports", "입금_반영회차.json")


def note(step, state, extra=None):
    try:
        os.makedirs(os.path.dirname(TRACE), exist_ok=True)
        d = {"단계": step, "상태": state,
             "시각": datetime.now().astimezone().isoformat(timespec="seconds")}
        if extra:
            d.update(extra)
        prev = []
        try:
            prev = json.load(open(TRACE, encoding="utf-8"))
            if not isinstance(prev, list):
                prev = []
        except (OSError, ValueError):
            prev = []
        prev.append(d)
        tmp = TRACE + f".{os.getpid()}.tmp"
        json.dump(prev[-40:], open(tmp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        os.replace(tmp, TRACE)
    except OSError:
        pass


def _csv_to_xlsx(src):
    """csv → xlsx 변환본 경로. 원본 옆이 아니라 임시 폴더에 만든다."""
    import csv as _csv
    import tempfile
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    # 은행 csv 는 cp949 가 흔하다 — utf-8 먼저, 실패하면 cp949.
    for enc in ("utf-8-sig", "cp949"):
        try:
            with open(src, encoding=enc, newline="") as fh:
                for row in _csv.reader(fh):
                    ws.append(row)
            break
        except UnicodeDecodeError:
            ws.delete_rows(1, ws.max_row)
            continue
    out = os.path.join(tempfile.gettempdir(),
                       os.path.splitext(os.path.basename(src))[0] + ".xlsx")
    wb.save(out)
    return out


def main(argv):
    now_flag = "--now" in argv
    files = [a for a in argv if not a.startswith("--")]
    if now_flag and os.environ.get("COUPANG_UNATTENDED") == "1":
        print("! --now 는 사람 명령 전용이다 — 무인 경로에서는 거부한다")
        return 3
    if not files:
        print("사용법: python receipt_apply.py <입금내역 파일...> [--now]")
        return 2

    # ① 없는 파일은 조용히 건너뛰지 않는다 — 하나라도 없으면 멈춘다([186]의 규칙).
    missing = [f for f in files if not os.path.isfile(f)]
    if missing:
        note("확인", "실패", {"없는파일": missing})
        print("! 파일이 없다 — 멈춘다(0건 성공 금지):")
        for m in missing:
            print("  ·", m)
        return 2

    from inbox_scan import classify
    import source_dirs as S
    from source_organizer import dated_dir
    from upload_intake import _sha256

    copied = []
    for f in files:
        ext = os.path.splitext(f)[1].lower()
        src = f
        if ext == ".csv":
            note("변환", "시작", {"파일": os.path.basename(f)})
            src = _csv_to_xlsx(f)
            print(f"  csv → xlsx 변환: {os.path.basename(src)}")
        elif ext == ".xls":
            note("확인", "실패", {"파일": os.path.basename(f), "왜": "xls 미지원"})
            print(f"! {os.path.basename(f)}: 구형 .xls 는 읽는 쪽(receipt_fill)이 "
                  f"못 읽는다 — 엑셀에서 .xlsx 로 저장해 다시 달라")
            return 2
        elif ext != ".xlsx":
            note("확인", "실패", {"파일": os.path.basename(f), "왜": "지원 밖 확장자"})
            print(f"! {os.path.basename(f)}: 입금내역은 .xlsx/.csv 만 받는다")
            return 2

        kind = None
        try:
            kind = classify(src).get("kind")
        except Exception as e:
            note("판별", "실패", {"파일": os.path.basename(f), "오류": str(e)[:80]})
        if kind != "receipt":
            note("판별", "실패", {"파일": os.path.basename(f), "판별": kind})
            print(f"! {os.path.basename(f)}: 입금내역 모양이 아니다(판별={kind}) — "
                  f"멈춘다. 은행 양식이 바뀌었으면 inbox_scan 판별 규칙에 더해야 한다")
            return 2

        # ② 정본 복사 — 해시가 같으면 이미 들어온 것이다.
        dst_dir = dated_dir(S.RECEIPT_DIR, src)
        os.makedirs(dst_dir, exist_ok=True)
        h = _sha256(src)
        dup = None
        for base, _dirs, names in os.walk(S.RECEIPT_DIR):
            for n in names:
                if n.lower().endswith(".xlsx"):
                    p = os.path.join(base, n)
                    try:
                        if _sha256(p) == h:
                            dup = p
                            break
                    except OSError:
                        continue
            if dup:
                break
        if dup:
            note("복사", "중복", {"파일": os.path.basename(f), "기존": os.path.basename(dup)})
            print(f"  이미 들어온 내용({os.path.basename(dup)}) — 다시 안 넣는다")
            continue
        import shutil
        dst = os.path.join(dst_dir, os.path.basename(src))
        shutil.copy2(src, dst)
        copied.append(dst)
        note("복사", "끝", {"파일": os.path.basename(f), "정본": dst})
        print(f"  정본 복사: {dst}")

    # ③ 대조 — 새로 들어온 것이 없어도 돌린다(이전 회차가 못 끝냈을 수 있다).
    note("대조", "시작", {"신규": len(copied)})
    r = subprocess.run([sys.executable, os.path.join(ROOT, "receipt_fill.py"), "--queue"],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", cwd=ROOT)
    tail = (r.stdout or "").strip().splitlines()[-3:]
    for ln in tail:
        print(" ", ln)
    note("대조", "끝" if r.returncode == 0 else "실패", {"code": r.returncode})
    if r.returncode != 0:
        print("! 대조가 실패했다 — 보관본 단계로 넘어가지 않는다")
        return 1

    # ④ 사람 명령일 때만 즉시 보관본.
    if now_flag:
        note("보관본", "시작", None)
        r2 = subprocess.run([sys.executable, os.path.join(ROOT, "ledger_db.py"),
                             "--intake", "--apply", "--force", "--now"],
                            capture_output=True, text=True, encoding="utf-8",
                            errors="replace", cwd=ROOT)
        note("보관본", "끝" if r2.returncode == 0 else "실패", {"code": r2.returncode})
        print("  즉시 보관본:", "완료" if r2.returncode == 0 else f"실패({r2.returncode})")
    else:
        print("  보관본은 11:00·15:00 회차가 만든다 (--now 로 즉시)")
    note("회차", "끝", {"신규": len(copied)})
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
