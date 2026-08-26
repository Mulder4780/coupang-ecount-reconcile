# -*- coding: utf-8 -*-
"""kakao_apply.py — 카톡 내보내기 한 파일을 **끝까지** 반영하는 회차 (2026-08-09 지시)

사용자 지시: "이거 반영하고 엑셀 및 앱에 반영해
              (내가 반영하라고 명령하면 규칙 무시하고 반영하는 알고리즘 추가)"

★ 빠져 있던 것은 '반영하는 기능'이 아니라 **한 줄로 끝까지 가는 길**이었다.
  조각은 다 있었다 — 흡수는 `download_intake`, 추출은 `kakao_extract --new --queue`,
  엑셀 쓰기는 `ledger_db --apply`. 그런데 사람이 카톡 파일을 건네면 그 셋을 **손으로
  이어야** 했고, 이어지는 도중 어디서 끊겼는지는 아무 화면에도 안 나왔다.
  그래서 "반영해 달라"는 말이 매번 사람(또는 클로드) 손을 탔다.

무엇을 하는가 (네 단계, 각 단계가 자국을 남긴다)
  ① 파일 확인 → ② 정본 폴더로 **복사** → ③ 등록 후보 추출 + 큐 적재 → ④ 엑셀 반영

★ 없는 파일을 조용히 건너뛰지 않는다. 이 회차에서 제일 위험한 것은
  "파일이 없었는데 나머지 단계가 멀쩡히 돌아 **0건 반영 성공**"으로 끝나는 것이다.
  숫자도 나오고 오류도 안 나서 아무도 안 본다([169] 와 같은 모양).
  그래서 **주어진 경로 중 하나라도 없으면 exit 2** 로 멈춘다.

★ 원본은 **옮기지 않고 복사**한다. 사람이 바탕화면에 둔 파일을 지우면
  "그때 정말 무엇을 받았나"를 사람 쪽에서 잃는다. 같은 내용이 이미 정본에 있으면
  이름이 달라도 다시 넣지 않는다(해시로 판정 — `kakao_extract.source_paths` 와 같은 근거).

★ 엑셀 반영은 **사람이 명령했을 때만** 지금 한다(`--now`).
  규칙이 뒤집힌 것이 아니다 — 막으려던 것은 *도구가 채울 때마다 저절로* vN+1 이
  쏟아지는 것이었지, 사람이 스스로 내린 명령이 아니다(2026-08-07 지시 · 검증 `[93]`).
  · `--now` 없이 `--apply` 만 주면 11:00·15:00 회차를 기다린다(평소 규칙 그대로).
  · 무인 경로(`COUPANG_UNATTENDED=1`)가 `--now` 를 부르면 **거부한다.**
    daily_run·session_wrapup 이 이 파일을 즉시반영으로 부르지 않는지는 검증이 지킨다.

실행
  python kakao_apply.py <파일...> --now      # 형님이 준 파일을 지금 끝까지 반영
  python kakao_apply.py --now                # 새 파일 없이, 이미 있는 카톡 자료로 반영
  python kakao_apply.py --find               # 최근 카톡 내보내기가 어디 있는지만 찾아본다
  python kakao_apply.py <파일...>            # 큐까지만(엑셀은 11:00·15:00 회차 몫)
"""
import glob
import hashlib
import io
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

REPORT = os.path.join(ROOT, "reports", "카톡_반영회차.json")

# 사람이 카톡 파일을 떨어뜨리는 자리들. **찾기용일 뿐** 여기서 지우지 않는다.
def drop_dirs():
    home = os.path.expanduser("~")
    out = [os.path.join(home, "Desktop"), os.path.join(home, "Downloads"),
           os.path.join(home, "OneDrive", "바탕 화면"),
           os.path.join(home, "OneDrive", "Desktop")]
    return [d for d in out if os.path.isdir(d)]


def canon_dir():
    """정본 자리 — 없으면 로컬 inbox 로 물러선다(Z: 가 안 붙는 PC 도 있다)."""
    try:
        import source_dirs as S
        d = os.path.join(S.KAKAO_DIR, "2026")
        if os.path.isdir(S.KAKAO_DIR):
            os.makedirs(d, exist_ok=True)
            return d
    except Exception:
        pass
    d = os.path.join(ROOT, "kakao", "inbox")
    os.makedirs(d, exist_ok=True)
    return d


def sha(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def known_hashes(folder):
    """정본 폴더가 이미 가진 내용들. 이름이 달라도 같은 내용은 다시 안 넣는다."""
    out = {}
    for p in glob.glob(os.path.join(folder, "**", "*.txt"), recursive=True):
        try:
            out[sha(p)] = p
        except OSError:
            continue
    return out


def find_recent(hours=48):
    """최근 카톡 내보내기를 사람이 떨구는 자리에서 찾는다(읽기만)."""
    cut = time.time() - hours * 3600
    hits = []
    for d in drop_dirs():
        for p in glob.glob(os.path.join(d, "KakaoTalk_*.txt")):
            try:
                m = os.path.getmtime(p)
            except OSError:
                continue
            if m >= cut:
                hits.append((m, p))
    hits.sort(reverse=True)
    return [p for _, p in hits]


def absorb(paths, folder):
    """정본 폴더로 **복사**. (넣음, 이미있음, 실패) 를 돌려준다."""
    have = known_hashes(folder)
    put, dup, fail = [], [], []
    for src in paths:
        try:
            d = sha(src)
        except OSError as e:
            fail.append((src, str(e)))
            continue
        if d in have:
            dup.append((src, have[d]))
            continue
        base = os.path.basename(src)
        dst = os.path.join(folder, base)
        n = 1
        while os.path.exists(dst):          # 이름은 같은데 내용이 다른 것 — 둘 다 남긴다
            stem, ext = os.path.splitext(base)
            dst = os.path.join(folder, "%s__ka%d%s" % (stem, n, ext))
            n += 1
        try:
            shutil.copy2(src, dst)
        except OSError as e:
            fail.append((src, str(e)))
            continue
        have[d] = dst
        put.append((src, dst))
    return put, dup, fail


def run(args, minutes=20):
    """자식은 Popen + communicate(timeout=) 로만 돌린다.
    편의 함수 쪽 timeout 은 윈도우에서 영원히 매달릴 수 있다([175])."""
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    p = subprocess.Popen([sys.executable] + args, cwd=ROOT, env=env,
                         stdout=subprocess.PIPE, stderr=subprocess.STDOUT, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    try:
        out, _ = p.communicate(timeout=minutes * 60)
        return p.returncode, (out or b"").decode("utf-8", "replace")
    except subprocess.TimeoutExpired:
        subprocess.call(["taskkill", "/F", "/T", "/PID", str(p.pid)],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        try:
            out, _ = p.communicate(timeout=30)
        except Exception:
            out = b""
        return 124, (out or b"").decode("utf-8", "replace") + "\n(시간 초과로 끊음)"


KNOWN_FLAGS = ("--now", "--apply", "--find", "--help", "-h")


def unknown_flags(argv):
    """모르는 깃발만 골라 돌려준다 (2026-08-19).

    ★ **조용히 무시하면 오타가 성공처럼 끝난다.** `main` 은 `--` 로 시작하는 것을
      통째로 버려 `given` 을 만들었다 — 그래서 `--nwo` 라고 잘못 치면 엑셀 반영을
      시켰는데 **큐까지만 하고 "완료"** 로 끝난다. 오류도 안 나고 숫자도 나와서
      아무도 안 본다([169] 와 같은 모양). 실측으로 `--help` 가 그렇게 **회차를
      통째로 돌렸다.**

    ⚠ `-` 하나로 시작하는 것도 깃발로 본다 — `-now` 를 파일 이름으로 읽으면
      "파일을 다시 내려받으십시오" 라는 **엉뚱한 안내**가 나간다([172]).
      경로가 정말 `-` 로 시작하면 `./` 를 붙인다.
    """
    return [a for a in argv if a.startswith("-") and a not in KNOWN_FLAGS]


def main(argv):
    now_flag = "--now" in argv
    apply_flag = "--apply" in argv or now_flag
    given = [a for a in argv if not a.startswith("-")]

    # ★ 모르는 깃발이면 **아무것도 하기 전에** 멈춘다(위 `unknown_flags` 설명).
    bad = unknown_flags(argv)
    if bad:
        print("멈춤: 모르는 깃발입니다 — %s" % " ".join(bad))
        print("   쓸 수 있는 것: %s" % " · ".join(KNOWN_FLAGS))
        print("   (파일 경로가 `-` 로 시작하면 `./` 를 붙여 주십시오)")
        return 2

    if "--help" in argv or "-h" in argv:
        print((__doc__ or "").strip())
        return 0

    if "--find" in argv:
        hits = find_recent()
        print("최근 48시간 카톡 내보내기 %d개" % len(hits))
        for p in hits:
            print("  ", time.strftime("%m-%d %H:%M", time.localtime(os.path.getmtime(p))),
                  p, os.path.getsize(p))
        return 0

    # ★ 무인 경로는 즉시반영을 못 쓴다. 사람 명령이라야 규칙을 넘는다.
    if now_flag and os.environ.get("COUPANG_UNATTENDED") == "1":
        print("거부: 무인 실행에서는 즉시 엑셀 반영을 하지 않습니다 "
              "— 11:00·15:00 회차로 들어갑니다 (검증 [93])")
        return 3

    step = []

    # ① 파일 확인 — 없는 것을 조용히 넘기지 않는다.
    missing = [p for p in given if not os.path.isfile(p)]
    if missing:
        print("멈춤: 주신 파일을 찾지 못했습니다 (%d개)" % len(missing))
        for p in missing:
            print("   없음:", p)
        hits = find_recent()
        if hits:
            print("\n최근 48시간 안에 이런 것들은 있습니다 — 이 중 하나입니까?")
            for p in hits[:6]:
                print("   ", p)
        print("\n파일을 다시 내려받으신 뒤 같은 명령을 주시거나, "
              "`0. 원본 자료/100. 업로드용 자료` 에 넣어 주십시오.")
        return 2

    # ② 정본 폴더로 복사
    folder = canon_dir()
    put = dup = fail = []
    if given:
        put, dup, fail = absorb(given, folder)
        print("① 원본 흡수 — 새로 넣음 %d · 이미 있음 %d · 실패 %d"
              % (len(put), len(dup), len(fail)))
        for s, d in put:
            print("   +", os.path.basename(d))
        for s, d in dup:
            print("   = 같은 내용이 이미 있음:", os.path.basename(d))
        for s, e in fail:
            print("   ! 실패:", s, e)
        if fail and not put:
            print("멈춤: 파일을 하나도 못 넣었습니다.")
            return 2
    else:
        print("① 새 파일 없음 — 이미 있는 카톡 자료로 진행합니다")
    step.append({"단계": "원본 흡수", "넣음": len(put), "이미있음": len(dup), "실패": len(fail)})

    # ③ 등록 후보 추출 + 큐 적재
    rc, out = run([os.path.join(ROOT, "kakao_extract.py"), "--new", "--queue"], minutes=25)
    tail = "\n".join([l for l in out.splitlines() if l.strip()][-12:])
    print("\n② 등록 후보 추출·큐 적재 (%s)" % ("성공" if rc == 0 else "실패 rc=%d" % rc))
    print(tail)
    held = _held_lines(out)
    step.append({"단계": "추출·큐", "rc": rc, "보류": held, "요약": tail[-800:]})
    if rc != 0:
        print("멈춤: 추출이 실패해 엑셀로 넘기지 않습니다 — "
              "실패한 채로 반영하면 없는 것이 반영된 것처럼 보입니다.")
        _save(step, given, now_flag)
        return 1

    # ④ 엑셀 반영
    if not apply_flag:
        print("\n③ 엑셀 반영은 하지 않았습니다 — 11:00·15:00 회차가 가져갑니다.")
        print("   지금 넣으시려면 같은 명령에 --now 를 붙이십시오.")
        _say_held(held)
        _save(step, given, now_flag)
        return 0

    args = [os.path.join(ROOT, "ledger_db.py"), "--intake", "--apply"]
    if now_flag:
        args += ["--force", "--now"]        # 사람 명령 — 시각을 무시하고 지금 쓴다
    rc2, out2 = run(args, minutes=30)
    tail2 = "\n".join([l for l in out2.splitlines() if l.strip()][-14:])
    print("\n③ 엑셀 반영 (%s)" % ("성공" if rc2 == 0 else "실패 rc=%d" % rc2))
    print(tail2)
    step.append({"단계": "엑셀 반영", "rc": rc2, "즉시": now_flag, "요약": tail2[-1200:]})
    _say_held(held)
    _save(step, given, now_flag)
    return 0 if rc2 == 0 else 1


# ② 추출 출력에서 '[보류]' 줄만 뽑는다 — **출력 전체**에서 센다.
#   tail(마지막 14줄)이 아니라 out 전체를 본다. 보류가 많으면 tail 밖으로
#   밀려나고, 그러면 이 계기 자신이 눈이 먼다([169]).
#   낱말 '[보류]' 는 kakao_extract 가 찍는 그 글자다([162]) — 어긋나면
#   한 건도 안 걸리면서 오류도 안 난다([165]). 검증이 그 계약을 얼린다.
def _held_lines(out):
    return [l.strip() for l in (out or '').splitlines() if '[보류]' in l]


# 보류를 **마지막에 다시** 적는다 — 끝 줄이 거짓말하지 않게.
#   왜 다시 적나: ② 에서 이미 찍었지만 그 뒤 '③ 엑셀 반영 (성공)' 이 오면
#   사람은 마지막 줄만 읽는다. 실측 2026-08-26 — 형님이 주신 카톡에서 새 접수
#   4건을 정확히 뽑고도 02_돌발AS접수 여유가 2행뿐이라 전량 보류였는데,
#   끝 줄은 '성공' 이고 exit 0 이었다([169]: 숫자도 나오고 오류도 안 난다).
#   ★ exit 코드는 **안 바꾼다**([172]). 부르는 곳이 automation_pipeline 하나인데
#   rc!=0 이면 그 갈래가 실패로 적혀 **지문이 안 커밋되고 같은 파일을 무한
#   재처리한다.** 보류는 파일 문제가 아니라 시트 여유행 문제라 다시 처리해도 같다.
#   대신 자국(reports/카톡_반영회차.json)과 인계 문서가 말한다.
def _say_held(held):
    if not held:
        return
    print()
    print('★ 보류 %d건 — 그만큼은 원장에 **안 들어갔습니다**(반쯤 넣는 것보다 낫습니다).'
          % len(held))
    for h in held[:6]:
        print('   ', h)
    if len(held) > 6:
        print('    ... 그 밖 %d건' % (len(held) - 6))
    print('   여유행이 모자란 것이면 expand_rows 로 늘린 뒤 다시 부릅니다 —')
    print('   ★ 자동으로 늘리지 않습니다: 관리대장 구조를 바꾸는 일은 사람이 정합니다.')


def _save(step, given, now_flag):
    rec = {"시각": datetime.now().isoformat(timespec="seconds"),
           "받은파일": [os.path.basename(p) for p in given],
           "즉시반영": now_flag, "단계": step}
    try:
        os.makedirs(os.path.dirname(REPORT), exist_ok=True)
        old = []
        if os.path.exists(REPORT):
            try:
                old = json.load(io.open(REPORT, encoding="utf-8"))
            except Exception:
                old = []
        old.insert(0, rec)
        with io.open(REPORT, "w", encoding="utf-8", newline="") as f:
            json.dump(old[:60], f, ensure_ascii=False, indent=1)
        print("\n자국: reports/카톡_반영회차.json")
    except Exception as e:
        print("(기록 실패:", e, ")")


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
