# -*- coding: utf-8 -*-
"""관리대장 보관본을 **읽기 전용**으로 잠근다 (2026-08-28 지시).

형님 지시: "이제 엑셀은 저장용으로만 쓰고 모든건 앱으로 관리하게 알고리즘 변경해"

★ 왜 필요한가 — 말과 실제가 어긋나 있었다.
  `ledger_db` 는 화면에 `Excel역할: 단방향 보관본` 이라 적고 주석은
  "Excel 은 읽기 전용 보관본" 이라 적어 두는데, **파일은 누구나 고칠 수 있었다**
  (실측 2026-08-28: 최신 v622 쓰기가능=True). 그래서 2026-08-27 21:43 에
  실제로 손입력이 났고(v620→v621), 그 값은 역수입 금지라 정본에 못 들어가
  형님이 **앱에 같은 것을 한 번 더** 넣으셔야 했다.
  이 파일은 그 왕복을 **처음부터 안 생기게** 한다.

★ 지키는 것
  · **읽는 것은 그대로 된다.** 읽기 전용은 여는 것을 막지 않는다 —
    류지영·오종현이 원장을 열어 보는 길은 한 톨도 안 좁아진다([172]).
  · **못 걸어도 보관본 생성을 안 죽인다**([171] 정신) — 잠금 하나로 회차를
    세우지 않는다. 그러나 **못 걸었으면 말한다**([169]) — 조용히 넘어가면
    "잠갔다"는 거짓이 화면에 남는다.
  · **우리 코드가 옮기거나 지울 때는 먼저 푼다**(`unlock`). 윈도우는 읽기 전용
    파일의 **삭제·이동도 막는다** — 안 풀면 `ledger_versions` 의 옛 버전 정리가
    그날부터 조용히 죽는다(고치려다 더 나쁘게 만드는 자리 · [172]).
  · **되돌리기 한 줄**: 환경변수 `COUPANG_ARCHIVE_READONLY=0` ([126] 과 같은 보호장치).

⚠ 이것은 담을 두르는 것이지 자물쇠가 아니다. 엑셀에서 "다른 이름으로 저장" 하면
  우회된다 — 그러나 그것은 **다른 파일**이라 원장이 더럽혀지지 않고,
  그래도 최신본을 갈아치우면 손입력 감지(`realtime_monitor`)가 잡는다.
"""
import os
import stat
import sys

ENV = "COUPANG_ARCHIVE_READONLY"


def enabled():
    """켜져 있나. 기본은 켬 — 끄려면 `COUPANG_ARCHIVE_READONLY=0`."""
    return (os.environ.get(ENV) or "").strip() not in ("0", "false", "off", "no")


def is_locked(path):
    """읽기 전용인가. **못 읽으면 None**(모름) — False 로 뭉개지 않는다([169])."""
    try:
        return not bool(os.stat(path).st_mode & stat.S_IWRITE)
    except OSError:
        return None


def in_use(path):
    """지금 사람이 쓰는 중인가. 판정은 `ledger_versions._in_use` 를 **빌린다**([162]) —
    `~$` 잠금파일과 최근 저장 시각을 본다. 못 빌리면 **True**(안 잠근다) 쪽으로
    기운다: 쓰는 중인 파일을 잠그면 그 사람의 입력이 날아간다([104] 류지영 우선)."""
    try:
        import ledger_versions
        return bool(ledger_versions._in_use(path))
    except Exception:
        return True


def lock(path, check_in_use=False):
    """보관본을 읽기 전용으로 만든다. 돌려주는 값은 (했나, 왜).

    끄면 `(False, "꺼짐")` · 이미 걸려 있으면 `(True, "이미")` ·
    못 걸면 `(False, 이유)` — **예외를 올리지 않는다**(회차를 안 죽인다).

    ★ `check_in_use` 는 **사람이 손으로 잠글 때만** 켠다. 회차가 방금 만든 새
    정본은 당연히 "방금 저장됨" 이라 그 문을 켜면 **영영 못 잠근다**."""
    if not enabled():
        return False, "꺼짐(%s=0)" % ENV
    if check_in_use and in_use(path):
        return False, "지금 누가 쓰는 중이다(~$ 잠금 또는 최근 저장) — 안 잠갔다"
    try:
        mode = os.stat(path).st_mode
    except OSError as exc:
        return False, "못읽음: %s" % exc
    if not (mode & stat.S_IWRITE):
        return True, "이미"
    try:
        os.chmod(path, mode & ~stat.S_IWRITE)
    except OSError as exc:
        return False, "못걸었다: %s" % exc
    got = is_locked(path)
    if got is not True:
        # 걸었다는데 안 걸렸다 — 성공이라 적지 않는다([169]).
        return False, "걸었는데 안 걸렸다(공유폴더 권한일 수 있다)"
    return True, "걸었다"


def unlock(path):
    """우리 코드가 옮기거나 지우기 **직전**에 푼다. (했나, 왜).

    ⚠ 사람에게 쓰라고 여는 것이 아니다 — 옮긴 뒤 그 파일은 OLD 로 가고,
    새 최신본은 다시 잠긴다."""
    try:
        mode = os.stat(path).st_mode
    except OSError as exc:
        return False, "못읽음: %s" % exc
    if mode & stat.S_IWRITE:
        return True, "원래 열려 있었다"
    try:
        os.chmod(path, mode | stat.S_IWRITE)
    except OSError as exc:
        return False, "못풀었다: %s" % exc
    return True, "풀었다"


def _say(msg):
    if hasattr(sys.stdout, "write"):
        try:
            print(msg)
        except Exception:
            pass


def main(argv=None):
    """사람이 보는 자리 — 인자 없이 부르면 지금 상태만 말한다(아무것도 안 바꾼다)."""
    import argparse
    import workbook_patch

    ap = argparse.ArgumentParser(description="관리대장 보관본 읽기 전용 잠금")
    ap.add_argument("--lock", action="store_true", help="최신 보관본을 잠근다")
    ap.add_argument("--unlock", action="store_true", help="최신 보관본을 푼다(사람 판단)")
    a = ap.parse_args(argv)

    try:
        path, ver = workbook_patch.latest_master()
    except SystemExit as exc:
        _say("관리대장을 못 찾았습니다: %s" % exc)
        return 2

    name = os.path.basename(path)
    if a.lock:
        ok, why = lock(path, check_in_use=True)
        _say("%s v%s -> %s (%s)" % ("잠갔습니다" if ok else "못 잠갔습니다", ver, name, why))
        return 0 if ok else 1
    if a.unlock:
        ok, why = unlock(path)
        _say("%s v%s -> %s (%s)" % ("풀었습니다" if ok else "못 풀었습니다", ver, name, why))
        return 0 if ok else 1

    st = is_locked(path)
    _say("최신 보관본 v%s · %s" % (ver, name))
    if st is None:
        _say("  상태: 확인 못 함 (파일을 못 읽었습니다)")
    elif st:
        _say("  상태: 읽기 전용 — 저장용입니다. 값은 앱에서 고칩니다.")
    else:
        _say("  상태: 쓰기 가능 — 손으로 고칠 수 있습니다(다음 보관본부터 잠깁니다).")
    _say("  기능: %s" % ("켜짐" if enabled() else "꺼짐(%s=0)" % ENV))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
