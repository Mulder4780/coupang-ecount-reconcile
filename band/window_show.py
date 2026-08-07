# -*- coding: utf-8 -*-
"""
window_show.py — 수집용 크롬 창을 **초점은 그대로 둔 채** 화면에 되살린다 (2026-08-07)

무엇을 풀려는 것인가
  밴드 수집기는 숨은 탭에서 시작을 거절한다 — 밴드 본문이 보이는 탭에서만 그려지기
  때문이다. 그런데 크롬 창이 **최소화**돼 있으면 탭이 그 창의 활성 탭이어도
  `document.hidden` 이 참이라 수집이 통째로 막힌다(2026-08-07 실측: 창 하나, 밴드 탭이
  활성, 그런데도 hidden=true → `__grabStart` 가 "탭이 뒤에 있다"로 거절).

왜 그냥 앞으로 꺼내지 않나
  사용자가 **두 창을 나란히** 굴리고 있다(이 창은 수집, 옆 창은 CSOS 앱 코딩+엑셀).
  수집할 때마다 크롬이 초점을 빼앗으면 옆 창에서 타이핑하던 것이 끊긴다 —
  병렬로 일하게 하려고 만든 판을 수집이 매번 걷어차는 꼴이다.
  그래서 `SW_SHOWNOACTIVATE`(4) 를 쓴다: **보이게는 하되 활성화하지 않는다.**
  키보드 초점은 사용자가 쓰던 창에 그대로 남는다.

  python band/window_show.py                 # 지금 상태만
  python band/window_show.py --apply         # 최소화된 크롬 창을 초점 없이 복원
  python band/window_show.py --apply --title 밴드

★ `--apply` 로도 안 풀릴 때가 있다 — **덮여 있음**(2026-08-07 실측, 반나절 손해)
  창이 최소화도 아니고 Win32 로는 '보임'인데 `document.hidden` 이 참인 경우가 있다.
  다른 창이 **완전히 덮고 있으면** 크롬이 그 창을 가려진 것으로 보고(native window
  occlusion) 탭 렌더링을 멈춘다. 이때:
    · `ShowWindow` 는 못 푼다 — 최소화가 아니니 할 일이 없다.
    · `SetWindowPos(HWND_TOP)` 도 **못 푼다** — Z순서 맨 위로 올려도 항상 위는 아니라서
      덮고 있던 창(작업표시줄 위 앱·항상 위 창)이 그대로 다시 덮는다.
    · `SetWindowPos(HWND_TOPMOST)` 는 **푼다.** '항상 위'로 잠깐 고정하면 아무도 못 덮는다.
  그래서 수집 동안만 고정하고 **반드시 되돌린다**(안 되돌리면 크롬이 영영 항상 위에
  떠서 사용자가 다른 창을 못 쓴다 — 이게 이 기능의 유일한 위험이다).

  python band/window_show.py --topmost       # 항상 위로 고정(수집 직전)
  python band/window_show.py --untopmost     # 되돌리기(수집 직후) — 잊지 말 것

  파이썬에서는 `with pinned():` 를 쓴다. 예외가 나도 복귀가 보장된다:
      from band.window_show import pinned
      with pinned():
          ...수집...
"""
import argparse
import contextlib
import ctypes
import sys
from ctypes import wintypes

SW_SHOWNOACTIVATE = 4          # 보이게 하되 활성화하지 않는다 (초점 유지)
HWND_TOP = 0
HWND_TOPMOST = -1              # 항상 위 — '덮여 있음'을 푸는 유일한 값
HWND_NOTOPMOST = -2            # 보통 창으로 복귀
SWP_NOMOVE, SWP_NOSIZE, SWP_NOACTIVATE = 0x0002, 0x0001, 0x0010

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _u32():
    if not hasattr(ctypes, "windll"):
        return None
    return ctypes.windll.user32


def windows(match=""):
    """보이는 최상위 창 중 제목에 `match` 가 든 것 → [(hwnd, 제목, 최소화여부)]."""
    u = _u32()
    if u is None:
        return []
    out = []
    cb_t = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def cb(hwnd, _lp):
        if not u.IsWindowVisible(hwnd):
            return True
        n = u.GetWindowTextLengthW(hwnd)
        if n <= 0:
            return True
        buf = ctypes.create_unicode_buffer(n + 1)
        u.GetWindowTextW(hwnd, buf, n + 1)
        title = buf.value
        if match and match not in title:
            return True
        out.append((hwnd, title, bool(u.IsIconic(hwnd))))
        return True

    u.EnumWindows(cb_t(cb), 0)
    return out


def _zorder(hwnd, after):
    """Z순서만 바꾼다 — 옮기지도, 크기를 바꾸지도, **활성화하지도** 않는다."""
    u = _u32()
    if u is None:
        return False
    return bool(u.SetWindowPos(wintypes.HWND(hwnd), wintypes.HWND(after), 0, 0, 0, 0,
                               SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE))


def set_topmost(on=True, title="Chrome", quiet=False):
    """크롬 창을 '항상 위'로 고정하거나 되돌린다 → 바뀐 창 수.

    `on=True` 는 **수집하는 동안만** 쓴다. 켜 둔 채로 두면 사용자가 다른 창을
    앞으로 못 꺼낸다 — 반드시 `set_topmost(False)` 나 `pinned()` 로 되돌릴 것.
    """
    found = windows(title)
    n = 0
    for hwnd, name, minimized in found:
        if minimized:
            # 최소화된 창은 항상 위로 만들어도 안 보인다. 먼저 초점 없이 되살린다.
            _u32().ShowWindow(hwnd, SW_SHOWNOACTIVATE)
        if _zorder(hwnd, HWND_TOPMOST if on else HWND_NOTOPMOST):
            n += 1
            if not quiet:
                print(f"  {'항상위' if on else '복귀 '} {name[:60]}")
    if not quiet and not found:
        print(f"✗ 제목에 '{title}' 이 든 창을 못 찾았다")
    return n


@contextlib.contextmanager
def pinned(title="Chrome", quiet=True):
    """수집하는 동안만 '항상 위'. 예외가 나도 **반드시** 되돌린다.

    되돌리기를 `finally` 에 두는 것이 이 함수의 전부다. 수집기는 중간에 자주
    죽는데(밴드가 로그인 화면을 주거나 번호가 없거나), 그때 크롬이 항상 위에
    남으면 사용자는 원인을 모른 채 창이 안 내려간다고 겪는다.
    """
    n = set_topmost(True, title, quiet=quiet)
    try:
        yield n
    finally:
        set_topmost(False, title, quiet=quiet)


def run(apply=False, title="Chrome"):
    found = windows(title)
    if not found:
        print(f"✗ 제목에 '{title}' 이 든 창을 못 찾았다 — 크롬이 떠 있는지 확인할 것")
        return 1
    u, changed = _u32(), 0
    for hwnd, name, minimized in found:
        state = "최소화됨" if minimized else "보임"
        print(f"  [{state}] {name[:60]}")
        if not apply:
            continue
        # 초점은 절대 옮기지 않는다 — 옆 세션에서 타이핑 중일 수 있다.
        if minimized:
            u.ShowWindow(hwnd, SW_SHOWNOACTIVATE)
        else:
            # ★ '보임'인데도 hidden 인 경우가 있다 (2026-08-07 실측).
            #   크롬은 **다른 창에 완전히 가려진** 창의 탭도 hidden 으로 본다
            #   (native window occlusion). 최소화가 아니므로 ShowWindow 로는 안 풀린다.
            #   Z순서만 위로 올리고 활성화는 하지 않는다 — 키보드 초점은 그대로다.
            #   ※ 이것으로도 안 풀리면 `--topmost`(HWND_TOPMOST) 다. HWND_TOP 은
            #     '맨 위'일 뿐 '항상 위'가 아니라서 덮던 창이 곧 다시 덮는다.
            u.SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0,
                           SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE)
        changed += 1
    if not apply:
        print("  (바꾸지 않음 — 실제로 되살리려면 --apply)")
    else:
        print(f"→ {changed}개 창을 초점 없이 앞으로 올렸다. 이제 수집기가 시작할 수 있다.")
        print("  (키보드 초점은 쓰던 창에 그대로 — 옆 세션 작업이 끊기지 않는다)")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="수집용 크롬 창을 초점 없이 되살린다")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--title", default="Chrome")
    ap.add_argument("--topmost", action="store_true",
                    help="'덮여 있음'을 푼다 — 항상 위로 고정(수집 직전)")
    ap.add_argument("--untopmost", action="store_true",
                    help="항상 위 해제(수집 직후) — 잊으면 창이 안 내려간다")
    a = ap.parse_args(argv)
    if a.topmost or a.untopmost:
        on = bool(a.topmost)
        n = set_topmost(on, a.title)
        if n and on:
            print(f"→ {n}개 창을 항상 위로 고정했다. **수집이 끝나면 --untopmost** 를 부를 것.")
        elif n:
            print(f"→ {n}개 창을 보통 창으로 되돌렸다.")
        return 0 if n else 1
    return run(a.apply, a.title)


if __name__ == "__main__":
    sys.exit(main())
