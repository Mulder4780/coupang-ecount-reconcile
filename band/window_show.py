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
"""
import argparse
import ctypes
import sys
from ctypes import wintypes

SW_SHOWNOACTIVATE = 4          # 보이게 하되 활성화하지 않는다 (초점 유지)

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


def run(apply=False, title="Chrome"):
    found = windows(title)
    if not found:
        print(f"✗ 제목에 '{title}' 이 든 창을 못 찾았다 — 크롬이 떠 있는지 확인할 것")
        return 1
    changed = 0
    for hwnd, name, minimized in found:
        state = "최소화됨" if minimized else "보임"
        print(f"  [{state}] {name[:60]}")
        if minimized and apply:
            # 초점을 옮기지 않는다 — 옆 세션에서 타이핑 중일 수 있다.
            _u32().ShowWindow(hwnd, SW_SHOWNOACTIVATE)
            changed += 1
    if not apply:
        print("  (바꾸지 않음 — 실제로 되살리려면 --apply)")
    elif changed:
        print(f"→ {changed}개 창을 초점 없이 되살렸다. 이제 수집기가 시작할 수 있다.")
    else:
        print("→ 최소화된 창이 없다. 그래도 hidden 이면 탭이 그 창의 활성 탭이 아니다.")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="수집용 크롬 창을 초점 없이 되살린다")
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--title", default="Chrome")
    a = ap.parse_args(argv)
    return run(a.apply, a.title)


if __name__ == "__main__":
    sys.exit(main())
