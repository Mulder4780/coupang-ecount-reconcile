# -*- coding: utf-8 -*-
"""pid_alive.py — "그 프로세스가 정말 살아 있나" 를 한 곳에서 판정한다

왜 따로 만들었나 (2026-08-06 실사고)
  `daily_run.py` 가 밤새 한 번도 안 돌았다. 이유는 `reports/.daily_run.lock` 이
  **죽은 프로세스의 이름으로 남아 있었기 때문**이다. 잠금은 "주인 pid 가 죽었으면
  회수한다" 는 옳은 규칙을 갖고 있었는데, 그 **판정이 틀렸다.**

  윈도우에서 `OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, …)` 는 **이미 끝난
  프로세스에도 핸들을 준다.** 프로세스 객체는 누군가 핸들을 쥐고 있는 동안 남아 있고,
  종료 여부는 `GetExitCodeProcess` 로 따로 물어야 한다(STILL_ACTIVE=259). 핸들이
  열렸다는 것만으로 '살아 있다' 고 본 탓에, 이미 끝난 pid 가 영원히 살아 있는 것으로
  나왔다. `Get-Process` 로는 안 보이는 pid 였다 — 그래서 사람 눈에도 안 띈다.
  잠금은 스스로 풀릴 길이 없어졌고, 그날 밤 대조는 통째로 비었다.

  덤으로 고친 것: `ctypes` 는 반환형을 지정하지 않으면 32비트 `int` 로 본다.
  64비트 윈도우의 핸들은 포인터라 **잘려서** 들어온다. 잘린 값으로 `CloseHandle` 을
  부르면 엉뚱한 핸들을 닫을 수도 있다. restype/argtypes 를 제대로 지정한다.

  모르면 '죽었다' 고 하지 않는다 — `None` 을 돌려 부르는 쪽이 시간 기준으로 넘기게 한다.
  살아 있는 옆 세션의 점유를 함부로 빼앗는 것이 그 반대 실수이며, 그쪽이 더 위험하다.

검증 [121].
"""
import os

STILL_ACTIVE = 259


def alive(pid):
    """살아 있으면 True · 확실히 죽었으면 False · 판정 불가면 None."""
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return None
    if pid <= 0:
        return None

    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes
            k = ctypes.windll.kernel32
            k.OpenProcess.restype = wintypes.HANDLE
            k.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
            k.CloseHandle.argtypes = [wintypes.HANDLE]
            k.GetExitCodeProcess.argtypes = [wintypes.HANDLE,
                                             ctypes.POINTER(wintypes.DWORD)]
            k.GetExitCodeProcess.restype = wintypes.BOOL

            h = k.OpenProcess(0x1000, False, pid)      # QUERY_LIMITED_INFORMATION
            if not h:
                # ERROR_ACCESS_DENIED(5) 는 '있는데 못 본다' 는 뜻이다 — 살아 있다.
                return True if ctypes.get_last_error() == 5 else False
            try:
                code = wintypes.DWORD()
                if not k.GetExitCodeProcess(h, ctypes.byref(code)):
                    return None
                # ★ 핸들이 열렸다고 살아 있는 것이 아니다. 끝난 프로세스도 핸들은 준다.
                return code.value == STILL_ACTIVE
            finally:
                k.CloseHandle(h)
        except Exception:
            return None

    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return True                # 내 것이 아니지만 있다
    except OSError:
        return None


def dead(pid):
    """확실히 죽었을 때만 True. 모르면 False — 산 것을 죽었다고 하지 않는다."""
    return alive(pid) is False
