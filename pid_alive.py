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

# 프로세스 생성시각 비교의 허용 오차(초). 잠금 파일을 쓰는 시각과 프로세스가 뜬
# 시각 사이에는 정상적으로도 몇 초의 틈이 있다 — 그 틈을 재사용으로 오판하지 않는다.
BORN_SLACK_S = 5.0
# Exact creation timestamps are serialized through JSON/ASCII floats.  Keep
# this transport tolerance far below the legacy five-second heuristic.
FINGERPRINT_TOLERANCE_S = 0.05


def started_at(pid):
    """프로세스 **생성 시각**(에포크 초). 모르면 None.

    ★ 왜 필요한가 (2026-08-11 실사고 — pid 재사용).
      09:50 회차가 '거래처코드 색인' 단계에서 죽었는데, 그 pid(37128)를 윈도우가
      `quick_share_server` 에 **재사용**했다. `alive(pid)` 는 '그 번호의 프로세스가
      있다'만 보므로 True 를 돌려줬고, 인계 문서는 다섯 시간 동안
      "지금 돌고 있다 — 기다려라"(정반대 지시)를 띄웠다. 잠금도 같은 판정을 쓰니
      스스로 풀리지 못했다 — 번호가 같다고 같은 프로세스가 아니다.
    """
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
            h = k.OpenProcess(0x1000, False, pid)   # QUERY_LIMITED_INFORMATION
            if not h:
                return None
            try:
                class FILETIME(ctypes.Structure):
                    _fields_ = [("lo", wintypes.DWORD), ("hi", wintypes.DWORD)]
                ct, xt, kt, ut = FILETIME(), FILETIME(), FILETIME(), FILETIME()
                k.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(FILETIME)] * 4
                k.GetProcessTimes.restype = wintypes.BOOL
                if not k.GetProcessTimes(h, ctypes.byref(ct), ctypes.byref(xt),
                                         ctypes.byref(kt), ctypes.byref(ut)):
                    return None
                # FILETIME = 1601-01-01 기준 100ns 단위 → 에포크 초
                ticks = (ct.hi << 32) | ct.lo
                return ticks / 10_000_000.0 - 11_644_473_600.0
            finally:
                k.CloseHandle(h)
        except Exception:
            return None
    try:
        return os.stat(f"/proc/{pid}").st_ctime
    except OSError:
        return None


def image_name(pid):
    """프로세스의 **실행파일 이름**(소문자, 예: 'python.exe'). 모르면 None.

    ★ 왜 필요한가 (2026-08-11 실사고 두 번째 — 검증 [211]).
      생성시각(born_before)은 '기록 시각보다 뒤에 태어난 남'만 거른다. 재사용된
      프로세스가 우연히 그 시각보다 **먼저** 떠 있던 것이면(오래 사는 상주 서비스가
      pid 를 물려받는 경우) 시각만으로는 못 가른다. 회차·잠금의 주인은 언제나
      python 이므로, 이름이 **읽히는데** python 이 아니면 번호만 같은 남이다.
      못 읽으면 None — 이름만으로 산 주인을 죽었다고 하지 않는다(alive 와 같은 원칙).
    """
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
            h = k.OpenProcess(0x1000, False, pid)   # QUERY_LIMITED_INFORMATION
            if not h:
                return None
            try:
                k.QueryFullProcessImageNameW.restype = wintypes.BOOL
                k.QueryFullProcessImageNameW.argtypes = [
                    wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR,
                    ctypes.POINTER(wintypes.DWORD)]
                buf = ctypes.create_unicode_buffer(1024)
                size = wintypes.DWORD(len(buf))
                if not k.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
                    return None
                return os.path.basename(buf.value).lower() or None
            finally:
                k.CloseHandle(h)
        except Exception:
            return None
    try:
        with open(f"/proc/{pid}/comm", encoding="utf-8", errors="replace") as fh:
            return fh.read().strip().lower() or None
    except OSError:
        return None


def alive(pid, born_before=None):
    """살아 있으면 True · 확실히 죽었으면 False · 판정 불가면 None.

    `born_before`(에포크 초)를 주면 **신원까지** 본다: 그 시각보다 **뒤에** 생긴
    프로세스라면 번호만 같은 남이므로 False. 원래 주인은 잠금·기록을 남기기 **전에**
    이미 떠 있었을 수밖에 없다 — 그래서 이 판정은 증명이지 짐작이 아니다.
    생성 시각을 모르면 예전과 똑같이 동작한다(산 것을 죽었다고 하지 않는다)."""
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return None
    if pid <= 0:
        return None
    res = _exists(pid)
    if res is True and born_before is not None:
        try:
            born = started_at(pid)
            if born is not None and born > float(born_before) + BORN_SLACK_S:
                return False    # 그 시각 뒤에 태어났다 — 번호만 같은 남이다
        except (TypeError, ValueError):
            pass
    return res


def _exists(pid):
    """번호 pid 의 프로세스가 지금 있나 — 신원은 보지 않는다."""
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
