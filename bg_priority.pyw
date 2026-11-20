# -*- coding: utf-8 -*-
"""
bg_priority.pyw : 백그라운드 작업이 마우스·화면을 못 뺏게 우선순위를 낮춘다 (2026-11-21 형님 지시)

  형님 지시: "마우스 드래그하면 맨날 뚝뚝 끊겨서 캡처하기 불편 · 팝업 띄우지 말고 전부 백그라운드에서"

실측 2026-11-21: CPU 100%. 범인은 창 없는 자동 작업들(헤드리스 크롬 16개 · node · python)이
'보통' 우선순위로 돌며 마우스와 같은 줄에 서 있던 것이다.
  · 끄지 않는다(남의 세션 일이다). **줄만 뒤로 세운다**(Idle/BelowNormal) : 일은 그대로 끝난다.
    사람이 마우스를 움직이면 윈도우가 그쪽을 먼저 처리한다.
  · 창이 있는 프로그램(웨일·크롬 본창·엑셀)은 안 건드린다 : 형님이 쓰는 것이다.
  · 이 파일은 pythonw 로 돈다 : 콘솔 창이 한 번도 안 뜬다.
작업 스케줄러 'CSOS_백그라운드_우선순위' 가 1분마다 부른다(숨김).
되돌리기: 그 예약 작업을 끄면 된다(낮춘 우선순위는 그 프로세스가 끝나면 사라진다).
"""
import subprocess

NAMES = ("chrome-headless-shell", "node", "python", "pythonw", "ffmpeg", "git", "rg")
PS = r"""
$n=@(%s)
Get-Process -Name $n -EA 0 | ? { $_.MainWindowHandle -eq 0 -and $_.PriorityClass -and $_.PriorityClass -ne 'Idle' -and $_.PriorityClass -ne 'BelowNormal' } | ForEach-Object {
  try { if($_.Name -eq 'chrome-headless-shell'){ $_.PriorityClass='Idle' } else { $_.PriorityClass='BelowNormal' } } catch {}
}
""" % ",".join("'%s'" % n for n in NAMES)

subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", PS],
               creationflags=0x08000000, stdin=subprocess.DEVNULL,
               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
