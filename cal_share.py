# -*- coding: utf-8 -*-
"""
cal_share.py — 쿠팡 캘린더만 **링크 하나로** 열리게 만든다 (고정 주소용)

사용자 지시(2026-08-06): "공유된 쿠팡 캘린더는 핀 번호 입력 기능 없애고 바로 볼 수
있게 정리해." · 공개 범위는 **링크 아는 사람만**으로 정했다.

왜 앱 PIN 을 링크에 담지 않나 — 그 PIN 하나면 정산·계산서·확인필요까지 다 열린다.
단톡방에 뿌린 링크로 회사 원장 전체가 열리는 셈이다. 그래서 **캘린더 전용 자물쇠**를
따로 만든다. 링크가 새어도 새는 것은 일정뿐이다.

어떻게 되나
  · 캘린더만 담은 꾸러미를 **무작위 열쇠**로 잠가 `docs/cal.enc` 로 올린다.
  · 열쇠는 `config/cal_share.local.json` 에만 둔다(git 밖 — .gitignore 의 config/*.local.json).
  · 공유 주소는 `.../cal.html#k=<열쇠>` 다. `#` 뒤는 **서버로 가지 않는다** —
    깃허브 접속 기록에도, 검색엔진에도 남지 않는다. 페이지에는 noindex 도 건다.
  · 받는 사람은 누르기만 하면 된다. PIN 을 묻지 않는다.

열쇠를 바꾸고 싶으면(누가 링크를 흘렸을 때): `python cal_share.py --new-key`
그러면 예전 링크는 그 즉시 못 연다. 새 링크를 다시 공유하면 된다.

  python cal_share.py            # cal.enc 갱신 + 공유 주소 출력
  python cal_share.py --new-key  # 열쇠를 새로 뽑고 갱신(예전 링크 무효)
  python cal_share.py --url      # 지금 공유 주소만 보기
"""
import base64
import json
import os
import secrets
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "webapp"))

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

KEYFILE = os.path.join(ROOT, "config", "cal_share.local.json")
OUT = os.path.join(ROOT, "docs", "cal.enc")
BASE = "https://mulder4780.github.io/coupang-ecount-reconcile"
KEEP = ("날짜", "시간", "분류", "캠프명", "장소", "제목", "프로젝트NO", "담당기사", "예측")


def key(new=False):
    """공유 열쇠. 없으면 만들고, --new-key 면 새로 뽑는다(예전 링크는 그 즉시 무효)."""
    os.makedirs(os.path.dirname(KEYFILE), exist_ok=True)
    if not new and os.path.exists(KEYFILE):
        try:
            return json.load(open(KEYFILE, encoding="utf-8"))["key"]
        except Exception:
            pass
    k = base64.urlsafe_b64encode(secrets.token_bytes(24)).decode().rstrip("=")
    json.dump({"key": k, "at": datetime.now().strftime("%Y-%m-%d %H:%M")},
              open(KEYFILE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"※ 공유 열쇠를 {'새로 ' if new else ''}만들었습니다 — config/cal_share.local.json"
          + (" · 예전 링크는 이제 열리지 않습니다" if new else ""))
    return k


def build():
    import app_server as A
    cal = A.get_calendar()
    year = str(datetime.now().year)
    return {
        "갱신": cal.get("갱신") or datetime.now().strftime("%Y-%m-%d %H:%M"),
        "분류목록": cal.get("분류목록") or [],
        # 올해 것만. 지난해까지 담으면 사본이 커지고, 현장이 볼 일도 거의 없다.
        "일정": [{k: e.get(k) for k in KEEP if e.get(k) not in (None, "")}
               for e in (cal.get("일정") or [])
               if str(e.get("날짜") or "").startswith(year)],
    }


def main():
    a = sys.argv[1:]
    k = key(new="--new-key" in a)
    url = f"{BASE}/cal.html#k={k}"
    if "--url" in a:
        print(url)
        return 0
    import csos_crypto as C
    data = build()
    body = json.dumps(data, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump(C.seal(body, k), open(OUT, "w", encoding="utf-8"))
    print(f"캘린더 공유본: 일정 {len(data['일정'])}건 · {len(body):,}바이트 → docs/cal.enc")
    print(f"공유 주소: {url}")
    print("  ※ 이 주소를 아는 사람은 PIN 없이 일정만 봅니다. 정산·계산서는 열리지 않습니다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
