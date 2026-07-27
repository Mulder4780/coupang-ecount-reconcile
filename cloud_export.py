# -*- coding: utf-8 -*-
"""
cloud_export.py — 클라우드 서버가 화면을 그리는 데 필요한 **데이터 한 덩어리**를 만든다
================================================================================
PC가 꺼져도 폰에서 보려면 데이터가 PC 밖에 있어야 한다. 그렇다고 관리대장 엑셀을
클라우드로 옮길 수는 없다(회사 Z: 드라이브가 원본이고, 매일 사람이 직접 입력한다).

  → **엑셀은 회사에 그대로 두고**, 앱이 화면에 쓰는 값만 뽑아 JSON 한 파일로 만든다.
    클라우드 서버는 이 파일만 읽어 같은 화면을 그린다(엑셀도, Z: 드라이브도 필요 없다).

담기는 것: 정산·업무·확인필요·대표보고·ERP서류·계산서구성·검증배지·상태
담기지 않는 것: 비밀키(config/*.json), 엑셀 원본, 밴드 원문 캐시

실행
  python cloud_export.py                      # reports/cloud_bundle.json 생성
  python cloud_export.py --upload             # 만들고 클라우드에 올리기
  python cloud_export.py --upload --url ... --token ...
설정은 config/cloud.json 에 둔다(비밀이므로 커밋 금지):
  {"url": "https://내주소/api/push", "token": "긴-임의-문자열"}
"""
import sys, os, json, gzip, urllib.request
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "webapp"))

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

OUT = os.path.join(ROOT, "reports", "cloud_bundle.json")
CFG = os.path.join(ROOT, "config", "cloud.json")


def build():
    import app_server as A
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])

    def safe(fn, default):
        try:
            return fn()
        except Exception as e:
            print(f"   ! {getattr(fn, '__name__', fn)} 실패: {type(e).__name__} {e}")
            return default

    works = safe(A.get_works, {"as": [], "pm": []})
    data = {
        "생성": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "원본": os.path.basename(master),
        # get_settlements는 행 리스트를 그대로 준다 — 앱 API 모양({rows:…})에 맞춰 감싼다
        "settlements": safe(lambda: {"rows": A.get_settlements()}, {"rows": []}),
        "works": works,
        "issues": safe(A.get_issues, {"rows": [], "cols": [], "source": ""}),
        "exec_report": safe(lambda: A.read_exec_report(master), {}),
        "erpdocs": safe(A.get_erpdocs, {"rows": [], "months": {}, "kinds": {}, "total": 0}),
        "checks": safe(A.get_checks, {}),
    }
    return data


def main():
    args = sys.argv[1:]
    print("클라우드 번들 생성 중…")
    data = build()
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
    open(OUT, "wb").write(raw)
    n = lambda k: len(((data.get(k) or {}).get("rows")) or [])
    print(f"정산 {n('settlements')} · AS {len(data['works'].get('as', []))} · "
          f"점검 {len(data['works'].get('pm', []))} · 확인필요 {n('issues')} · "
          f"ERP서류 {n('erpdocs')}")
    print(f"→ {OUT}  ({len(raw)/1024:.0f} KB)")

    if "--upload" not in args:
        print("\n올리려면: python cloud_export.py --upload")
        return
    url = token = ""
    if os.path.exists(CFG):
        c = json.load(open(CFG, encoding="utf-8"))
        url, token = c.get("url", ""), c.get("token", "")
    if "--url" in args:
        url = args[args.index("--url") + 1]
    if "--token" in args:
        token = args[args.index("--token") + 1]
    if not url or not token:
        sys.exit("올릴 주소·토큰이 없습니다 — config/cloud.json 을 만들어 주세요(예시는 이 파일 맨 위 설명 참고)")

    body = gzip.compress(raw)            # 1MB 넘는 JSON이라 압축해서 보낸다
    req = urllib.request.Request(url, data=body, method="POST", headers={
        "Content-Type": "application/json", "Content-Encoding": "gzip",
        "X-Token": token})
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            print("업로드 완료:", r.status, r.read(200).decode("utf-8", "replace"))
    except Exception as e:
        print("업로드 실패:", type(e).__name__, e)
        sys.exit(1)


if __name__ == "__main__":
    main()
