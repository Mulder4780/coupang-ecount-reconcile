# -*- coding: utf-8 -*-
"""
publish_endpoint.py — 고정 주소(GitHub Pages)에 현재 접속 주소를 게시
=======================================================================
외부 터널(trycloudflare)은 재시작하면 주소가 바뀐다. 폰 북마크를 매번 고치지 않도록
**바뀌지 않는 GitHub Pages 주소**를 진입점으로 두고, 이 스크립트가 현재 주소를
docs/endpoint.json 에 써서 커밋·푸시한다. 폰은 고정 주소만 열면 자동으로 전달된다.

  폰이 여는 고정 주소 :  https://<user>.github.io/<repo>/
  실제 접속 주소       :  endpoint.json 의 url (바뀌어도 폰은 신경 쓸 필요 없음)

주소가 이전과 같으면 아무것도 하지 않는다(불필요한 커밋 방지).
워치독이 매 주기 호출하므로 사람이 실행할 일은 없다.

실행:  python publish_endpoint.py [--force]
"""
import sys, os, json, subprocess
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
URL_FILE = os.path.join(ROOT, "reports", "tunnel_url.txt")
EP_FILE = os.path.join(ROOT, "docs", "endpoint.json")


def git(*args, timeout=90):
    return subprocess.run(["git"] + list(args), cwd=ROOT, capture_output=True,
                          text=True, encoding="utf-8", errors="replace", timeout=timeout)


def main():
    force = "--force" in sys.argv
    try:
        url = open(URL_FILE, encoding="utf-8").read().strip()
    except Exception:
        url = ""
    if not url:
        print("게시할 주소 없음 (터널 미가동)")
        return

    prev = ""
    try:
        prev = json.load(open(EP_FILE, encoding="utf-8")).get("url", "")
    except Exception:
        pass
    if url == prev and not force:
        print("주소 동일 — 게시 생략")
        return

    os.makedirs(os.path.dirname(EP_FILE), exist_ok=True)
    json.dump({"url": url, "updated": datetime.now().isoformat(timespec="seconds")},
              open(EP_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    git("add", "docs/endpoint.json")
    st = git("diff", "--cached", "--name-only").stdout.strip()
    if not st:
        print("변경 없음")
        return
    c = git("commit", "-q", "-m", f"endpoint: {url.split('//')[-1][:40]}")
    p = git("push", "-q", "origin", "master")
    ok = p.returncode == 0
    print(f"게시 {'완료' if ok else '실패'} → {url}")
    if not ok:
        print((p.stderr or "")[:200])


if __name__ == "__main__":
    main()
