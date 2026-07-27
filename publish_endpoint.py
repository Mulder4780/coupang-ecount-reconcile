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

    # ★ 마지막 안전장치. 여기서 한 번 더 거르지 않으면 잘못된 주소가 그대로 게시돼
    #   폰에서 '사이트에 연결할 수 없습니다'만 뜬다(2026-07-27에 api.trycloudflare.com이 실렸다).
    #   무료 터널 주소는 '단어-단어-단어-단어.trycloudflare.com' 꼴이다.
    import re as _re
    if not _re.fullmatch(r"https://[a-z0-9]+(?:-[a-z0-9]+){2,}\.trycloudflare\.com/?", url):
        print(f"게시 취소 — 터널 주소 형식이 아님: {url}")
        return
    # 정말 살아 있는지 확인하고 올린다(죽은 주소를 게시하면 폰이 그대로 막힌다).
    # ★ 단 **회사 DNS로 판단하면 안 된다** — 사내망은 *.trycloudflare.com 을 풀어 주지 않아
    #   멀쩡한 주소도 늘 '죽음'으로 나온다. 그 오판 때문에 게시가 계속 취소되고 터널만
    #   새로 만들어져, 폰은 영영 못 들어왔다(2026-07-27 원인 확정). net_probe가 공개 DNS로
    #   우회해 **바깥에서 보이는 진짜 상태**를 본다.
    from net_probe import probe
    ok, why = probe(url.rstrip("/") + "/api/ping")
    if not ok:
        print(f"게시 취소 — {why}: {url}")
        return
    print(f"   살아 있음 확인 — {why}")

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
