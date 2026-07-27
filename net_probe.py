# -*- coding: utf-8 -*-
"""
net_probe.py — 터널이 **바깥에서** 살아 있는지 회사 DNS를 거치지 않고 확인한다
================================================================================
2026-07-27, 오래 끌던 접속 불가의 진짜 원인:

    회사 DNS는 `*.trycloudflare.com` 을 풀어 주지 않는다(차단).
    공개 DNS(8.8.8.8)로 물으면 정상으로 나온다.

그래서 사무실 PC에서 터널 주소를 찔러 보면 늘 `getaddrinfo failed` 가 났고,
  · publish_endpoint 는 "죽은 주소"라며 게시를 취소하고,
  · tunnel_run 의 감시 루프는 "터널이 죽었다"며 주소를 새로 만들었다.
정작 폰(LTE)에서는 그 주소가 **멀쩡히 살아 있었다**. PC 혼자 못 보고 있었을 뿐이다.
주소만 끝없이 바뀌니 폰 북마크는 매번 무용지물이 됐다.

해결: 회사 DNS를 건너뛴다.
  ① 평범하게 한 번 시도한다(집·테더링 등 정상 DNS 환경에서는 이걸로 끝).
  ② 이름을 못 풀면 **공개 DNS에 HTTPS로 직접 물어** IP를 얻고,
     그 IP에 붙되 SNI·Host 헤더는 원래 이름으로 보낸다(인증서·라우팅 정상).

표준 라이브러리만 쓴다(설치 0 원칙).
"""
import sys, json, socket, ssl, urllib.request, urllib.parse

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# 이 두 곳은 회사 DNS로도 잘 풀린다(막힌 건 trycloudflare 쪽이다).
DOH = ("https://dns.google/resolve?name={name}&type=A",
       "https://cloudflare-dns.com/dns-query?name={name}&type=A")


def resolve_public(host, timeout=8):
    """공개 DNS에 HTTPS로 물어 A 레코드를 얻는다(회사 DNS를 거치지 않는다)."""
    for tpl in DOH:
        try:
            req = urllib.request.Request(tpl.format(name=urllib.parse.quote(host)),
                                         headers={"accept": "application/dns-json"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                d = json.load(r)
            ips = [a["data"] for a in (d.get("Answer") or []) if a.get("type") == 1]
            if ips:
                return ips
        except Exception:
            continue
    return []


def _get_via_ip(ip, host, path, timeout):
    """IP에 직접 붙되 **이름은 그대로** 알린다 — SNI와 Host가 맞아야 터널이 응답한다."""
    ctx = ssl.create_default_context()
    with socket.create_connection((ip, 443), timeout=timeout) as raw:
        with ctx.wrap_socket(raw, server_hostname=host) as s:
            s.sendall(("GET %s HTTP/1.1\r\nHost: %s\r\n"
                       "User-Agent: CSOS-probe\r\nConnection: close\r\n\r\n"
                       % (path, host)).encode())
            head = b""
            while b"\r\n" not in head and len(head) < 4096:
                b = s.recv(1024)
                if not b:
                    break
                head += b
    line = head.split(b"\r\n", 1)[0].decode("latin-1", "replace")
    parts = line.split()
    return int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0


def probe(url, timeout=12):
    """(살아있나, 설명) — 회사 DNS가 막고 있어도 진실을 말한다."""
    u = urllib.parse.urlparse(url if "://" in url else "https://" + url)
    host, path = u.hostname, (u.path or "/")
    if not host:
        return False, "주소 형식 오류"

    # ① 평범한 경로
    try:
        with urllib.request.urlopen(url, timeout=timeout) as r:
            return r.status == 200, f"정상 응답 {r.status}"
    except urllib.error.HTTPError as e:
        # 530 등은 '터널은 있는데 뒤쪽 앱이 안 붙은' 상태 — 살아 있다고 하면 안 된다
        return False, f"HTTP {e.code} (터널은 있으나 앱이 응답하지 않음)"
    except Exception as e:
        first = f"{type(e).__name__}"

    # ② 이름을 못 풀었을 뿐일 수 있다 — 공개 DNS로 우회해 다시 본다
    ips = resolve_public(host)
    if not ips:
        return False, f"이름을 어디서도 풀지 못함 ({first})"
    for ip in ips[:3]:
        try:
            code = _get_via_ip(ip, host, path, timeout)
            if code == 200:
                return True, f"정상(회사 DNS가 막아 공개 DNS로 확인 · {ip})"
            if code:
                return False, f"HTTP {code} (앱이 응답하지 않음 · {ip})"
        except Exception:
            continue
    return False, f"공개 DNS로는 이름이 있으나 접속 실패 ({first})"


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else None
    if not target:
        import os
        p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs", "endpoint.json")
        target = json.load(open(p, encoding="utf-8"))["url"]
    if "/api/" not in target:                 # 주소만 준 경우에만 붙인다
        target = target.rstrip("/") + "/api/ping"
    ok, why = probe(target)
    print(("살아 있음" if ok else "응답 없음") + " — " + why)
    print("  " + target)
