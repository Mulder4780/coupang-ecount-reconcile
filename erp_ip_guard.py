# -*- coding: utf-8 -*-
"""
erp_ip_guard.py — 공인 IP가 바뀌면 **ERP를 두드리기 전에** 알아채고 등록을 안내한다
===============================================================================
사용자 지시(2026-07-30): "앞으로 IP 주소가 변경되면 이 화면에서 IP 주소 등록해서 진행해
(알고리즘 반영)"

이카운트 OAPI는 **등록된 IP에서만** 동작한다(최대 20개). 등록 화면:
  이카운트 → Self-Customizing → 정보관리 → **API인증키발급** → [IP등록] → 저장(F8)

왜 자동 등록을 하지 않는가
  그 화면은 회사 ERP의 **보안 설정**이다. 자동으로 눌러 바꾸면 안 되는 종류의 조작이라
  이 도구는 **판단과 안내까지만** 한다. 사람이 20초면 끝나는 일을, 되돌리기 어려운
  자동 조작으로 바꾸지 않는다.

무엇을 하는가
  1. 지금 공인 IP를 확인한다(여러 서비스로 교차 확인 — 한 곳이 죽어도 판정은 계속된다).
  2. 등록 목록(config/erp_allowed_ips.json)과 비교한다.
  3. 안 맞으면 **ERP API 호출을 미리 막고**(require) 무엇을 해야 하는지 정확히 알린다.
     ★ 막는 이유: 미등록 IP로 호출하면 실패가 인증 오류처럼 보여 원인을 엉뚱한 데서 찾는다.
       게다가 반복 실패는 트래픽 제한을 건드려 ERP 차단으로 번질 수 있다(AGENTS.md 절대규칙).
  4. 앱·리포트가 읽을 캐시를 남긴다(reports/erp_ip.json).

IP 목록은 회사 설정값이라 `config/erp_allowed_ips.json` 에 두고 커밋하지 않는다(규칙 1).

사용
  python erp_ip_guard.py                # 지금 상태 확인 + 캐시 갱신
  python erp_ip_guard.py --register IP  # 등록 목록에 추가(사람이 ERP에 넣은 뒤 기록용)
  python erp_ip_guard.py --self-test
"""
import sys, os, re, json, urllib.request
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

CFG = os.path.join(ROOT, "config", "erp_allowed_ips.json")
CACHE = os.path.join(ROOT, "reports", "erp_ip.json")
NOTE = os.path.join(ROOT, "reports", "ERP_IP_등록필요.md")
REGISTER_URL = "https://loginab.ecount.com/ec5/view/erp?w_flag=1"
HOWTO = ("이카운트 → Self-Customizing → 정보관리 → API인증키발급 → [IP등록] "
         "→ 빈 칸에 IP 입력 → 저장(F8)")
# 여러 곳에 물어본다. 한 서비스가 죽었거나 이상한 값을 주면 다수결로 걸러진다.
IP_SERVICES = ("https://api.ipify.org", "https://ifconfig.me/ip", "https://icanhazip.com",
               "https://ipv4.icanhazip.com")
RE_IP = re.compile(r"^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$")


# ── 순수 판정 (합성 검증 대상) ─────────────────────────────────
def valid_ip(text):
    m = RE_IP.match((text or "").strip())
    return bool(m) and all(0 <= int(g) <= 255 for g in m.groups())


def is_private(ip):
    """사설 IP는 OAPI 허용목록에서 **아무 일도 하지 않는다** — 20칸 중 한 칸만 버린다.

    ★ 실제로 목록에 192.168.219.108 이 등록돼 있었다(2026-07-30 확인). 사내망 주소라
      이카운트 서버가 볼 수 있는 주소가 아니다. 이런 칸은 지적해 준다."""
    if not valid_ip(ip):
        return False
    a, b = (int(x) for x in ip.split(".")[:2])
    return (a == 10 or a == 127 or (a == 172 and 16 <= b <= 31)
            or (a == 192 and b == 168) or (a == 169 and b == 254))


def decide(current, registered):
    """(상태, 사람이 읽는 한 줄). 상태: ok | need_register | unknown"""
    if not valid_ip(current):
        return "unknown", "공인 IP를 확인하지 못했습니다(인터넷 확인 필요) — ERP 호출을 보류합니다."
    if current in set(registered or []):
        return "ok", f"현재 IP {current} 는 등록돼 있습니다."
    return "need_register", (f"★ 현재 IP {current} 가 이카운트에 등록되지 않았습니다. "
                             f"등록 전에는 OAPI가 동작하지 않습니다.")


def slots_left(registered, limit=20):
    """남은 등록 칸. 사설 IP는 **쓸모없이 칸만 차지**하므로 따로 세어 알려준다."""
    reg = [r for r in (registered or []) if valid_ip(r)]
    return max(0, limit - len(reg)), [r for r in reg if is_private(r)]


# ── 원천 ─────────────────────────────────────────────────────
def registered_ips():
    try:
        d = json.load(open(CFG, encoding="utf-8"))
        return [str(x).strip() for x in d.get("ips", []) if str(x).strip()]
    except Exception:
        return []


def public_ip(timeout=8):
    """여러 서비스에 물어 **가장 많이 나온 값**을 쓴다. 한 곳만 믿지 않는다."""
    votes = {}
    for u in IP_SERVICES:
        try:
            req = urllib.request.Request(u, headers={"User-Agent": "coupang-work/1.0"})
            ip = urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "replace").strip()
            if valid_ip(ip):
                votes[ip] = votes.get(ip, 0) + 1
                if votes[ip] >= 2:            # 두 곳이 같으면 그걸로 확정(불필요한 호출 안 함)
                    return ip
        except Exception:
            continue
    return max(votes, key=votes.get) if votes else ""


def save(state, ip, registered, msg):
    left, private = slots_left(registered)
    doc = {"확인": datetime.now().isoformat(timespec="seconds"), "상태": state,
           "현재IP": ip, "등록수": len(registered), "남은칸": left,
           "쓸모없는_사설IP": private, "안내": msg,
           "등록방법": HOWTO, "등록주소": REGISTER_URL}
    os.makedirs(os.path.dirname(CACHE), exist_ok=True)
    json.dump(doc, open(CACHE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    if state == "need_register":
        with open(NOTE, "w", encoding="utf-8") as fh:
            fh.write(f"# ERP IP 등록 필요 ({datetime.now():%Y-%m-%d %H:%M})\n\n")
            fh.write(f"현재 공인 IP: **{ip}**\n\n등록해야 ERP(OAPI) 조회가 동작합니다.\n\n")
            fh.write(f"1. {HOWTO}\n2. 아래 값을 빈 칸에 넣고 저장\n\n```\n{ip}\n```\n\n")
            fh.write(f"3. 등록한 뒤 이 명령으로 기록: `python ecount/erp_ip_guard.py --register {ip}`\n\n")
            fh.write(f"- 등록 가능 20개 중 현재 {len(registered)}개 사용 · 남은 칸 {left}개\n")
            if private:
                fh.write(f"- ⚠ 사설 IP {', '.join(private)} 는 OAPI에서 쓸 수 없습니다(칸만 차지) — 지우면 칸이 늘어납니다.\n")
    elif os.path.exists(NOTE):
        os.remove(NOTE)                       # 해결되면 흔적을 남기지 않는다
    return doc


def check(quiet=True):
    reg = registered_ips()
    ip = public_ip()
    state, msg = decide(ip, reg)
    doc = save(state, ip, reg, msg)
    if not quiet:
        print(msg)
    return doc


def require():
    """ERP API를 부르기 직전에 호출한다. 등록 안 됐으면 **부르지 않고 멈춘다**."""
    doc = check()
    if doc["상태"] == "ok":
        return doc
    raise SystemExit(f"{doc['안내']}\n  → {HOWTO}\n  → 등록 후: python ecount/erp_ip_guard.py --register {doc['현재IP']}")


def register(ip):
    if not valid_ip(ip):
        sys.exit(f"IP 형식이 아닙니다: {ip}")
    reg = registered_ips()
    if ip in reg:
        print(f"이미 기록돼 있습니다: {ip}")
        return
    left, _ = slots_left(reg)
    if left <= 0:
        print("★ 등록 칸(20개)이 꽉 찼습니다 — 이카운트에서 안 쓰는 IP를 먼저 지우세요.")
    reg.append(ip)
    os.makedirs(os.path.dirname(CFG), exist_ok=True)
    json.dump({"_주의": "회사 ERP 설정값 — 커밋 금지(.gitignore 처리됨)",
               "_등록방법": HOWTO, "ips": reg},
              open(CFG, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"기록 완료: {ip} (총 {len(reg)}개) — 이카운트에도 실제로 등록했는지 확인하세요.")


# ── 합성 검증 ─────────────────────────────────────────────────
def self_test():
    bad = 0
    for ip, want in (("112.154.53.212", True), ("1.2.3.4", True), ("256.1.1.1", False),
                     ("1.2.3", False), ("", False), ("abc", False)):
        if valid_ip(ip) != want:
            print(f"  [FAIL] valid_ip({ip!r})"); bad += 1
    for ip, want in (("192.168.219.108", True), ("10.0.0.1", True), ("172.16.0.1", True),
                     ("172.32.0.1", False), ("112.154.53.212", False), ("127.0.0.1", True)):
        if is_private(ip) != want:
            print(f"  [FAIL] is_private({ip!r})"); bad += 1
    reg = ["112.154.53.212", "192.168.219.108"]
    if decide("112.154.53.212", reg)[0] != "ok":
        print("  [FAIL] 등록된 IP를 막는다"); bad += 1
    st, msg = decide("203.0.113.9", reg)
    if st != "need_register" or "203.0.113.9" not in msg:
        print(f"  [FAIL] 미등록 판정 {st}"); bad += 1
    # IP를 못 구했을 때 'ok' 로 넘어가면 미등록 상태로 ERP를 두드린다 — 그게 최악이다
    if decide("", reg)[0] != "unknown":
        print("  [FAIL] IP 불명인데 통과시킨다"); bad += 1
    left, private = slots_left(reg)
    if left != 18 or private != ["192.168.219.108"]:
        print(f"  [FAIL] 칸 계산 {left} {private}"); bad += 1
    print("erp_ip_guard self-test:", "OK" if not bad else f"{bad}건 실패")
    return bad == 0


def main():
    if "--self-test" in sys.argv:
        sys.exit(0 if self_test() else 1)
    if "--register" in sys.argv:
        i = sys.argv.index("--register")
        if i + 1 >= len(sys.argv):
            sys.exit("사용법: python erp_ip_guard.py --register 1.2.3.4")
        return register(sys.argv[i + 1])
    doc = check(quiet=False)
    print(f"  등록 {doc['등록수']}개 / 20 · 남은 칸 {doc['남은칸']}개")
    if doc["쓸모없는_사설IP"]:
        print(f"  ⚠ 사설 IP {', '.join(doc['쓸모없는_사설IP'])} 는 OAPI에서 무효 — 지우면 칸이 늘어납니다")
    if doc["상태"] != "ok":
        print(f"  → {HOWTO}")
        print(f"  → 안내 문서: {NOTE}")


if __name__ == "__main__":
    main()
