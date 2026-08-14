# -*- coding: utf-8 -*-
"""
ecount_client.py — 이카운트(ECOUNT) OAPI 클라이언트
====================================================
인증 흐름(공식): Zone 조회 → OAPILogin(세션ID 발급) → 각 조회 API 호출.

- 외부 라이브러리 불필요(파이썬 표준 urllib만 사용).
- 인증키 등 비밀정보는 config/ecount_config.json 에서만 읽습니다(코드에 하드코딩 금지).
- 세션ID는 .session_cache.json 에 잠시 저장해 반복 로그인을 줄입니다.

사용:
    from ecount_client import EcountClient, load_config
    cfg = load_config()
    cli = EcountClient(cfg)
    cli.login()                       # Zone 자동조회 + 로그인
    rows = cli.inquiry("sale", {"BASE_DATE_FROM": "20260701", "BASE_DATE_TO": "20260731"})
"""
import json, os, ssl, sys, time, urllib.request, urllib.error
from datetime import datetime, date

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config", "ecount_config.json")
SESSION_CACHE = os.path.join(BASE_DIR, ".session_cache.json")


def load_config(path=CONFIG_PATH):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def is_placeholder(s):
    """템플릿 문구가 그대로 남은 것은 **넣은 것이 아니다.**
    비어 있으면 사람이 알아채지만 '여기에_...' 는 값처럼 보여서 안 띈다 -
    그대로 로그인을 시도하면 그 실패가 '키가 틀렸다' 로 읽힌다(넣지 않은 것인데도)."""
    s = str(s or "").strip()
    return (not s) or s.startswith("여기에") or s.startswith("<")


class EcountError(RuntimeError):
    pass


class EcountClient:
    def __init__(self, config):
        self.cfg = config
        self.auth = config["auth"]
        self.net = config.get("network", {})
        self.endpoints = config["endpoints"]
        self.zone = (self.auth.get("ZONE") or "").strip()
        self.session_id = None
        self._ctx = ssl.create_default_context()

    # ---------- 도메인 ----------
    def _root(self, with_zone=True):
        sub = "sboapi" if self.auth.get("IS_TEST") else "oapi"
        return f"https://{sub}{self.zone if with_zone else ''}.ecount.com"

    # ---------- 저수준 POST ----------
    def _post(self, url, body):
        data = json.dumps(body).encode("utf-8")
        req = urllib.request.Request(
            url, data=data, method="POST",
            headers={"Content-Type": "application/json; charset=utf-8"},
        )
        timeout = self.net.get("timeout_sec", 30)
        retry = self.net.get("retry", 2)
        last = None
        for attempt in range(retry + 1):
            try:
                with urllib.request.urlopen(req, timeout=timeout, context=self._ctx) as r:
                    raw = r.read().decode("utf-8", "replace")
                return json.loads(raw) if raw.strip() else {}
            except urllib.error.HTTPError as e:
                detail = e.read().decode("utf-8", "replace")[:500]
                raise EcountError(f"HTTP {e.code} {url}\n{detail}")
            except (urllib.error.URLError, TimeoutError) as e:
                last = e
                time.sleep(1.0 * (attempt + 1))
        raise EcountError(f"네트워크 실패({url}): {last}\n"
                          f"→ 이카운트에서 이 PC의 공인 IP를 [IP등록]했는지 확인하세요.")

    @staticmethod
    def _dig(obj, *keys):
        """중첩 dict에서 후보 키를 순서대로 탐색(응답 구조 변형 대비)."""
        cur = obj
        for k in keys:
            if isinstance(cur, dict) and k in cur:
                cur = cur[k]
            else:
                return None
        return cur

    def _find_first(self, obj, target):
        """응답 어딘가에 있는 target 키의 첫 값을 재귀 탐색."""
        if isinstance(obj, dict):
            for k, v in obj.items():
                if k == target and v not in (None, ""):
                    return v
                found = self._find_first(v, target)
                if found is not None:
                    return found
        elif isinstance(obj, list):
            for it in obj:
                found = self._find_first(it, target)
                if found is not None:
                    return found
        return None

    # ---------- 1) Zone ----------
    def fetch_zone(self):
        com = (self.auth.get("COM_CODE") or "").strip()
        if not com:
            raise EcountError("COM_CODE(회사코드)가 비어 있습니다. config/ecount_config.json 의 auth.COM_CODE 를 채우세요.")
        url = self._root(with_zone=False) + self.endpoints["zone"]["path"]
        resp = self._post(url, {"COM_CODE": com})
        zone = self._find_first(resp, "ZONE")
        if not zone:
            raise EcountError(f"Zone 조회 실패. 응답={json.dumps(resp, ensure_ascii=False)[:400]}")
        self.zone = str(zone).strip()
        return self.zone

    # ---------- 2) Login ----------
    def login(self, force=False):
        if not force and self._load_cached_session():
            return self.session_id
        # ★ 사용자 지시(2026-07-30): IP가 바뀌면 이카운트 [IP등록] 화면에 넣고 진행한다.
        #   OAPI는 등록된 IP에서만 동작한다. 미등록 상태로 부르면 실패가 **인증 오류처럼**
        #   보여 원인을 엉뚱한 데서 찾게 되고, 반복 실패는 트래픽 제한을 건드려 ERP 전체
        #   차단으로 번질 수 있다(AGENTS.md 절대규칙). 그래서 부르기 전에 멈춘다.
        #   캐시된 세션이 살아 있으면 이 관문을 지나지 않는다 — 이미 되는 IP라는 뜻이다.
        try:
            import erp_ip_guard
            erp_ip_guard.require()
        except SystemExit:
            raise
        except Exception:
            pass                      # 판정 자체가 불가하면(모듈·인터넷 문제) 막지는 않는다
        if not self.zone:
            self.fetch_zone()
        # ★ 테스트존(sboapi)과 실서비스(oapi)는 **인증키가 서로 다르다.**
        #   예전에는 테스트키가 비면 `or` 로 실서비스 키를 흘려보냈다. 그러면 sboapi 가
        #   그 키를 거절하고, 그 실패가 **"테스트키가 잘못됐다"** 로 읽힌다 —
        #   실제로는 키를 아예 안 넣은 것이다. 넣지 않은 것과 틀린 것은 다른 사실이므로
        #   갈라서 말한다(안 그러면 멀쩡한 키를 의심하며 재발급을 신청하게 된다).
        key_field = "API_CERT_KEY_TEST" if self.auth.get("IS_TEST") else "API_CERT_KEY"
        for field in ("COM_CODE", "USER_ID", key_field):
            if is_placeholder(self.auth.get(field)):
                hint = ""
                if field == "API_CERT_KEY_TEST":
                    hint = ("\n  → IS_TEST=true 입니다. 테스트존 인증키를 받아 이 칸에 넣으세요."
                            "\n  → 실서비스로 쓰려면 auth.IS_TEST 를 false 로 되돌리세요.")
                raise EcountError(f"{field} 가 비어 있습니다. config 를 확인하세요.{hint}")
        url = self._root() + self.endpoints["login"]["path"]
        key = self.auth[key_field]
        body = {
            "COM_CODE": self.auth["COM_CODE"],
            "USER_ID": self.auth["USER_ID"],
            "API_CERT_KEY": key,
            "LAN_TYPE": self.auth.get("LAN_TYPE", "ko-KR"),
            "ZONE": self.zone,
        }
        resp = self._post(url, body)
        sid = self._find_first(resp, "SESSION_ID")
        if not sid:
            raise EcountError(f"로그인 실패(세션ID 없음). 응답={json.dumps(resp, ensure_ascii=False)[:400]}")
        self.session_id = str(sid)
        self._save_cached_session()
        return self.session_id

    # ---------- 3) 조회 ----------
    def inquiry(self, endpoint_key, body=None):
        """config.endpoints[endpoint_key] 의 domain/method 로 조회. 결과 rows(list) 반환."""
        if not self.session_id:
            self.login()
        ep = self.endpoints[endpoint_key]
        if "domain" not in ep or "method" not in ep:
            raise EcountError(f"endpoints.{endpoint_key} 에 domain/method 가 없습니다.")
        url = f"{self._root()}/OAPI/V2/{ep['domain']}/{ep['method']}?SESSION_ID={self.session_id}"
        resp = self._post(url, body or {})
        # ECOUNT 조회 응답은 보통 Data.Result / Data.Datas 아래 배열
        for path in (("Data", "Result"), ("Data", "Datas"), ("Data",)):
            node = self._dig(resp, *path)
            if isinstance(node, list):
                return node
        node = self._dig(resp, "Data")
        if isinstance(node, dict):
            for v in node.values():
                if isinstance(v, list):
                    return v
        return resp if isinstance(resp, list) else [resp]

    # ---------- 세션 캐시 ----------
    def _save_cached_session(self):
        try:
            json.dump(
                {"zone": self.zone, "session_id": self.session_id,
                 "com": self.auth.get("COM_CODE"),
                 "is_test": bool(self.auth.get("IS_TEST")),
                 "ts": datetime.now().isoformat()},
                open(SESSION_CACHE, "w", encoding="utf-8"),
            )
        except Exception:
            pass

    def _load_cached_session(self, max_age_min=30):
        try:
            d = json.load(open(SESSION_CACHE, "r", encoding="utf-8"))
            if d.get("com") != self.auth.get("COM_CODE"):
                return False
            # ★ 실서비스(oapi)와 테스트존(sboapi)은 **세션이 서로 통하지 않는다.**
            #   그런데 예전에는 회사코드만 비교해서, 실서비스로 로그인해 둔 세션이
            #   IS_TEST 로 바꾼 뒤에도 그대로 재사용됐다. 그러면 테스트인 줄 알고
            #   부른 SaveSale 이 **실제 ERP 에 전표를 넣는다** — 되돌릴 수 없는 쪽이다.
            #   ★ 옛 캐시에는 이 칸이 아예 없다. 없으면 '실서비스였겠지'로 짐작하지
            #     않고 **모름으로 보고 버린다** — 로그인 한 번이 잘못 쓴 전표보다 싸다.
            if d.get("is_test") is None or bool(d["is_test"]) != bool(self.auth.get("IS_TEST")):
                return False
            age = (datetime.now() - datetime.fromisoformat(d["ts"])).total_seconds() / 60
            if age > max_age_min:
                return False
            self.zone = d.get("zone") or self.zone
            self.session_id = d.get("session_id")
            return bool(self.session_id)
        except Exception:
            return False


def days_until_expiry(config):
    try:
        end = datetime.strptime(config["auth"]["유효기간_만료"], "%Y-%m-%d").date()
        return (end - date.today()).days
    except Exception:
        return None


# ===================== 사람이 손으로 부르는 점검 =====================
# ★ 여기서 **조회 API 는 부르지 않는다.** Zone -> OAPILogin 까지가 끝이다.
#   "되는지 보자"고 이것저것 찔러 보는 것이 곧 무차별 탐침이고, 트래픽 제한을
#   건드리면 ERP 계정 전체가 막힌다(AGENTS.md 절대규칙 3).

def _mask(s):
    s = str(s or "")
    if not s:
        return "(비어있음)"
    if is_placeholder(s):
        return "(템플릿 그대로 - 미설정)"
    return "설정됨 len=%d %s...%s" % (len(s), s[:3], s[-2:])


def _save_config(cfg, path=CONFIG_PATH):
    """비밀설정을 원자적으로 갈아끼운다(반쯤 쓰다 죽으면 로그인 자체가 막힌다)."""
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _check(argv):
    cfg = load_config()
    a = cfg["auth"]
    is_test = bool(a.get("IS_TEST"))
    key_field = "API_CERT_KEY_TEST" if is_test else "API_CERT_KEY"
    print("[1] 설정")
    print("    모드        :", "테스트존(sboapi)" if is_test else "실서비스(oapi)")
    print("    COM_CODE    :", a.get("COM_CODE") or "(비어있음)")
    print("    USER_ID     :", a.get("USER_ID") or "(비어있음)")
    print("    %-12s:" % key_field, _mask(a.get(key_field)))
    if is_test:
        print("    (실서비스키):", _mask(a.get("API_CERT_KEY")), "<- 이 모드에서는 안 씁니다")
    d = days_until_expiry(cfg)
    print("    만료까지    :", ("%d일" % d) if d is not None else "(모름)")
    if is_placeholder(a.get(key_field)):
        print("\n  x %s 가 비어 있습니다. 아직 키를 안 넣었습니다." % key_field)
        print("    넣기: python ecount/ecount_client.py --set-test-key")
        return 2

    print("[2] 허용 IP")
    try:
        import erp_ip_guard
        erp_ip_guard.require()
        print("    OK - 지금 공인 IP 가 등록돼 있습니다")
    except SystemExit as e:
        print("    x", str(e.code).splitlines()[0])
        print("    (등록 전에는 OAPI 가 동작하지 않습니다. 여기서 멈춥니다)")
        return 2
    except Exception as e:
        # ★ 판정 자체를 못 한 것과 '등록 안 됨'은 다른 사실이다. 뭉치면
        #   못 물어본 것을 이상 없음으로 읽는다.
        print("    ? 확인 못 함(%s: %s) - 아래가 실패하면 IP 부터 의심하세요" % (type(e).__name__, e))

    cli = EcountClient(cfg)
    print("[3] Zone 조회")
    try:
        print("    OK ->", cli.fetch_zone())
    except Exception as e:
        print("    x", e)
        return 1
    print("[4] 로그인(OAPILogin)")
    try:
        sid = cli.login(force=True)
        print("    OK - 세션ID 발급됨 len=%d" % len(sid))
    except Exception as e:
        print("    x", e)
        return 1
    print("\n=> 인증까지 정상입니다. 조회 API 는 여기서 부르지 않습니다(절대규칙 3).")
    return 0


def _set_test_key(argv):
    """테스트존 인증키를 넣는다. 키는 **인자로 받지 않는다** - 명령 인자는 셸
    기록에 남고 그 기록은 아무도 안 지운다. 붙여넣기로만 받는다."""
    print("이카운트 테스트존 인증키를 붙여넣고 Enter (취소는 빈 줄):")
    try:
        key = sys.stdin.readline().strip()
    except Exception:
        key = ""
    if not key:
        print("취소했습니다. 아무것도 바꾸지 않았습니다.")
        return 1
    cfg = load_config()
    cfg["auth"]["API_CERT_KEY_TEST"] = key
    cfg["auth"].setdefault("_API_CERT_KEY_TEST_설명",
                           "테스트존(sboapi) 전용 인증키. IS_TEST=true 일 때만 쓴다.")
    _save_config(cfg)
    print("넣었습니다:", _mask(key))
    print("테스트존으로 전환: python ecount/ecount_client.py --use test")
    return 0


def _use(argv):
    want = (argv[0] if argv else "").lower()
    if want not in ("test", "live"):
        print("사용: --use test | --use live")
        return 1
    cfg = load_config()
    is_test = (want == "test")
    if is_test and is_placeholder(cfg["auth"].get("API_CERT_KEY_TEST")):
        # ★ 키 없이 켜면 로그인이 실패하는데 그 실패가 '키가 틀렸다'로 읽힌다.
        print("x 테스트키가 아직 없습니다. 먼저 --set-test-key 로 넣으세요.")
        return 2
    cfg["auth"]["IS_TEST"] = is_test
    _save_config(cfg)
    try:
        os.remove(SESSION_CACHE)          # 모드가 바뀌면 옛 세션은 통하지 않는다
    except OSError:
        pass
    print("전환:", "테스트존(sboapi)" if is_test else "실서비스(oapi)")
    print("확인: python ecount/ecount_client.py --check")
    return 0


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "--check"
    rest = sys.argv[2:]
    fn = {"--check": _check, "--set-test-key": _set_test_key, "--use": _use}.get(cmd)
    if fn is None:
        print("사용: ecount_client.py [--check | --set-test-key | --use test|live]")
        raise SystemExit(1)
    raise SystemExit(fn(rest))
