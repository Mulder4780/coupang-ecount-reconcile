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
import json, os, ssl, time, urllib.request, urllib.error
from datetime import datetime, date

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config", "ecount_config.json")
SESSION_CACHE = os.path.join(BASE_DIR, ".session_cache.json")


def load_config(path=CONFIG_PATH):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


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
        if not self.zone:
            self.fetch_zone()
        for field in ("COM_CODE", "USER_ID", "API_CERT_KEY"):
            if not (self.auth.get(field) or "").strip():
                raise EcountError(f"{field} 가 비어 있습니다. config 를 확인하세요.")
        url = self._root() + self.endpoints["login"]["path"]
        key = (self.auth.get("API_CERT_KEY_TEST") if self.auth.get("IS_TEST")
               else self.auth.get("API_CERT_KEY")) or self.auth.get("API_CERT_KEY")
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
                 "com": self.auth.get("COM_CODE"), "ts": datetime.now().isoformat()},
                open(SESSION_CACHE, "w", encoding="utf-8"),
            )
        except Exception:
            pass

    def _load_cached_session(self, max_age_min=30):
        try:
            d = json.load(open(SESSION_CACHE, "r", encoding="utf-8"))
            if d.get("com") != self.auth.get("COM_CODE"):
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
