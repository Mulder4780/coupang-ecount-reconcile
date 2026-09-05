# -*- coding: utf-8 -*-
"""이카운트 OAPI에서 **공식적으로 조회가 확인된 자료만** 가져온다.

공개 안내와 이 회사 API인증현황에서 조회 가능함이 확인된 품목·발주서만 호출한다.
판매/매입전표·세금계산서·수금 조회 method를 짐작해 탐침하지 않는다. 그 자료는 기존
ERP 화면 XLSX → download_intake 경로가 정본이다.

    python erp_api_collect.py              # 6시간 안의 성공본이 있으면 재사용
    python erp_api_collect.py --force      # 품목 + 최근 120일 발주서 새로 조회
    python erp_api_collect.py --items
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from ecount_client import EcountClient, EcountError, load_config, mask_secrets


class ApiTrouble(EcountError):
    """무엇이 잘못됐는지 **갈래·조치·원문까지** 들고 다니는 실패."""

    def __init__(self, message: str, **fields: Any) -> None:
        super().__init__(message)
        self.fields = fields


# 문구 → 갈래. ★ 좁게 본다: 잘못 지목하면 사람이 **멀쩡한 것을 고치러 간다**(오타탐지 [172] 와 같은 문).
# ⚠ 맨몸 `"ip"` 를 표시로 쓰지 말 것 — `multiple`·`description`·`recipient` 에 다 들어 있어
#   엉뚱한 오류를 'IP 미등록' 이라 부른다. 그래서 표시는 낱말 경계를 가진 **정규식**이다.
_SIGNS = (
    ("IP 미등록", r"\bip\b|아이피|허용되지|등록되지 않은|not allowed",
     "이카운트 [API인증키발급 > IP등록] 에 이 PC 공인 IP 를 넣는다 · python erp_ip_guard.py"),
    ("인증 만료", r"session|login|expire|unauthor|세션|로그인|인증|만료",
     "python erp_api_collect.py --force  (세션 캐시를 버리고 다시 로그인)"),
    ("조회 API 아님", r"\bmethod\b|not found|notfound|지원하지|no service|invalid api",
     "이 자료는 ERP 화면 XLSX → download_intake 경로가 정본이다 (method 를 짐작해 탐침하지 않는다)"),
    # 2026-08-16 실측으로 드러난 갈래다. po_list 가 EXP00001 을 돌려줬다 —
    # method 는 살아 있고(404 가 아니라 구조화된 오류를 준다) **요청 본문**이 문제다.
    ("요청 본문 형식", r"exp00001|데이터 입력에 오류|입력 데이터 확인",
     "config/ecount_config.json 의 endpoints.<갈래> 요청 칸 이름을 이카운트 API 문서와 맞춘다"
     " (짐작으로 바꿔 가며 반복 호출하지 않는다 — 트래픽 제한이 ERP 전체 차단으로 번진다)"),
    # 2026-09-05 실측: 66회를 `code`(코드가 깨졌다)로 재시도했는데 진짜 원인은
    # **저쪽 서버가 503 을 준 것**이었다.  우리가 고칠 코드가 없다 —
    # ★ 그런데 조치가 `--force` 였다.  503 에 다시 두드리는 것은 절대규칙
    #   (이카운트 무차별 API 탐침 금지)이 막는 방향이다.
    # ★ 표시는 **우리 클라이언트가 만드는 한 줄**(`HTTP {code} {url}`)만 본다 —
    #   남의 HTML 본문까지 훑으면 'login' 같은 낱말에 걸려 옆 갈래를 삼킨다([172]).
    ("저쪽 서버가 안 받음", r"http 5[0-9][0-9][^0-9]|service unavailable|bad gateway",
     "이카운트 서버가 지금 안 받는다 — 코드에는 고칠 것이 없다. 돌아올 때까지 기다린다"
     " (다시 두드리지 않는다 — 트래픽 제한이 ERP 전체 차단으로 번진다)."
     " 급하면 ERP 화면 XLSX → download_intake 경로가 정본이다"),
)


def _diagnose(text: str) -> tuple[str, str]:
    """오류 문구에서 갈래를 고른다. **근거가 없으면 '모름' 이다 — 지어내지 않는다.**

    후보가 **유일할 때만** 지목한다. 둘에 걸리면 그것은 '둘 다'가 아니라 '모른다'이며,
    그때는 원문을 그대로 보여 사람이 판단한다."""
    low = str(text or "").lower()
    hits = [(kind, how) for kind, sign, how in _SIGNS if re.search(sign, low)]
    if len(hits) == 1:
        return hits[0]
    if len(hits) > 1:
        return "모름", ("여러 갈래에 걸린다(%s) — 원문을 보고 사람이 고른다"
                       % ", ".join(k for k, _ in hits))
    return "모름", "reports/erp_api_latest.json 의 '원문' 을 이카운트에 문의한다"


def _envelope(only: dict[str, Any]) -> tuple[str, str]:
    """오류 봉투에서 코드·문구를 **꺼낸다.** 지금까지는 있는 것을 버렸다."""
    # `Status:false` 는 코드가 아니라 실패 표시다 — 여기 넣으면 '코드 False' 라고 적힌다.
    code = _first(only, ("Code", "CODE", "ERROR_CODE", "ErrorCode", "ErrCode"), "")
    parts: list[str] = []
    for key in ("Message", "MESSAGE", "ErrorMessage", "Error", "Errors", "Detail", "RESULT_MSG"):
        value = only.get(key)
        if value in (None, "", [], {}):
            continue
        parts.append(value if isinstance(value, str)
                     else json.dumps(value, ensure_ascii=False, default=str))
    raw = " | ".join(parts) or json.dumps(only, ensure_ascii=False, default=str)
    return str(code), raw[:400]


ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "inbox" / "api"
STATUS = ROOT / "reports" / "erp_api_latest.json"
CONFIRMED = {
    "items": ("ERP:api_items", "품목"),
    "po_list": ("ERP:api_purchase_orders", "발주서"),
}


def _atomic_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    with temp.open("w", encoding="utf-8", newline="\n") as fh:
        json.dump(value, fh, ensure_ascii=False, indent=1, default=str)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(temp, path)


def _read(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
        return value if isinstance(value, dict) else {}
    except (OSError, ValueError, TypeError):
        return {}


#: 받아 둔 것을 다시 안 받는 창.  **한 곳에서 정한다**([162]) — 전체 캐시와
#: 갈래별 캐시가 서로 다른 값을 쓰면 같은 자료를 놓고 두 답이 나온다.
FRESH_HOURS = 6


def _fresh(hours: float = FRESH_HOURS) -> bool:
    try:
        return (datetime.now().timestamp() - STATUS.stat().st_mtime) < hours * 3600 and bool(_read(STATUS).get("ok"))
    except OSError:
        return False


def _fresh_source(endpoint: str, hours: float = FRESH_HOURS) -> dict[str, Any] | None:
    """그 갈래가 **최근에 성공했으면** 그때 결과를 돌려준다 — 아니면 ``None``.

    ★ **왜 갈래마다 보나** (2026-08-31 실사고).  `_fresh()` 는 `ok` 하나만 보는데
    한 갈래라도 실패하면 `ok=False` 다.  그래서 발주서가 못 풀리는 동안
    **멀쩡히 성공한 품목 7,346건을 18일간 57번 다시 받았다**(inbox/api 916.1MB).
    실패한 갈래는 사람이 `config/ecount_config.json` 을 고쳐야 풀리는데(요청 칸
    이름) 그때까지 **같은 요청이 이카운트로 계속 나간다** — 절대규칙이 막는
    '무차별 탐침'에 닿는 자리다(트래픽 제한이 ERP 전체 차단으로 번진다).

    ★ **실패한 갈래는 여기서 안 걸린다** — `count is None` 이라 매번 다시 시도한다.
      곧 이 문은 **성공을 아끼는 것**이지 실패를 덮는 것이 아니다([172]).
    ★ **못 읽으면 ``None``**([169]) — 모름을 '받아 뒀다'로 치면 새 자료가
      영영 안 들어오면서 화면은 멀쩡해 보인다.
    """
    try:
        age = datetime.now().timestamp() - STATUS.stat().st_mtime
    except OSError:
        return None
    if age >= hours * 3600:
        return None
    got = (_read(STATUS).get("sources") or {}).get(endpoint)
    if isinstance(got, dict) and got.get("count") is not None:
        return dict(got)
    return None


def _first(row: dict[str, Any], names: tuple[str, ...], default: Any = "") -> Any:
    for name in names:
        value = row.get(name)
        if value not in (None, ""):
            return value
    return default


def _day(value: Any) -> str:
    text = "".join(ch for ch in str(value or "") if ch.isdigit())
    return f"{text[:4]}-{text[4:6]}-{text[6:8]}" if len(text) >= 8 else ""


def _amount(value: Any) -> int | None:
    try:
        return int(round(float(str(value).replace(",", ""))))
    except (TypeError, ValueError):
        return None


def _natural(endpoint: str, row: dict[str, Any]) -> str:
    candidates = {
        "items": ("PROD_CD", "PROD_CODE", "ITEM_CD", "CODE"),
        "po_list": ("IO_NO", "DOC_NO", "ORDER_NO", "SER_NO"),
    }[endpoint]
    found = str(_first(row, candidates)).strip()
    if found:
        return f"{endpoint}/{found}"
    body = json.dumps(row, ensure_ascii=False, sort_keys=True, default=str)
    return f"{endpoint}/sha1:{hashlib.sha1(body.encode('utf-8')).hexdigest()}"


def _ingest(endpoint: str, rows: list[dict[str, Any]]) -> dict[str, int]:
    import datalake
    kind = CONFIRMED[endpoint][0]
    made = changed = same = 0
    con = datalake.connect()
    try:
        for row in rows:
            if not isinstance(row, dict):
                continue
            biz_day = _day(_first(row, ("IO_DATE", "BASE_DATE", "ORDER_DATE", "REG_DATE")))
            party = str(_first(row, ("CUST_DES", "CUST_NAME", "BUSINESS_DES")))[:60]
            amount = _amount(_first(row, ("SUPPLY_AMT", "TOTAL_AMT", "AMT"), None))
            status = str(_first(row, ("STATUS", "PROGRESS_STATUS", "USE_YN")))[:40]
            _rid, how = datalake.put_record(
                con, kind, _natural(endpoint, row), row, biz_date=biz_day,
                party=party, amount=amount, status=status,
                why="이카운트 공식 조회 OAPI", actor="erp_api_collect",
            )
            if how == "new": made += 1
            elif how == "changed": changed += 1
            else: same += 1
        datalake.log(con, "erp", "api_collect", detail={
            "endpoint": endpoint, "rows": len(rows), "new": made,
            "changed": changed, "same": same,
        }, actor="erp_api_collect")
        con.commit()
    finally:
        con.close()
    return {"신규": made, "변경": changed, "그대로": same}


def _validated_rows(client: EcountClient, endpoint: str, body: dict[str, Any]) -> list[dict[str, Any]]:
    rows = client.inquiry(endpoint, body)
    if not isinstance(rows, list):
        raise EcountError(f"{endpoint} 응답이 목록이 아닙니다")
    if len(rows) == 1 and isinstance(rows[0], dict):
        only = rows[0]
        # 업무 행이 아니라 API 오류 봉투만 돌아온 것을 1건 수집으로 세지 않는다.
        if (only.get("Status") is False or only.get("Error") or only.get("Errors")):
            # ★ 봉투를 **열어서** 올린다. 예전에는 '오류 응답을 돌려줬습니다' 한 줄만 남아
            #   리포트가 무엇이 잘못됐는지 말하지 못했다 — 화면에는 실패가 떠 있는데
            #   고칠 자리를 아무도 못 찾는 자리였다.
            code, raw = _envelope(only)
            raw = client.safe(raw)          # 응답이 세션ID 를 되비칠 수 있다
            kind, how = _diagnose(f"{code} {raw}")
            raise ApiTrouble(
                "%s %s%s · %s" % (endpoint, kind, f" (코드 {code})" if code else "", raw[:180]),
                갈래=kind, 조치=how, 코드=code, 원문=raw, endpoint=endpoint)
    return [dict(row) for row in rows if isinstance(row, dict)]


def collect(endpoints: list[str], days: int = 120, force: bool = False) -> dict[str, Any]:
    if not force and _fresh():
        cached = _read(STATUS)
        cached["cached"] = True
        return cached
    cfg = load_config()
    missing = [name for name in endpoints if name not in cfg.get("endpoints", {})]
    if missing:
        raise EcountError("설정에 확인된 endpoint가 없습니다: " + ", ".join(missing))
    client = EcountClient(cfg)
    client.login()
    today = date.today()
    start = today - timedelta(days=max(1, int(days)))
    result: dict[str, Any] = {
        "ok": True, "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "mode": "test" if cfg.get("auth", {}).get("IS_TEST") else "live",
        "sources": {}, "limits": {
            "supported": ["품목 조회", "발주서 조회"],
            "not_available_as_read_api": ["판매/매입전표", "세금계산서", "수금"],
        },
    }
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    broken: dict[str, str] = {}
    for endpoint in endpoints:
        # ★ **이미 받아 둔 갈래는 다시 안 받는다**(위 `_fresh_source` 설명).
        #   `--force` 면 예전 그대로 전부 받는다 — 좁히는 것도 고장이다([172]).
        if not force:
            keep = _fresh_source(endpoint)
            if keep is not None:
                # 조용히 넘기지 않는다([169]) — 이 회차에 새로 받은 것이
                # 아니라는 사실을 화면·리포트가 그대로 말한다.
                keep["cached"] = True
                result["sources"][endpoint] = keep
                continue
        body = {} if endpoint == "items" else {
            "BASE_DATE_FROM": start.strftime("%Y%m%d"),
            "BASE_DATE_TO": today.strftime("%Y%m%d"),
        }
        # ★ 한 갈래가 죽어도 **다른 갈래의 성공을 지우지 않는다.** 예전에는 여기서
        #   그대로 올려서, 품목이 멀쩡히 들어왔는데도 리포트에는 그 사실이 통째로
        #   사라지고 `ok:false` 한 줄만 남았다(성공을 실패로 적는 자리).
        try:
            rows = _validated_rows(client, endpoint, body)
        except Exception as exc:
            fields = getattr(exc, "fields", {})
            broken[endpoint] = fields.get("갈래", "모름")
            result["sources"][endpoint] = {
                "label": CONFIRMED[endpoint][1], "count": None,
                "실패": client.safe(str(exc))[:300],
                "갈래": broken[endpoint], "조치": fields.get("조치", ""),
                "코드": fields.get("코드", ""), "원문": fields.get("원문", ""),
            }
            continue
        payload = {
            "source": "ECOUNT OAPI", "endpoint": endpoint,
            "collected_at": result["generated_at"], "request_range": body,
            "count": len(rows), "rows": rows,
        }
        path = OUT_DIR / f"ecount_{endpoint}_{stamp}.json"
        _atomic_json(path, payload)
        latest = OUT_DIR / f"ecount_{endpoint}_latest.json"
        _atomic_json(latest, payload)
        ingested = _ingest(endpoint, rows)
        result["sources"][endpoint] = {
            "label": CONFIRMED[endpoint][1], "count": len(rows),
            "file": str(path), "ingested": ingested,
        }
    if broken:
        result["ok"] = False
        # 갈래를 **맨 앞**에 둔다 — 읽는 쪽(system_audit)이 220자로 자르는데,
        # 뒤에 두면 정작 필요한 한 마디가 잘려 나간다.
        result["error"] = " · ".join(f"{k} {v}" for k, v in broken.items())
        result["조치"] = " / ".join(
            dict.fromkeys(str(result["sources"][k].get("조치") or "") for k in broken)) or ""
        result["살아남은것"] = [k for k in result["sources"] if k not in broken]
    _atomic_json(STATUS, result)
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--items", action="store_true")
    parser.add_argument("--purchase-orders", action="store_true")
    parser.add_argument("--days", type=int, default=120)
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args(argv)
    selected = []
    if args.items: selected.append("items")
    if args.purchase_orders: selected.append("po_list")
    if not selected: selected = list(CONFIRMED)
    try:
        result = collect(selected, days=args.days, force=args.force)
    except Exception as exc:
        fields = getattr(exc, "fields", {})
        # 클라이언트가 못 만든 실패(설정·로그인)도 밖으로 나가기 전에 지운다.
        msg = f"{type(exc).__name__}: {exc}"
        kind, how = fields.get("갈래"), fields.get("조치")
        # ★ 진단기는 같은 파일에 있는데 이 길에서만 안 불렸다 — 그래서 갈래가
        #   늘 '모름' 이고 읽는 쪽은 'code'(코드가 깨졌다)라 적었다([289]).
        # ★ **첫 줄만** 본다: 그 줄은 우리 클라이언트가 만든 말이고, 뒤에 붙는
        #   남의 HTML 본문을 같이 훑으면 옆 갈래를 삼킨다([172]).
        if not kind or kind == "모름":
            head_line = (msg.splitlines() or [msg])[0]
            kind, how = _diagnose(head_line)
        # ★ 갈래를 **맨 앞**에 둔다 — 이 값은 300자에서 잘리는데(읽는 쪽은 220자)
        #   뒤에 두면 정작 필요한 한 마디가 남의 HTML 에 밀려 사라진다([292]·[325]).
        tag = f"[{kind}] " if kind and kind != "모름" else ""
        failed = {"ok": False, "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
                  "error": (tag + mask_secrets(msg))[:300],
                  "갈래": kind or "모름", "조치": how or "",
                  "원문": fields.get("원문", "")}
        _atomic_json(STATUS, failed)
        print("ERP API 수집 실패:", failed["error"])
        return 1
    detail = " · ".join(
        f"{v['label']} {v['count']:,}건"
        + (" (이미 받음)" if v.get("cached") else "") if v.get("count") is not None
        else f"{v['label']} 실패({v.get('갈래', '모름')})"
        for v in result.get("sources", {}).values())
    print(("ERP API 캐시 재사용" if result.get("cached") else
           "ERP API 수집·DB 반영 완료" if result.get("ok") else "ERP API 일부 실패")
          + (" · " + detail if detail else ""))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
