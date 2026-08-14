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
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from ecount_client import EcountClient, EcountError, load_config


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


def _fresh(hours: float = 6) -> bool:
    try:
        return (datetime.now().timestamp() - STATUS.stat().st_mtime) < hours * 3600 and bool(_read(STATUS).get("ok"))
    except OSError:
        return False


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
            raise EcountError(f"{endpoint} API가 오류 응답을 돌려줬습니다")
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
    for endpoint in endpoints:
        body = {} if endpoint == "items" else {
            "BASE_DATE_FROM": start.strftime("%Y%m%d"),
            "BASE_DATE_TO": today.strftime("%Y%m%d"),
        }
        rows = _validated_rows(client, endpoint, body)
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
        failed = {"ok": False, "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
                  "error": f"{type(exc).__name__}: {str(exc)[:300]}"}
        _atomic_json(STATUS, failed)
        print("ERP API 수집 실패:", failed["error"])
        return 1
    detail = " · ".join(f"{v['label']} {v['count']:,}건" for v in result.get("sources", {}).values())
    print(("ERP API 캐시 재사용" if result.get("cached") else "ERP API 수집·DB 반영 완료")
          + (" · " + detail if detail else ""))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
