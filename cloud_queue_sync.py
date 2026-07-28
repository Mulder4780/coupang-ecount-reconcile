# -*- coding: utf-8 -*-
"""클라우드 입력 큐를 안전하게 관리대장에 반영한다.

PC가 꺼져 있는 동안에는 서버가 입력을 영구 보관한다. PC가 켜지면 이 스크립트가
항목을 임대(lease)하고, 기존 project_resolve + ledger_writer ZIP 패치 경로로 반영한
뒤 성공한 항목만 확인(ack)한다. 실패 항목은 다시 대기 상태로 돌려 데이터 유실을 막는다.
"""
from __future__ import annotations

import json
import os
import sys
import urllib.error
import urllib.request
from datetime import datetime

from operation_window import input_window_label, is_input_window

ROOT = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(ROOT, "config", "cloud_queue.local.json")
REPORT = os.path.join(ROOT, "reports", "cloud_queue_sync.json")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def load_config(path=CONFIG):
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    url = str(cfg.get("api_base_url", "")).rstrip("/")
    token = str(cfg.get("worker_token", ""))
    if not url.startswith("https://") or len(token) < 32:
        raise ValueError("클라우드 큐 주소 또는 작업자 토큰 설정이 올바르지 않습니다")
    return url, token


def api(url, token, path, body=None, timeout=20):
    data = None if body is None else json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        url + path,
        data=data,
        method="GET" if body is None else "POST",
        headers={
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json",
            "User-Agent": "CSOS-Local-Sync/1",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def save_report(data):
    os.makedirs(os.path.dirname(REPORT), exist_ok=True)
    tmp = REPORT + f".{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump({"time": datetime.now().isoformat(), **data}, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, REPORT)


def sync_once():
    if is_input_window():
        return {"ok": True, "skipped": f"입력 보호시간({input_window_label()})"}
    if not os.path.exists(CONFIG):
        return {"ok": True, "skipped": "클라우드 큐 설정 없음"}

    import ai_claim

    if not ai_claim.take("cloud-sync", "ledger", "휴대폰 클라우드 예약 반영"):
        return {"ok": True, "skipped": "관리대장 작업 점유 중"}

    lease_id = ""
    leased_ids = []
    try:
        url, token = load_config()
        lease = api(url, token, "/api/queue/lease", {"limit": 50, "leaseMinutes": 20})
        lease_id = str(lease.get("leaseId", ""))
        items = lease.get("items") or []
        leased_ids = [int(item["id"]) for item in items]
        if not items:
            return {"ok": True, "leased": 0, "applied": 0}

        import project_resolve as resolver
        from webapp.app_server import enqueue_codes

        evidence = resolver.evidence()
        already, ready, retry = [], [], []
        for item in items:
            resolved = resolver.resolve(item["code"], evidence)
            if not resolved.get("ok"):
                retry.append((int(item["id"]), str(resolved.get("reason", "프로젝트 미확인"))))
            elif resolved.get("state") == "등록됨":
                already.append(int(item["id"]))
            else:
                ready.append(item)

        acknowledged = list(already)
        if ready:
            result = enqueue_codes([item["code"] for item in ready])
            if result.get("ok"):
                acknowledged.extend(int(item["id"]) for item in ready)
            else:
                retry.extend((int(item["id"]), "관리대장 반영 실패") for item in ready)

        if acknowledged:
            api(url, token, "/api/queue/ack", {"leaseId": lease_id, "ids": acknowledged})
        if retry:
            retry_ids = [item_id for item_id, _ in retry]
            reason = "; ".join(dict.fromkeys(reason for _, reason in retry))[:500]
            api(
                url,
                token,
                "/api/queue/release",
                {"leaseId": lease_id, "ids": retry_ids, "error": reason},
            )
        return {
            "ok": not retry,
            "leased": len(items),
            "applied": len(acknowledged) - len(already),
            "already": len(already),
            "retry": len(retry),
        }
    except (OSError, ValueError, urllib.error.URLError, json.JSONDecodeError) as exc:
        if lease_id and leased_ids:
            try:
                url, token = load_config()
                api(
                    url,
                    token,
                    "/api/queue/release",
                    {"leaseId": lease_id, "ids": leased_ids, "error": str(exc)[:500]},
                )
            except Exception:
                pass
        return {"ok": False, "error": str(exc)[:500]}
    finally:
        ai_claim.free("cloud-sync", "ledger")


def main():
    result = sync_once()
    save_report(result)
    print(json.dumps(result, ensure_ascii=False))
    if not result.get("ok"):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
