# -*- coding: utf-8 -*-
"""클라우드 큐 로컬 비밀 설정을 생성·갱신한다.

토큰 원문은 gitignored config/cloud_queue.local.json에만 저장하고 화면에는 출력하지 않는다.
서버에는 SHA-256 해시만 등록하므로 이 스크립트의 출력은 비밀값이 아니다.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import secrets

ROOT = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(ROOT, "config", "cloud_queue.local.json")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", default="")
    args = ap.parse_args()
    try:
        cfg = json.load(open(PATH, encoding="utf-8"))
    except Exception:
        cfg = {}
    cfg.setdefault("enqueue_token", secrets.token_urlsafe(48))
    cfg.setdefault("worker_token", secrets.token_urlsafe(48))
    if args.url:
        if not args.url.startswith("https://"):
            raise SystemExit("HTTPS 주소만 사용할 수 있습니다")
        cfg["api_base_url"] = args.url.rstrip("/")
    os.makedirs(os.path.dirname(PATH), exist_ok=True)
    tmp = PATH + f".{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, PATH)
    print(json.dumps({
        "QUEUE_ENQUEUE_TOKEN_HASH": hashlib.sha256(
            cfg["enqueue_token"].encode("utf-8")).hexdigest(),
        "QUEUE_WORKER_TOKEN_HASH": hashlib.sha256(
            cfg["worker_token"].encode("utf-8")).hexdigest(),
        "url_set": bool(cfg.get("api_base_url")),
    }))


if __name__ == "__main__":
    main()
