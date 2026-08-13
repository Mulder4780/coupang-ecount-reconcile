# -*- coding: utf-8 -*-
"""Synchronize extracted Band work directly into the canonical app DB.

Excel row capacity and formulas used to decide whether a Band item could be
registered.  That is no longer valid once the application DB is the system of
record: a new work item must be committed even when the archival workbook has
no spare row (``archive_worker`` expands the archive safely later).

The importer is intentionally conservative.  It creates missing AS/inspection
records, fills blank human/legacy fields, and accepts an explicit completion as
an upgrade.  It never turns an existing job into cancelled from the broad Band
extract; cancellation continues to require the stricter cross-signal evidence
gate.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import tempfile
from datetime import date, datetime
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from app_store import AppStore, VersionConflict, canonical_json, default_store


MAP: Dict[str, Dict[str, Any]] = {
    "돌발AS": {
        "prefix": "AS",
        "status_field": "진행상태",
        "fields": {
            "프로젝트NO": "프로젝트NO",
            "캠프명": "캠프명",
            "접수일자": "작업일",
            "담당기사": "담당기사",
            "진행상태": "_진행상태",
            "작업완료일": "_완료일",
            "유상·무상·보험": "비용구분",
            "신청내용": "_내용",
            "비고": "_출처",
        },
    },
    "정기점검": {
        "prefix": "PM",
        "status_field": "점검상태",
        "fields": {
            "프로젝트NO": "프로젝트NO",
            "캠프명": "캠프명",
            "점검예정일": "작업일",
            "담당기사": "담당기사",
            "점검상태": "_진행상태",
            "실제점검일": "_완료일",
            "유상·무상·보험": "비용구분",
            "비고": "_출처",
        },
    },
}


def _text(value: Any) -> str:
    return str(value or "").strip()


def _kind(record: Mapping[str, Any]) -> str:
    raw = _text(record.get("업무유형"))
    if "정기점검" in raw or "정기 점검" in raw:
        return "정기점검"
    if "돌발" in raw or "AS" in raw.upper():
        return "돌발AS"
    return ""


def _work_day(record: Mapping[str, Any]) -> str:
    for raw in (record.get("작업일"), record.get("게시일")):
        text = _text(raw)[:10]
        try:
            date.fromisoformat(text)
            return text
        except ValueError:
            continue
    return date.today().isoformat()


def _source_ref(record: Mapping[str, Any], project: str) -> str:
    band = _text(record.get("밴드") or record.get("밴드명") or record.get("band"))
    posted = _text(record.get("게시일"))
    return "band:" + ":".join(x for x in (band, posted, project) if x)


#: 단계 낱말을 여기 적지 않는다([196]).  관리대장 10_코드관리 드롭다운이 정본이고
#: ``work_flow`` 가 그것을 읽는다.  낱말을 이 파일에 박아 두면 드롭다운이 바뀐 날
#: 이 파일만 옛 낱말을 계속 쓰는데, 그 값은 엑셀에서 '목록 밖'이 되면서도 화면에는
#: 멀쩡히 보인다 — 오류가 나지 않는 종류의 잘못이다.
#:
#: 아래 짝은 work_flow 를 못 읽었을 때만 쓰는 대비값이며 **실측한 정본 낱말**이다.
#: (예전에는 'as' 기본값이 ``접수`` 로 박혀 있었는데 그 낱말은 드롭다운에 없다 —
#:  원장에 78건이 그 상태로 쌓여 있었다.  못 읽었다고 아는 낱말까지 버리지는 않는다.)
_STAGE_FALLBACK: Dict[str, tuple] = {"as": ("신규접수", "작업완료"), "pm": ("예정", "완료")}
_STAGE_CACHE: Dict[str, tuple] = {}


def _stage_words(kind_key: str) -> tuple:
    """(기본단계, 완료단계) — 낱말은 work_flow 한 곳에서 받아온다.

    work_flow 는 워크북을 여는 비싼 일을 하므로 **레코드마다 부르지 않는다**([168]).
    한 회차에 1,600건 넘게 도는 자리라 호출을 프로세스당 한 번으로 묶는다.
    """
    cached = _STAGE_CACHE.get(kind_key)
    if cached is not None:
        return cached
    words = _STAGE_FALLBACK[kind_key]
    try:
        import work_flow

        first = _text(work_flow.default_stage(kind_key))
        last = _text(work_flow.done_stage(kind_key))
        if first and last:
            words = (first, last)
    except Exception:
        # 정의를 못 읽어도 등록 자체는 계속한다 — 낱말 하나 때문에 접수를 잃는 쪽이
        # 더 나쁘다.  대비값이 정본과 같으므로 조용히 틀리지는 않는다.
        pass
    _STAGE_CACHE[kind_key] = words
    return words


def _status(kind: str, record: Mapping[str, Any]) -> str:
    first, last = _stage_words("pm" if kind == "정기점검" else "as")
    raw = _text(record.get("_진행상태") or record.get("진행상태"))
    if "완료" in raw:
        return last
    if "취소" in raw:
        return "취소"
    return first


def _is_done(value: Any) -> bool:
    return "완료" in _text(value)


def _prepare(records: Iterable[Mapping[str, Any]]) -> List[Dict[str, Any]]:
    # Reuse the proven grouping/enrichment rules without invoking its Excel
    # row allocator or writer queue.
    from backfill_rows import dedupe, enrich

    selected = [dict(row) for row in records if _kind(row) and _text(row.get("프로젝트NO"))]
    return [enrich(row) for row in dedupe(selected)]


def sync_records(
    records: Iterable[Mapping[str, Any]],
    *,
    store: Optional[AppStore] = None,
) -> Dict[str, Any]:
    store = (store or default_store()).initialize()
    prepared = _prepare(records)
    existing = store.list_work(limit=10_000)
    by_project: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}
    for work in existing:
        project = _text(work.get("project_no") or (work.get("fields") or {}).get("프로젝트NO"))
        if project:
            by_project.setdefault((_text(work.get("kind")), project), []).append(work)

    result: Dict[str, Any] = {
        "ok": True,
        "source_records": len(prepared),
        "created": 0,
        "updated": 0,
        "unchanged": 0,
        "ambiguous": 0,
        "unsupported": 0,
        "errors": [],
    }
    for row in prepared:
        kind = _kind(row)
        spec = MAP.get(kind)
        project = _text(row.get("프로젝트NO"))
        if not spec or not project:
            result["unsupported"] += 1
            continue
        matches = by_project.get((kind, project), [])
        if len(matches) > 1:
            result["ambiguous"] += 1
            result["errors"].append(f"{kind}/{project}: 앱 DB에 같은 프로젝트가 {len(matches)}건")
            continue

        desired_status = _status(kind, row)
        # Broad extraction may contain a cancellation word.  Creating it as a
        # cancelled record is useful evidence, but an existing live record is
        # never closed here; cancel_watch/cross_signal owns that decision.
        fields = {
            header: row.get(source)
            for header, source in spec["fields"].items()
            if _text(row.get(source))
        }
        fields[spec["status_field"]] = desired_status
        observed_day = _work_day(row)
        observed_at = observed_day + "T00:00:00+09:00"
        source_ref = _source_ref(row, project)
        source_sha = hashlib.sha256(canonical_json(row).encode("utf-8")).hexdigest()
        evidence = f"밴드 자동수집 · {source_ref}"
        meta = {
            key: {
                "source": "band",
                "evidence": evidence,
                "source_ref": source_ref,
                "source_sha256": source_sha,
                "source_observed_at": observed_at,
            }
            for key in fields
        }

        try:
            if not matches:
                reserved = store.reserve_public_id(kind, observed_day, spec["prefix"])
                public_id = reserved["public_id"]
                created = store.create_work(
                    kind=kind,
                    business_key=project,
                    public_id=public_id,
                    project_no=project,
                    camp_name=_text(row.get("캠프명")) or None,
                    status=desired_status,
                    fields=fields,
                    field_meta=meta,
                    actor="automation",
                    source="band",
                    evidence=evidence,
                    source_ref=source_ref,
                    source_sha256=source_sha,
                    source_observed_at=observed_at,
                    idempotency_key=f"band-create:{kind}:{project}:{source_sha}",
                )["work"]
                by_project[(kind, project)] = [created]
                result["created"] += 1
                continue

            current = matches[0]
            current_fields = dict(current.get("fields") or {})
            current_meta = dict(current.get("field_meta") or {})
            patch_fields: Dict[str, Any] = {}
            for key, value in fields.items():
                old = current_fields.get(key)
                old_source = _text((current_meta.get(key) or {}).get("source"))
                strong_completion = (
                    key in {spec["status_field"], "작업완료일", "실제점검일"}
                    and _is_done(desired_status)
                )
                if _text(old) == _text(value):
                    continue
                if not _text(old) or old_source == "band" or strong_completion:
                    patch_fields[key] = value

            current_status = _text(current.get("status"))
            next_status = current_status
            if not current_status or (_is_done(desired_status) and not _is_done(current_status)):
                next_status = desired_status
            # Do not downgrade/close an existing record from this broad signal.
            if desired_status == "취소":
                next_status = current_status
                patch_fields.pop(spec["status_field"], None)

            patch: Dict[str, Any] = {}
            if patch_fields:
                patch["fields"] = patch_fields
                patch["field_meta"] = {key: meta[key] for key in patch_fields}
            if next_status != current_status:
                patch["status"] = next_status
            if not _text(current.get("camp_name")) and _text(row.get("캠프명")):
                patch["camp_name"] = _text(row.get("캠프명"))
            if not patch:
                result["unchanged"] += 1
                continue
            updated = store.update_work(
                current["id"],
                expected_version=int(current["record_version"]),
                patch=patch,
                actor="automation",
                source="band",
                evidence=evidence,
                source_ref=source_ref,
                source_sha256=source_sha,
                source_observed_at=observed_at,
                idempotency_key=f"band-update:{current['id']}:{current['record_version']}:{source_sha}",
            )["work"]
            by_project[(kind, project)] = [updated]
            result["updated"] += 1
        except VersionConflict as exc:
            result["errors"].append(f"{kind}/{project}: 동시수정 충돌({exc}) · 다음 회차 재시도")
        except Exception as exc:
            result["errors"].append(f"{kind}/{project}: {type(exc).__name__}: {exc}")

    result["ok"] = not result["errors"]
    return result


def self_test() -> bool:
    with tempfile.TemporaryDirectory(prefix="band-canonical-") as temp:
        store = AppStore(os.path.join(temp, "app.db")).initialize()
        base = {
            "프로젝트NO": "UJ-BAND-001",
            "업무유형": "돌발 AS",
            "캠프명": "합성캠프",
            "작업일": "2026-08-10",
            "게시일": "2026-08-10",
            "담당기사": "김필우",
            "진행상태": "접수·예정",
            "문서상태": "사진 확인",
            "비용구분": "무상",
            "밴드": "합성밴드",
        }
        first = sync_records([base], store=store)
        assert first["ok"] and first["created"] == 1, first
        seq = store.status()["change_seq"]
        second = sync_records([base], store=store)
        assert second["ok"] and second["unchanged"] == 1, second
        assert store.status()["change_seq"] == seq
        done = sync_records([{**base, "진행상태": "작업완료"}], store=store)
        assert done["ok"] and done["updated"] == 1, done
        work = store.list_work(kind="돌발AS")[0]
        assert work["status"] == "작업완료" and work["fields"]["작업완료일"] == "2026-08-10"
        cancelled = sync_records([{**base, "진행상태": "취소"}], store=store)
        assert cancelled["ok"] and store.list_work(kind="돌발AS")[0]["status"] == "작업완료"
    return True


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args(argv)
    if args.self_test:
        self_test()
        print("band_canonical self-test: OK")
        return 0
    from band_extract import load_records

    result = sync_records(load_records())
    print(json.dumps(result, ensure_ascii=False, indent=1))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
