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
import io
import json
import os
import sys
import tempfile
from datetime import date, datetime
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from app_store import AppStore, VersionConflict, canonical_json, default_store


#: 모호는 코드가 못 푼다 — **무엇을 해야 하는지** 같이 적는다([289]).
#: ★ 2026-08-25 실측: `erp_pick` 이 돌려주는 다섯 문구는 **왜 모호한지**만
#:   말하고 조치가 없었다. 그러면 읽는 사람은 "그래서 뭘 하라는 거지" 로 끝난다.
#:   조치는 갈래마다 다르므로 그 말이 없으면 자국이 반쪽이다.
AMBIGUOUS_FIX = "앱에서 그 프로젝트를 열어 사람이 정한다"

#: ERP 자료가 **아직 안 닿은 구간** — 그것은 "등록이 안 됐다" 가 아니다
#: (2026-09-02 실사고).  실측: 색인이 담은 마지막 전표일이 **2026-08-05** 인데
#: UJ2601393(완료 8/10) · UJ2601416(예정 8/6)은 그 뒤였다.  그런데 문구가
#: "ERP 색인에 이 프로젝트가 없다" 하나뿐이라 읽는 사람은 *"류지영이 ERP 에
#: 등록을 안 했나"* 로 읽는다 — **틀린 지목은 못 잡는 것보다 나쁘다**([172]).
#: 조치가 정반대다([289]): 앞은 "ERP 에 등록해 주세요", 뒤는
#: **"판매조회를 그 뒤 기간으로 다시 받는다"** 이고 그러면 스스로 풀린다.
#: ⚠ 문구를 만드는 곳과 조치를 고르는 곳이 **이 상수 하나**를 본다([162]) —
#:   손으로 적으면 문구를 다듬는 날 그 갈래가 조용히 죽는다([165]).
ERP_BEHIND_MARK = "ERP 자료가 "
ERP_BEHIND_FIX = ("ERP 판매조회를 그 뒤 기간으로 다시 받는다"
                  " — 들어오면 이 건은 스스로 풀린다(사람이 고를 것 없다)")


def ambiguous_fix(why):
    """왜 모호한지에 따라 조치가 갈린다([289]) — 문구를 지어내지 않는다."""
    return ERP_BEHIND_FIX if _text(why).startswith(ERP_BEHIND_MARK) else AMBIGUOUS_FIX

#: 모호한 건(같은 프로젝트에 행이 여럿)을 남기는 자리.  **실패 자국이 아니다** —
#: 이름을 `*_오류.json` 으로 두면 `schedule_watch.traces()` 가 실패로 모아 다시
#: 경보가 된다([170]).  사람이 정할 것을 모아 두는 자리다.
AMBIGUOUS_TRACE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               "reports", "밴드등록_모호.json")

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


_ERP_IDX = None
ERP_INDEX = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "reports", "ERP판매_프로젝트색인.json")


def _erp_index():
    """ERP 판매 프로젝트 색인 — **회차가 만들어 둔 파일만** 읽는다.

    여기서 `erp_sales_index.build()` 를 부르면 Z: 를 재귀로 훑어(실측 수십 초 · [409])
    5분 회차가 그만큼 느려진다([168] — 비싼 탐색은 캐시 검사 뒤에).
    ★ 못 읽으면 **빈 것**이다 — 그러면 아래 문이 하나도 안 열려 예전과 똑같이
      '사람이 정한다' 로 남는다([169] — 못 읽은 것을 '가릴 수 있다' 로 치지 않는다).
    """
    global _ERP_IDX
    if _ERP_IDX is None:
        try:
            with open(ERP_INDEX, encoding="utf-8") as fh:
                _ERP_IDX = json.load(fh).get("index") or {}
        except Exception:
            _ERP_IDX = {}
    return _ERP_IDX


def _erp_day(value):
    """ERP `date` 는 `20260810-4`(일자-No.) 다 — 앞 여덟 자리만 날짜다."""
    digits = "".join(ch for ch in _text(value) if ch.isdigit())[:8]
    if len(digits) != 8:
        return ""
    return "%s-%s-%s" % (digits[:4], digits[4:6], digits[6:])


def _done_day(work):
    """그 행의 작업완료일 — `YYYY-MM-DD` 앞 열 글자만 본다."""
    fields = work.get("fields") or {}
    for key in ("작업완료일", "점검완료일", "완료일"):
        day = _text(fields.get(key))[:10]
        if len(day) == 10:
            return day
    return ""


def _erp_last_day():
    """ERP 색인이 담은 **마지막 전표일** — 그 뒤 건은 없는 것이 아니라 못 받은 것이다.

    ★ 색인을 다시 만들지 않는다([168]) — 이미 읽어 둔 것을 훑을 뿐이다(1822건 · 밀리초).
    ★ **캐시하지 않는다** — 검증이 `_ERP_IDX` 를 갈아 끼우는데 여기만 옛 값을 들고
      있으면 그 검사는 **아무것도 안 재면서 통과한다**([371]).  이 함수는 모호한
      건에서만 불리므로(하루 몇 건) 캐시가 필요 없다.
    ★ 못 재면 `""` 다([169]) — 그러면 아래 갈래가 안 열려 예전 문구 그대로다.
    """
    best = ""
    for entry in _erp_index().values():
        day = _erp_day((entry or {}).get("date"))
        if day > best:
            best = day
    return best


def _latest_day(matches):
    """그 건의 아는 날짜 중 **가장 늦은 것** — 전표는 일이 끝난 뒤에 끊긴다.

    ★ 완료일만 보면 안 된다: 정기점검은 완료 칸이 비고 `점검예정일` 만 있는 행이
      실재한다(실측 UJ2601416 · 두 행 다 점검상태만 "완료").  그러면 갈래가 안 열려
      예전 문구로 떨어지고, 사람은 다시 엉뚱한 데를 본다([172]).
    """
    best = ""
    for work in matches or ():
        fields = (work or {}).get("fields") or {}
        for key in ("작업완료일", "점검완료일", "완료일",
                    "점검예정일", "접수일자", "작업일"):
            day = _text(fields.get(key))[:10]
            if len(day) == 10 and day > best:
                best = day
    return best


def erp_pick(project, matches):
    """ERP 가 가려 주면 그 행을, 아니면 `(None, 왜)` 를 돌려준다.

    형님 지시(2026-08-25): **"erp 기준으로 판단해"**

    ★ **원본이 말하게 한다**([170] 의 유형E 와 같은 자리). 실측 2026-08-25:
      UJ2601393·UJ2601394 는 앱 DB 에 행이 **2건**인데 ERP 판매전표는 **각각 1건**이다
      — 곧 하나가 중복이다. 그 사실만으로도 사람의 조치가 달라진다([289]):
      '어느 행인지 앱에서 정한다' 가 아니라 **'중복으로 보인다'** 다.

    ★ 문은 좁다([172]) — 잘못 고르면 **엉뚱한 행에 완료가 박히고** 그것은
      되돌리기 어렵다. 둘 다 맞을 때만 고른다:
        ① ERP 판매전표가 **정확히 1건**
        ② 그 전표 날짜와 같은 **완료일을 가진 행이 정확히 하나**
      실측 UJ2601393 이 그 자리다(ERP `20260810-4` = `AS-2608-603` 완료일 2026-08-10).
      UJ2601394 는 두 행이 **완전히 같아**(캠프·접수일·신청내용·완료일 없음) 못 고른다 —
      그때는 **안 고르고 근거를 적는다**([169]).
    """
    entry = _erp_index().get(_text(project).upper())
    if not entry:
        last, day = _erp_last_day(), _latest_day(matches)
        if last and day and day > last:
            fmt = "%s%s 까지만 들어왔다 — 이 건(%s)은 그 뒤라 아직 못 받은 것이다"
            fmt += " (ERP 에 등록이 안 된 것이 아니다)"
            return None, fmt % (ERP_BEHIND_MARK, last, day)
        return None, "ERP 색인에 이 프로젝트가 없다"
    rows = entry.get("rows")
    if rows != 1:
        return None, "ERP 판매전표가 %s건이라 1건으로 못 좁힌다" % rows
    day = _erp_day(entry.get("date"))
    if not day:
        return None, "ERP 판매전표 1건인데 그 날짜를 못 읽었다"
    hit = [w for w in matches if _done_day(w) == day]
    if len(hit) == 1:
        return hit[0], "ERP 판매전표 1건 · 그 날짜(%s)와 맞는 행이 하나다" % day
    return None, ("ERP 판매전표는 **1건**인데 앱 DB 는 %d건이다 — 하나가 중복으로 보인다"
                  " (ERP 날짜 %s 와 맞는 행 %d개)" % (len(matches), day, len(hit)))


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
    # ★ **사람이 지운 건까지 한 번에 받는다** (2026-08-31 실사고).  옛 코드는
    #   살아 있는 것만 받아서, 사람이 앱에서 지운 건이 `matches` 에 안 잡히고
    #   **매 회차 새로 만들려 했다.**  그러면 DB 가 유니크 제약으로 거부하고
    #   (`duplicate public_id or kind/business_key`) 그 실패가 `errors` 로 가서
    #   `ok=False` 가 되고, 그러면 파이프라인이 **지문을 안 적어** 같은 자료를
    #   처음부터 다시 처리한다([365]).  실측: 정기점검 6건이 8/28 에 지워진 뒤
    #   **3일 · 56회** 그 고리를 돌았고, 6시간 창에서 회차가 기계를 **70%**
    #   물고 있었다(30회 262분).  앱이 느려지고 아침 관문이 25분에 걸려
    #   **그날 대조가 통째로 안 돌았다.**
    # ⚠ `by_project` 에는 **살아 있는 것만** 담는다 — 지워진 행이 거기 들어가면
    #   `update_work` 가 **지워진 기록을 되살린다**(되돌릴 수 없는 쪽 · [172]).
    existing = store.list_work(limit=10_000, include_deleted=True)
    by_project: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}
    deleted_keys: set = set()
    for work in existing:
        project = _text(work.get("project_no") or (work.get("fields") or {}).get("프로젝트NO"))
        if not project:
            continue
        key = (_text(work.get("kind")), project)
        if work.get("deleted_at"):
            deleted_keys.add(key)
            continue
        by_project.setdefault(key, []).append(work)

    result: Dict[str, Any] = {
        "ok": True,
        "source_records": len(prepared),
        "created": 0,
        "updated": 0,
        "unchanged": 0,
        "ambiguous": 0,
        "unsupported": 0,
        # ★ **사람이 정해야 하는 것은 실패가 아니다**([170]).  전에는 이것을
        #   `errors` 에 담아 `ok=False` 로 만들었고, 그러면 5분 회차가
        #   **하루 288번** 실패로 적혔다 — 실측 2026-08-20: 한 건
        #   (UJ2601394 · 목포1 에 AS-2608-601·604 두 행)이 매 회차를
        #   빨갛게 만들고 있었다.  나머지 1,702건은 멀쩡히 처리됐다.
        #   경보가 늘 켜져 있으면 진짜 실패가 그 안에 묻힌다.
        #   ⚠ 그렇다고 **조용히 넘기지도 않는다**([169]) — 아래
        #   `AMBIGUOUS_TRACE` 에 자국을 남기고, 없어지면 지운다([228]).
        "모호": [],
        # ★ **사람이 지운 건은 실패가 아니다**([170]) — 지운 것은 사람의 결정이고
        #   회차가 되살리면 그것이 사고다.  조용히 넘기지도 않는다([169]).
        "지운건": 0,
        "지운건목록": [],
        # ERP 가 가려 준 것 — 왜 그렇게 골랐는지 근거를 같이 남긴다
        "erp선택": 0,
        "erp근거": [],
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
        if not matches and (kind, project) in deleted_keys:
            # 사람이 앱에서 지운 건이다.  다시 만들지 않는다 — 되살리려면
            # 앱에서 사람이 되살린다(`app_store.undelete_work`).
            result["지운건"] += 1
            result["지운건목록"].append(
                "%s/%s: 사람이 앱에서 지운 건 — 다시 만들지 않는다"
                % (kind, project))
            continue
        if len(matches) > 1:
            # ★ **ERP 가 가른다**(2026-08-25 형님 지시 "erp 기준으로 판단해").
            #   원본이 직접 말한 것이라 근거가 세다 — 짐작으로 고르지 않는다.
            picked, why = erp_pick(project, matches)
            if picked is not None:
                result["erp선택"] += 1
                result["erp근거"].append(
                    "%s/%s: %s -> %s" % (kind, project, why,
                                        picked.get("public_id") or "?"))
                matches = [picked]
            else:
                result["ambiguous"] += 1
                result["모호"].append(
                    f"{kind}/{project}: 앱 DB에 같은 프로젝트가 {len(matches)}건"
                    f" — {why} · {ambiguous_fix(why)}")
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


def write_ambiguous_trace(result, path=None):
    """모호한 건을 자국으로 남긴다 — 실패로 세지 않는 대신 **잃지도 않는다**([169]).

    없어지면 지운다([228]) — 옛 자국이 남으면 이미 풀린 것을 계속 보고한다.
    자국을 못 남겨도 회차를 세우지 않는다(그것은 실패가 아니다).
    """
    p = AMBIGUOUS_TRACE if path is None else path
    rows = list(result.get("모호") or [])
    try:
        if not rows:
            if os.path.exists(p):
                os.remove(p)
            return False
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with io.open(p, "w", encoding="utf-8", newline="") as fh:
            json.dump({"적은때": datetime.now().isoformat(timespec="seconds"),
                       "건수": len(rows), "모호": rows,
                       "뜻": "실패가 아니다 — 같은 프로젝트에 행이 여럿이라"
                             " 어느 행을 고칠지 원본이 말해 주지 않는다. 사람이 정한다."},
                      fh, ensure_ascii=False, indent=1)
        return True
    except Exception:
        return False


def main(argv: Optional[Sequence[str]] = None) -> int:
    # ★ 콘솔은 cp949 라 `—` 한 글자에 죽고, 무인 회차는 `sys.stdout` 이 **None** 이다
    #   ([235]). 인계 문서가 알려 주는 그 명령(`--ambiguous`)이 실측으로 여기서
    #   `UnicodeEncodeError` 로 죽었다 — 사람이 붙여넣으면 그 오류만 본다.
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--ambiguous", action="store_true",
                        help="모호 자국을 사람 말로 찍는다 — 읽기 전용(등록하지 않는다)")
    args = parser.parse_args(argv)
    if args.self_test:
        self_test()
        print("band_canonical self-test: OK")
        return 0
    if args.ambiguous:
        # ★ 읽기 전용이다 — `sync_records()` 를 부르지 않는다.  물어봤을 뿐인데
        #   앱 DB 가 바뀌면 안 된다([181] 과 같은 규칙).
        if not os.path.exists(AMBIGUOUS_TRACE):
            print("밴드 앱 DB 등록 — 사람이 정할 것 없음(모호 자국이 없다).")
            return 0
        try:
            with io.open(AMBIGUOUS_TRACE, encoding="utf-8") as fh:
                d = json.load(fh)
        except Exception as exc:
            print("모호 자국을 못 읽었다 — %s: %s" % (type(exc).__name__, exc))
            print("  ('모호가 없다'는 뜻이 아니다: %s)" % AMBIGUOUS_TRACE)
            return 1
        rows = list(d.get("모호") or [])
        print("밴드 앱 DB 등록 — 사람이 정할 것 %d건 (적은때 %s)"
              % (len(rows), d.get("적은때") or "모름"))
        print("  실패가 아니다 — 같은 프로젝트에 앱 DB 행이 여럿이라")
        print("  어느 행을 고칠지 원본이 말해 주지 않는다.  앱에서 그 프로젝트를 열어 사람이 정한다.")
        for r in rows:
            print("  - %s" % r)
        return 0
    from band_extract import load_records

    result = sync_records(load_records())
    print(json.dumps(result, ensure_ascii=False, indent=1))
    write_ambiguous_trace(result)
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
