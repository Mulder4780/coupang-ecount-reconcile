# -*- coding: utf-8 -*-
"""접수취소·유선 원격해결을 앱 DB 정본에 안전하게 반영한다.

밴드 글의 명시적 프로젝트번호와 접수취소 문구를 근거로 원천 AS/정기점검만
갱신한다. 카톡 교차근거는 있으면 함께 남기지만, 모호한 카톡 한 줄만으로 업무를
취소하지 않는다. 실제 발행된 거래서류·금액은 절대 지우지 않는다.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
from datetime import datetime
from typing import Any, Dict, Iterable, Mapping, Optional


_REMOTE_CHANNEL = re.compile(
    r"(유\s*선\s*(?:전|접)?\s*화|전\s*화|통\s*화|원\s*격)"
)
_REMOTE_DONE = re.compile(
    r"(해\s*결(?:\s*완\s*료)?|정\s*상\s*(?:작\s*동|동\s*작)|"
    r"작\s*동\s*(?:원\s*활|정\s*상)|조\s*치\s*완\s*료)"
)
_CANCELLED = ("취소", "철회")
_OPEN_STATES = {
    "", "접수", "접수완료", "배정대기", "배정완료", "방문예정", "진행중", "작업중",
    "미실시", "예정", "예정월", "대기", "미점검",
}

# ★ 여기 적힌 낱말만 보면 **정본을 모른다**([187] · 2026-08-21 실측).
#   관리대장 10_코드관리 의 돌발AS 목록은 `신규접수`(10건)·기사배정·일정확정·
#   방문중·재방문예정·보류 이고 정기점검에는 `일정변경` 이 있는데, 위 표에는
#   그중 하나도 없다. 그래서 **열린 행**이 `not in _OPEN_STATES` 로 떨어져
#   "완료일 또는 종료상태가 있어" 라는 **거짓 사유**를 리포트에 적었다
#   (실측 그날 충돌 4건 중 2건이 그것 — UJ2601444·UJ2601445 둘 다 `신규접수`).
#   경보의 절반이 가짜면 나머지도 아무도 안 본다([170]).
#
#   그래서 낱말은 **정본 한 곳**(work_flow.definition)에서 가져온다([162]).
#   여기 손으로 베껴 적으면 관리대장 드롭다운이 바뀐 날 사본만 낡는다.
#   위 표는 **지우지 않고 합집합으로 남긴다** — 원장에 이미 쓰인 옛 값이 있다
#   (`접수` 80건 · `예정월` 69건은 목록 밖 값인데 둘 다 열린 뜻이다).
#
# ⚠ 정본은 '이 단계가 끝인가'를 말해 주지 않는다 — 완료단계와 취소류만 안다.
#   `AS전환` 은 목록에 있지만 **그 정기점검 건은 끝난 것**이다(AS 로 넘어갔다).
#   열린 것으로 세면 이미 넘어간 건을 접수취소로 뒤집을 수 있다([172]).
#   그러므로 이 한 낱말만 손으로 적고, 왜 적었는지를 여기 남긴다.
_CLOSED_EXTRA = {"AS전환"}


def open_states():
    """열린 상태 낱말 — 정본 + 위 표의 합집합. `(낱말들, 못읽은이유)`.

    ★ 못 읽으면 '전부 열림'으로도 '전부 닫힘'으로도 치지 않는다([169]).
      위 표로 물러나고 **못 읽었다는 사실을 같이 돌려준다** — 부르는 쪽이
      그것을 사유에 적어야 사람이 왜 이렇게 판정됐는지 알 수 있다.
    """
    base = set(_OPEN_STATES)
    try:
        import work_flow
        kinds = (work_flow.definition() or {}).get("갈래") or {}
    except Exception as exc:
        return base, "%s: %s" % (type(exc).__name__, str(exc)[:80])
    if not kinds:
        return base, "정본 갈래가 비어 있다"
    for g in kinds.values():
        if not isinstance(g, dict):
            continue
        done = str(g.get("완료단계") or "").strip()
        for s in g.get("단계") or []:
            name = str((s or {}).get("단계") or "").strip()
            if not name or name == done or name in _CLOSED_EXTRA:
                continue
            if any(word in name for word in _CANCELLED):
                continue
            base.add(name)
    return base, ""


def known_states():
    """정본에 **적혀 있는** 상태 낱말 전부(열림·닫힘 가리지 않음).

    ★ '종료상태다' 와 '모르는 낱말이다' 는 다른 사실이다([169]).
      정본 목록에 있는데 열림이 아니면 그것은 진짜 종료상태이고,
      목록에 아예 없으면 **우리가 모르는 것**이다 — 종료라고 확언하면
      사람이 없는 완료 기록을 찾아 나선다([172]).
    """
    out = set()
    try:
        import work_flow
        kinds = (work_flow.definition() or {}).get("갈래") or {}
    except Exception:
        return out
    for g in kinds.values():
        if not isinstance(g, dict):
            continue
        for s in g.get("단계") or []:
            name = str((s or {}).get("단계") or "").strip()
            if name:
                out.add(name)
    return out


def closed_reason(status, done_value, open_words, unread=""):
    """왜 자동 취소를 안 했는지 **갈래를 갈라** 말한다([289]).

    예전에는 셋을 한 문장으로 뭉쳤다 — "완료일 또는 종료상태가 있어".
    그러면 `신규접수` 인 행에도 그 말이 붙어 사람이 **없는 완료일을 찾아
    나선다**([172]). 조치가 갈래마다 다르므로 문장도 갈라야 한다.
    """
    status = str(status or "").strip()
    if str(done_value or "").strip():
        return "완료일이 이미 있어 자동 취소하지 않음"
    if any(word in status for word in _CANCELLED):
        return "이미 취소 상태라 자동 취소하지 않음"
    if unread:
        return ("상태 낱말이 열린 단계인지 확인 못 함(정본을 못 읽었다: %s) — %s"
                % (unread, status or "(빈칸)"))
    if status not in open_words:
        if status and status in known_states():
            return "종료상태(%s)라 자동 취소하지 않음" % status
        # 정본 목록에 없는 낱말이다 — 끝난 것인지 아닌지 **모른다**([169]).
        return ("상태 낱말 '%s' 이 관리대장 목록에 없어 열린 단계인지 판단할 수 없음"
                % (status or "(빈칸)"))
    return "자동 취소하지 않음"

SOURCE_SPECS = (
    ("돌발AS", "02_돌발AS접수", "접수ID", "진행상태"),
    ("정기점검", "04_정기점검", "점검ID", "점검상태"),
)


def remote_resolution_hit(text: Any) -> bool:
    """유선·전화·통화로 현장 방문 없이 해결됐다는 명시 문구인가.

    실제 밴드 오타 ``유선접화``도 수용한다. 단 채널 낱말만 있거나 해결 낱말만
    있으면 원격해결로 확정하지 않는다.
    """

    value = str(text or "")
    return bool(_REMOTE_CHANNEL.search(value) and _REMOTE_DONE.search(value))


def outcome_kind(text: Any) -> str:
    """명시적 접수취소 문구를 ``원격해결`` 또는 ``접수취소``로 분류한다."""

    from band_extract import cancel_hit

    value = str(text or "")
    if not cancel_hit(value):
        return ""
    return "원격해결" if remote_resolution_hit(value) else "접수취소"


def _as_day(value: Any) -> str:
    if value in (None, "", 0, "0"):
        return ""
    try:
        number = int(value)
        if number > 10_000_000_000:
            return datetime.fromtimestamp(number / 1000).strftime("%Y-%m-%d")
        if number > 1_000_000_000:
            return datetime.fromtimestamp(number).strftime("%Y-%m-%d")
    except (TypeError, ValueError, OSError, OverflowError):
        pass
    match = re.search(r"\d{4}-\d{2}-\d{2}", str(value))
    return match.group(0) if match else ""


def load_corroborations(path: Optional[str] = None) -> Dict[str, Dict[str, Any]]:
    """최신 카톡·밴드 교차 결과를 프로젝트번호로 읽는다(읽기 전용)."""

    if path is None:
        root = os.path.dirname(os.path.abspath(__file__))
        report_dir = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(root, "reports")
        path = os.path.join(report_dir, "카톡_밴드_교차.json")
    try:
        with open(path, encoding="utf-8") as handle:
            payload = json.load(handle)
    except (OSError, ValueError, TypeError):
        return {}
    out: Dict[str, Dict[str, Any]] = {}
    for row in payload.get("짝지어짐") or []:
        project = str(row.get("프로젝트NO") or "").strip().upper()
        if project:
            out[project] = dict(row)
    return out


def source_outcomes(store=None) -> Dict[str, Dict[str, Any]]:
    """원천업무 ID별 현재 취소·원격해결 사실을 앱 DB에서 읽는다."""

    if store is None:
        from app_store import default_store

        store = default_store()
    out: Dict[str, Dict[str, Any]] = {}
    for kind, sheet, id_col, status_col in SOURCE_SPECS:
        for row in store.list_sheet_rows(sheet):
            source_id = str(row.get(id_col) or row.get("_business_key") or "").strip()
            if not source_id:
                continue
            status = str(row.get(status_col) or "").strip()
            reason = str(row.get("접수취소사유") or "").strip()
            treatment = str(row.get("처리구분") or "").strip()
            cancelled = any(word in status for word in _CANCELLED) or (
                str(row.get("접수취소여부") or "").strip() == "예"
            )
            # ★ 취소와 '청구 제외'는 다른 사실이다 (2026-08-13 류지영 요청).
            #   취소 = 현장에 아무도 안 갔다. 제외 = **다녀왔는데** 이 건으로는 청구를
            #   안 한다(정기점검과 동시 진행 등). 실측 UJ2601032 는 작업완료·완료일
            #   2026-06-12·담당기사 김필우·사진·완료보고서·ERP등록까지 다 있는 건이라
            #   취소로 적으면 **다녀온 사실이 기록에서 뒤집힌다**. 그래서 낱말을 가른다.
            #   청구에서 빠지는 결과는 같으므로 **전파 경로는 하나를 같이 쓴다**([162]) —
            #   갈리는 것은 '왜 빠졌나'뿐이고 그것을 화면이 그대로 말한다.
            excluded = str(row.get("청구제외") or "").strip() == "예"
            out[source_id] = {
                "source_id": source_id,
                "kind": kind,
                "project_no": str(row.get("프로젝트NO") or "").strip().upper(),
                "status": status,
                "cancelled": cancelled,
                "billing_excluded": excluded,
                "excluded_reason": str(row.get("청구제외사유") or "").strip(),
                "reason": reason,
                "treatment": treatment,
                "evidence": str(row.get("접수취소근거") or row.get("_evidence") or "").strip(),
                "confirmed_on": str(row.get("접수취소확인일") or "")[:10],
            }
    return out


def settlement_document_evidence(store=None) -> Dict[str, list[str]]:
    """정산ID별 실제 청구·발행·입금 근거를 exact 관계키로 모은다.

    15/16 sidecar는 06 정산행과 별도 레코드다. 06만 보면 실제 발행·입금이
    존재하는 취소건을 자료 없는 건으로 오판해 확인필요에서 지울 수 있다.
    관리ID/기한관리 placeholder만으로는 실제 서류라고 확정하지 않고, 날짜·번호·
    0보다 큰 금액처럼 객관 사실이 있는 레코드만 센다.
    """

    if store is None:
        from app_store import default_store

        store = default_store()

    out: Dict[str, list[str]] = {}

    def add(settle_id: Any, label: str) -> None:
        key = str(settle_id or "").strip()
        if key:
            out.setdefault(key, []).append(label)

    def present(value: Any) -> bool:
        return str(value or "").strip() not in ("", "0", "0.0")

    for row in store.list_sheet_rows("06_거래서류청구수금"):
        sid = row.get("정산ID") or row.get("_business_key")
        for field in (
            "거래명세서번호", "거래명세서발행일", "세금계산서발행일",
            "세금계산서승인번호", "PO번호", "PO발행일", "청구일", "입금일", "입금액",
        ):
            if present(row.get(field)):
                add(sid, f"06:{field}")

    for row in store.list_sheet_rows("15_세금계산서관리"):
        sid = row.get("정산ID")
        for field in ("실제발행일", "승인번호", "세금계산서승인번호", "발행금액"):
            if present(row.get(field)):
                add(sid, f"15:{field}")

    for row in store.list_sheet_rows("16_입금수금관리"):
        sid = row.get("정산ID")
        for field in ("입금일", "입금액", "은행거래키"):
            if present(row.get(field)):
                add(sid, f"16:{field}")

    return {key: sorted(set(values)) for key, values in out.items()}


def _hit_evidence(hit: Mapping[str, Any], corroboration: Mapping[str, Any]) -> str:
    band = str(hit.get("밴드") or "").strip()
    post = str(hit.get("게시글") or "").strip()
    where = str(hit.get("자리") or "").strip()
    basis = str(hit.get("근거") or "").strip()
    parts = [f"밴드 {band}/{post} {where} — {basis}".strip()]
    if corroboration:
        kakao = str(corroboration.get("카톡") or "").strip()
        kakao_text = re.sub(r"\s+", " ", str(corroboration.get("카톡글") or "")).strip()
        parts.append(f"카톡 교차 {kakao} — {kakao_text[:180]}".strip())
    return " · ".join(part for part in parts if part)[:900]


def sync_hits(
    hits: Mapping[str, Mapping[str, Any]],
    *,
    store=None,
    corroborations: Optional[Mapping[str, Mapping[str, Any]]] = None,
    actor: str = "automation:cancel-watch",
) -> Dict[str, Any]:
    """명시 취소 근거를 원천 AS/정기점검에 멱등·낙관잠금으로 반영한다.

    프로젝트와 업무종류가 정확히 한 행으로 연결될 때만 쓴다. 이미 취소인 행도
    사유·근거가 비어 있으면 보강한다. 정산·계산서·PO·입금 레코드는 건드리지 않는다.
    """

    if store is None:
        from app_store import default_store

        store = default_store()
    corroborations = dict(corroborations or {})
    by_project: Dict[str, list] = {}
    for kind, _sheet, _id_col, status_col in SOURCE_SPECS:
        for work in store.list_work(kind=kind, limit=10_000):
            project = str(work.get("project_no") or "").strip().upper()
            if project:
                by_project.setdefault(project, []).append((kind, status_col, work))

    result = {"ok": True, "updated": 0, "unchanged": 0, "ambiguous": 0,
              "missing": 0, "conflicts": 0, "errors": [], "records": []}
    for project, raw_hit in sorted(hits.items()):
        hit = dict(raw_hit or {})
        project = str(project or hit.get("프로젝트NO") or "").strip().upper()
        if not project:
            continue
        wanted_kind = str(hit.get("업무종류") or "").strip()
        candidates = list(by_project.get(project) or [])
        if wanted_kind:
            candidates = [item for item in candidates if item[0] == wanted_kind]
        if not candidates:
            result["missing"] += 1
            result["records"].append({"project": project, "action": "missing"})
            continue
        if len(candidates) != 1:
            result["ambiguous"] += 1
            result["records"].append({"project": project, "action": "ambiguous",
                                      "count": len(candidates)})
            continue

        kind, status_col, current = candidates[0]
        full_text = str(hit.get("원문") or hit.get("근거") or "")
        treatment = str(hit.get("처리구분") or outcome_kind(full_text) or "접수취소")
        reason = "유선전화 원격해결" if treatment == "원격해결" else "접수취소"
        # 접수취소가 실제로 발생한 날은 밴드 게시일/댓글일이다. 수집기가 글을
        # 뒤늦게 다시 읽은 ``captured_at``을 먼저 쓰면 1월 취소가 8월 취소로
        # 기록된다. 원문 사건일을 우선하고, 과거 캐시에 사건일이 없을 때만
        # 관측시각을 최후 폴백으로 쓴다.
        confirmed_on = _as_day(
            hit.get("게시일") or hit.get("event_day")
            or hit.get("관측시각") or hit.get("captured_at")
        )
        corroboration = corroborations.get(project) or {}
        fields = dict(current.get("fields") or {})
        evidence = _hit_evidence(hit, corroboration)
        # 교차 리포트는 재생성 중 잠깐 프로젝트가 빠질 수 있다. 그 결측만으로 이미
        # 검증해 둔 카톡 근거를 지우면 다음 회차에 다시 나타나 같은 행을 되풀이해 쓴다.
        # 새 교차근거가 없을 때만 기존 교차 부분을 보존하고, 밴드 부분은 현재의
        # 정규화된 숫자 밴드 ID로 다시 만든다.
        current_evidence = str(fields.get("접수취소근거") or "").strip()
        cross_marker = " · 카톡 교차 "
        if not corroboration and cross_marker in current_evidence:
            old_cross = current_evidence.split(cross_marker, 1)[1].strip()
            evidence = (_hit_evidence(hit, {}) + cross_marker + old_cross)[:900]
        source_ref = str(hit.get("근거URL") or "").strip()
        desired_fields = {
            status_col: "취소",
            "접수취소여부": "예",
            "접수취소사유": reason,
            "처리구분": treatment,
            "접수취소근거": evidence,
        }
        if confirmed_on:
            desired_fields["접수취소확인일"] = confirmed_on
        current_status = str(current.get("status") or fields.get(status_col) or "").strip()
        done_value = (fields.get("작업완료일") if kind == "돌발AS"
                      else fields.get("실제점검일"))
        already_cancelled = any(word in current_status for word in _CANCELLED)
        # 이후 완료·현장실적을 접수취소 본문 하나로 뒤집으면 실제 작업과 청구를
        # 잃는다. 이미 취소인 행의 근거 보강은 허용하고, 열린 상태만 취소 전환한다.
        열린낱말, 못읽음 = open_states()
        if (not already_cancelled
                and (str(done_value or "").strip() or current_status not in 열린낱말)):
            result["conflicts"] += 1
            result["records"].append({
                "project": project, "work_id": current["id"], "action": "conflict",
                "reason": closed_reason(current_status, done_value, 열린낱말, 못읽음),
                "current_status": current_status,
            })
            continue
        same = str(current.get("status") or "").strip() == "취소" and all(
            fields.get(key) == value for key, value in desired_fields.items()
        )
        if same:
            result["unchanged"] += 1
            result["records"].append({"project": project, "work_id": current["id"],
                                      "action": "unchanged", "treatment": treatment})
            continue

        token_payload = {
            "algorithm": "cancel-resolution/v2",
            "project": project,
            "kind": kind,
            "band": hit.get("밴드"),
            "post": hit.get("게시글"),
            "confirmed_on": confirmed_on,
            "treatment": treatment,
            "evidence": evidence,
            # 같은 사실 A→B→A 로 입력 근거가 되돌아와도, 과거 버전에서 쓴 멱등키를
            # 현재 버전 요청에 재사용하면 AppStore가 올바르게 IdempotencyConflict를 낸다.
            # 낙관잠금 버전을 키에 포함하면 같은 버전의 정확한 재시도만 재생되고,
            # 이후 버전의 새 요청은 별개의 안전한 쓰기가 된다.
            "record_version": int(current["record_version"]),
        }
        token = hashlib.sha256(
            json.dumps(token_payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
        ).hexdigest()[:32]
        try:
            response = store.update_work(
                current["id"],
                expected_version=int(current["record_version"]),
                patch={"status": "취소", "fields": desired_fields},
                actor=actor,
                source="band+kakao:cancel-resolution",
                evidence=evidence,
                source_ref=source_ref,
                source_observed_at=(confirmed_on or None),
                idempotency_key=f"cancel-resolution:{token}",
            )
            result["updated"] += 1
            result["records"].append({"project": project, "work_id": current["id"],
                                      "action": "updated", "treatment": treatment,
                                      "event_id": response.get("event_id")})
        except Exception as exc:
            result["ok"] = False
            result["errors"].append({"project": project, "type": type(exc).__name__,
                                     "message": str(exc)[:300]})
    return result
