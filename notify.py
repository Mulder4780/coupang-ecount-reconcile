# -*- coding: utf-8 -*-
"""올린 것을 알린다 — 알림을 **만드는 자리는 여기 하나다**.

2026-08-14 지시: "류지영이 카톡 텍스트 파일 앱에 업로드 하면 바로 정리 시작하고
나에게 올렸다고 알려주는 구조 코딩해 / 업로드 알림 만들어"

실측(2026-08-14): '올리면 바로 정리 시작'은 **이미 됐다**(`/api/automation/kakao-upload`
가 그 자리에서 `start_task("automation")`). 빠져 있던 것은 **알리는 쪽**이고,
코드 전체에 사람에게 알리는 길이 **0건**이었다.

지키는 것 (SPEC_업로드알림.md ④)
  · **판단을 두 곳에 두지 않는다**([162]) — 알림을 만드는 함수는 `push` 하나다.
  · **못 보냈으면 못 보냈다고 적는다**([169]) — 채널 실패를 성공으로 안 적는다.
    `reports/알림_기록.json` 에 보냄·실패가 채널별로 남는다.
  · **알림이 대부분이면 아무도 안 본다**([170]) — 같은 갈래·같은 받는이가
    `MERGE_WINDOW_SEC` 안에 여러 번이면 **한 줄로 합친다**.
  · **업무값을 밖으로 안 내보낸다** — 외부 채널에는 갈래·상태·건수만.
    프로젝트NO·금액·캠프명 모양이 섞이면 그 문장을 통째로 버리고 맨 줄로 대신한다.
  · **알림 실패가 업로드를 죽이지 않는다** — `push` 는 어떤 예외도 밖으로 안 낸다.
    알리려다 원본 접수를 막으면 본말전도다.

★ '올렸다'만 알리면 반쪽이다. 처리가 죽어도 형님은 된 줄 알고 넘어간다([169]).
  그래서 `expect_upload()` 로 접수를 적어 두고 `sweep_uploads()` 가 회차가 남긴
  자국을 **읽어서**([168] — 다시 세지 않는다) 끝남·실패를 뒤따라 알린다.
  회차가 끝났는지 **못 물어본** 경우도 조용히 넘기지 않고 그렇게 말한다.
"""

from __future__ import annotations

import json
import os
import sys
import time
import uuid
from datetime import datetime, timedelta, timezone

ROOT = os.path.dirname(os.path.abspath(__file__))
REPORTS = os.path.join(ROOT, "reports")

STORE = os.path.join(REPORTS, "알림_대기.json")        # 사람이 볼 알림(앱이 읽는다)
LOG = os.path.join(REPORTS, "알림_기록.json")          # 보냄·실패 자국
PENDING = os.path.join(REPORTS, "알림_대기업로드.json")  # 접수했고 결과를 아직 못 본 것
CONF = os.path.join(ROOT, "config", "notify.json")     # 외부 채널(없으면 앱 안 알림만)
PIPELINE_STATE = os.path.join(REPORTS, "automation_pipeline_state.json")

MERGE_WINDOW_SEC = 300      # 같은 갈래가 5분 안에 여러 번이면 한 줄로 합친다
KEEP_ITEMS = 200            # 저장소가 무한히 자라지 않게
KEEP_LOG = 400
DEFAULT_TTL_H = 72          # 알림이 화면에 남는 기본 기간
FAIL_TTL_H = 24 * 14        # 실패는 오래 남는다 — 사람이 볼 때까지
UNRESOLVED_WARN_H = 6       # 이만큼 지나도 회차 자국이 없으면 '확인 못 함'이라 말한다

# 받는이 토큰. 사람 이름을 코드에 적지 않는다 — 역할·slug 로만 적는다.
ADMIN = "admin"


def staff(slug):
    """업무센터 담당자 토큰. slug 가 비면 관리자만 받는다."""
    slug = str(slug or "").strip()
    return f"staff:{slug}" if slug else ADMIN


def actor_token(role, slug=""):
    """app_server `_actor()` 의 role/slug 를 받는이 토큰으로 옮긴다."""
    role = str(role or "").strip()
    if role == "admin":
        return ADMIN
    if role == "staff":
        return staff(slug)
    return ""


# ───────────────────────── 파일 입출력 ─────────────────────────

def _now_local():
    return datetime.now().astimezone()


def _iso(dt):
    return dt.isoformat(timespec="seconds")


def _parse(value):
    """ISO 문자열 → aware datetime. 못 읽으면 None(모른다고 말한다)."""
    text = str(value or "").strip()
    if not text:
        return None
    try:
        dt = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except (TypeError, ValueError):
        return None
    if dt.tzinfo is None:                       # 순진한 시각은 이 PC 시간대로 읽는다
        dt = dt.replace(tzinfo=_now_local().tzinfo)
    return dt


def _load(path, default):
    try:
        with open(path, encoding="utf-8") as fh:
            value = json.load(fh)
    except (OSError, ValueError):
        return default
    if isinstance(default, list) and not isinstance(value, list):
        return default
    if isinstance(default, dict) and not isinstance(value, dict):
        return default
    return value


def _save(path, payload):
    """원자적 저장. 읽는 쪽이 물고 있으면 물러서며 다시 건다([171] 와 같은 자리)."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = f"{path}.{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=1)
        fh.flush()
        os.fsync(fh.fileno())
    delay = 0.05
    for attempt in range(5):
        try:
            os.replace(tmp, path)
            return
        except OSError:
            if attempt == 4:
                try:
                    os.remove(tmp)
                except OSError:
                    pass
                raise
            time.sleep(delay)
            delay *= 2


# ───────────────────────── 채널 ─────────────────────────

def channels():
    """쓸 수 있는 채널 목록. **앱 안 알림은 항상 켜져 있다.**

    외부 채널은 `config/notify.json` 에 사람이 적어야 켜진다 — 지금은 확실히 도는
    길이 앱 안 알림뿐이라 그것부터 채운다(SPEC ①).
    """
    out = [{"name": "app", "kind": "app", "enabled": True}]
    conf = _load(CONF, {})
    for row in conf.get("channels") or []:
        if not isinstance(row, dict):
            continue
        name = str(row.get("name") or "").strip()
        if not name or name == "app":
            continue
        out.append({
            "name": name,
            "kind": str(row.get("kind") or "webhook"),
            "enabled": bool(row.get("enabled")),
            "url": str(row.get("url") or ""),
            "timeout": float(row.get("timeout") or 5),
        })
    return out


# 업무값 모양 — 이런 것이 섞이면 외부로 안 내보낸다.
#   프로젝트NO(UJ2601321) · PO번호 · 금액(1,234,500 / 3049310원) · 캠프명
_BIZ_PATTERNS = (
    r"\b[A-Z]{2}\d{6,8}\b",          # 프로젝트NO
    r"\bPO\s?\d{5,8}\b",             # PO번호
    r"\b\d{1,3}(?:,\d{3})+\b",       # 1,234,500
    r"\d{4,}\s*원",                   # 3049310원
    r"[가-힣]{2,}\d*(?:캠프|MB|FC|Sub-FC)",   # 캠프명
)


def leaks_business_value(text):
    """외부로 나가면 안 되는 업무값 모양이 있나. 있으면 그 문장은 안 보낸다."""
    import re
    body = str(text or "")
    for pat in _BIZ_PATTERNS:
        if re.search(pat, body):
            return True
    return False


def external_text(rec):
    """외부 채널용 한 줄 — **건수·상태만**. 제목·본문을 그대로 싣지 않는다.

    제목에는 파일명·사람 이름이 섞이고, 본문에는 업무값이 섞인다. 제3자 중계를
    거치는 길에 그것을 올리지 않는다(SPEC ④).
    """
    parts = [f"[CSOS] {rec.get('갈래') or '알림'}"]
    상태 = str(rec.get("상태") or "").strip()
    if 상태:
        parts.append(상태)
    n = int(rec.get("건수") or 1)
    if n > 1:
        parts.append(f"{n}건")
    line = " · ".join(parts)
    # 갈래 이름 자체에 업무값을 넣는 날이 와도 새어 나가지 않게 마지막에 한 번 더 본다.
    return line if not leaks_business_value(line) else "[CSOS] 새 알림"


def _send_external(ch, rec):
    """한 채널로 보낸다. 성공하면 True. **확인 못 하면 성공이 아니다.**"""
    if not ch.get("enabled"):
        return False, "꺼짐"
    url = str(ch.get("url") or "")
    if ch.get("kind") != "webhook" or not url.startswith("https://"):
        return False, "보낼 길이 없다(https webhook 만 지원)"
    import urllib.request
    body = json.dumps({"text": external_text(rec)}).encode("utf-8")
    req = urllib.request.Request(
        url, data=body, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=ch.get("timeout") or 5) as resp:
        code = int(getattr(resp, "status", 0) or resp.getcode() or 0)
    if 200 <= code < 300:
        return True, f"HTTP {code}"
    return False, f"HTTP {code}"


# ───────────────────────── 알림 만들기 ─────────────────────────

def push(kind, title, body="", evidence=None, audience=None,
         severity="info", 상태="", ttl_hours=None):
    """알림 하나. **어떤 예외도 밖으로 내지 않는다** — 알리려다 업무를 막지 않는다.

    돌려주는 것: `{"보냄": [채널…], "실패": [{"채널","왜"}…], "id": …}`
    """
    result = {"보냄": [], "실패": [], "id": ""}
    try:
        audience = [a for a in (audience or [ADMIN]) if a]
        if not audience:
            audience = [ADMIN]
        now = _now_local()
        ttl = FAIL_TTL_H if severity == "error" else (ttl_hours or DEFAULT_TTL_H)
        rec = {
            "id": uuid.uuid4().hex[:12],
            "갈래": str(kind or "알림"),
            "제목": str(title or "")[:200],
            "본문": str(body or "")[:600],
            "상태": str(상태 or "")[:40],
            "심각도": severity if severity in ("info", "warning", "error") else "info",
            "받는이": sorted(set(str(a) for a in audience)),
            "근거": str(evidence or "")[:200],
            "때": _iso(now),
            "건수": 1,
            "만료": _iso(now + timedelta(hours=ttl)),
            "읽음": False,
        }

        items = _load(STORE, [])
        merged = None
        for old in items:
            if not isinstance(old, dict) or old.get("읽음"):
                continue
            if old.get("갈래") != rec["갈래"] or old.get("받는이") != rec["받는이"]:
                continue
            when = _parse(old.get("때"))
            if when and (now - when).total_seconds() <= MERGE_WINDOW_SEC:
                merged = old
                break
        if merged is not None:
            # 합치되 숨기지 않는다([169]) — 몇 건이 합쳐졌는지 화면이 말한다.
            merged["건수"] = int(merged.get("건수") or 1) + 1
            merged["제목"] = rec["제목"]
            merged["본문"] = rec["본문"]
            merged["상태"] = rec["상태"]
            merged["때"] = rec["때"]
            merged["만료"] = rec["만료"]
            if _SEV_RANK.get(rec["심각도"], 0) > _SEV_RANK.get(merged.get("심각도"), 0):
                merged["심각도"] = rec["심각도"]
            rec = merged
        else:
            items.append(rec)
        _save(STORE, items[-KEEP_ITEMS:])
        result["id"] = rec["id"]
        result["보냄"].append("app")
    except Exception as exc:                     # noqa: BLE001 — 절대 밖으로 안 낸다
        result["실패"].append({"채널": "app", "왜": f"{type(exc).__name__}: {exc}"[:180]})
        rec = {"갈래": str(kind or "알림"), "상태": str(상태 or ""), "건수": 1}

    for ch in channels():
        if ch["name"] == "app":
            continue
        try:
            ok, why = _send_external(ch, rec)
        except Exception as exc:                 # noqa: BLE001
            ok, why = False, f"{type(exc).__name__}: {exc}"[:180]
        if ok:
            result["보냄"].append(ch["name"])
        else:
            result["실패"].append({"채널": ch["name"], "왜": why})

    _log(rec, result)
    return result


_SEV_RANK = {"info": 0, "warning": 1, "error": 2}


def _log(rec, result):
    """보냄·실패를 남긴다. **못 보냈으면 못 보냈다고 적는다**([169])."""
    try:
        rows = _load(LOG, [])
        rows.append({
            "때": _iso(_now_local()),
            "갈래": rec.get("갈래"),
            "제목": str(rec.get("제목") or "")[:120],
            "받는이": rec.get("받는이"),
            "보냄": result.get("보냄") or [],
            "실패": result.get("실패") or [],
        })
        _save(LOG, rows[-KEEP_LOG:])
    except Exception:                            # noqa: BLE001
        pass


# ───────────────────────── 읽기 ─────────────────────────

def feed(role="admin", slug="", include_read=False):
    """이 사람이 볼 알림. 만료된 것은 빼고 최신순."""
    token = actor_token(role, slug)
    now = _now_local()
    out = []
    for row in _load(STORE, []):
        if not isinstance(row, dict):
            continue
        if not include_read and row.get("읽음"):
            continue
        until = _parse(row.get("만료"))
        if until and until < now:
            continue
        who = row.get("받는이") or [ADMIN]
        if token and token not in who:
            continue
        if not token:
            continue
        out.append(row)
    out.sort(key=lambda r: str(r.get("때") or ""), reverse=True)
    return out


def ack(ids, role="admin", slug=""):
    """사람이 확인했다고 표시한다. **내가 받는 알림만** 내린다."""
    token = actor_token(role, slug)
    want = set(str(i) for i in (ids or []))
    if not token:
        return 0
    items = _load(STORE, [])
    n = 0
    for row in items:
        if not isinstance(row, dict) or row.get("읽음"):
            continue
        if want and str(row.get("id")) not in want:
            continue
        if token not in (row.get("받는이") or [ADMIN]):
            continue
        row["읽음"] = True
        row["읽은때"] = _iso(_now_local())
        n += 1
    if n:
        _save(STORE, items)
    return n


# ───────────────────────── 접수 → 끝남·실패 ─────────────────────────

def expect_upload(kind, label, audience, source="automation", note=""):
    """접수를 적어 둔다 — 결과를 뒤따라 알리기 위해서다.

    `source="automation"` 이면 `automation_pipeline` 이 남기는 자국을 읽어 판정한다.
    """
    try:
        rows = _load(PENDING, [])
        rows.append({
            "id": uuid.uuid4().hex[:12],
            "갈래": str(kind or "업로드"),
            "이름": str(label or "")[:160],
            "받는이": sorted(set(str(a) for a in (audience or [ADMIN]) if a)) or [ADMIN],
            "원천": str(source or "automation"),
            "메모": str(note or "")[:160],
            "올린때": _iso(_now_local()),
        })
        _save(PENDING, rows[-100:])
        return True
    except Exception:                            # noqa: BLE001
        return False


def _runs():
    """회차가 남긴 자국을 **읽기만** 한다([168]). 못 읽으면 None — 0 이 아니다."""
    state = _load(PIPELINE_STATE, None)
    if not isinstance(state, dict):
        return None
    runs = []
    last = state.get("last_run")
    if isinstance(last, dict):
        runs.append(last)
    for row in state.get("history") or []:
        if isinstance(row, dict):
            runs.append(row)
    return runs


def _run_verdict(run):
    """회차 하나의 갈래·설명. 칸 이름은 실측한 그대로다([165]).

    `status` 는 성공일 때 **`success`** 다(`ok` 가 아니다).
    """
    status = str(run.get("status") or "")
    summary = str(run.get("summary") or "")
    if status == "success":
        changed = run.get("changed_sources") or []
        detail = f"변경 원천 {', '.join(changed)}" if changed else "새로 바뀐 원천 없음"
        return "끝", f"{detail} · {summary}".strip(" ·")
    if status in ("partial", "failed", "lost_lock"):
        stage = str(run.get("current_stage") or "")
        bad = run.get("failures") or []
        for st in run.get("stages") or []:
            if isinstance(st, dict) and not st.get("ok"):
                stage = str(st.get("name") or stage)
                summary = str(st.get("summary") or summary)
                break
        why = ", ".join(str(b) for b in bad) if bad else summary
        return "실패", f"{stage} · {why}".strip(" ·")
    return "", summary


def sweep_uploads():
    """접수해 둔 것의 결과를 알린다. 아무것도 없으면 파일을 안 건드린다.

    ★ **누가 돌렸든 결과를 읽는다** — 앱이 시작한 회차든 5분 스케줄러가 돌린 회차든
      끝나면 같은 자국(`automation_pipeline_state.json`)에 남는다.
    """
    result = {"끝": 0, "실패": 0, "확인못함": 0, "남음": 0}
    try:
        rows = _load(PENDING, [])
        if not rows:
            return result
        runs = _runs()
        now = _now_local()
        keep = []
        for row in rows:
            if not isinstance(row, dict):
                continue
            born = _parse(row.get("올린때"))
            who = row.get("받는이") or [ADMIN]
            name = row.get("이름") or ""
            kind = row.get("갈래") or "업로드"
            done = None
            if runs is not None and born is not None:
                for run in runs:
                    fin = _parse(run.get("finished_at"))
                    if fin and fin >= born:
                        done = run
                        break
            if done is not None:
                verdict, detail = _run_verdict(done)
                if verdict == "끝":
                    push(f"{kind} 정리 끝",
                         f"{name} 정리가 끝났습니다",
                         detail, evidence="reports/automation_pipeline_state.json",
                         audience=who, severity="info", 상태="끝")
                    result["끝"] += 1
                    continue
                if verdict == "실패":
                    push(f"{kind} 정리 실패",
                         f"{name} 정리가 실패했습니다 — 확인이 필요합니다",
                         detail, evidence="reports/automation_pipeline_state.json",
                         audience=who, severity="error", 상태="실패")
                    result["실패"] += 1
                    continue
                # status 가 running·already_running 이면 아직 결과가 아니다 — 기다린다.
            if born is not None and (now - born) > timedelta(hours=UNRESOLVED_WARN_H):
                # ★ 못 물어본 것을 '이상 없음'이라 하지 않는다([169]).
                why = ("회차 자국을 못 읽었습니다" if runs is None
                       else "회차가 끝난 자국이 아직 없습니다")
                push(f"{kind} 결과 확인 못 함",
                     f"{name} 정리 결과를 확인하지 못했습니다",
                     why, evidence="reports/automation_pipeline_state.json",
                     audience=who, severity="warning", 상태="확인못함")
                result["확인못함"] += 1
                continue
            keep.append(row)
        result["남음"] = len(keep)
        if len(keep) != len(rows):
            _save(PENDING, keep)
    except Exception:                            # noqa: BLE001
        pass
    return result


# ───────────────────────── 사람이 볼 때 ─────────────────────────

def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if "--sweep" in argv:
        print("결과 확인:", json.dumps(sweep_uploads(), ensure_ascii=False))
        return 0
    if "--test" in argv:
        print("보냄/실패:", json.dumps(
            push("시험", "알림 시험", "이 줄은 시험입니다", audience=[ADMIN]),
            ensure_ascii=False))
        return 0
    rows = feed("admin")
    print(f"관리자 알림 {len(rows)}건 · 채널 "
          f"{', '.join(c['name'] for c in channels() if c.get('enabled'))}")
    for row in rows[:15]:
        print(f"  [{row.get('심각도')}] {row.get('때')} · {row.get('제목')}"
              + (f" (x{row['건수']})" if int(row.get("건수") or 1) > 1 else ""))
    return 0


if __name__ == "__main__":
    try:                                          # 무인 회차는 stdout 이 None 이다([235])
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:                             # noqa: BLE001
        pass
    sys.exit(main())
