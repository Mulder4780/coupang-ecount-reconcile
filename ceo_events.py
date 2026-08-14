# -*- coding: utf-8 -*-
"""유수비 대표 카카오톡 → 대표보고용 쿠팡 접수 전달 추출.

대표는 쿠팡 대화방에 프로젝트번호 배정 전 접수와 쿠팡 업무지시를 함께 올린다.
기존 ``kakao_extract`` 는 프로젝트번호가 없는 글을 버리므로 접수가 대표보고·캡처에서
조용히 빠졌다. 이 모듈은 원문을 고치거나 원장에 쓰지 않고, 대표가 올린 메시지를
``쿠팡 건 / 쿠팡 무관 / 모름``으로 가른 뒤 쿠팡 접수와 대표 지시를 읽기 전용 캐시에 남긴다.

실행::

    python ceo_events.py --sync
"""
import argparse
import hashlib
import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime

# 콘솔 cp949 는 '—' 를 못 쓰고, 무인 회차(pythonw)는 sys.stdout 이 None 이다.
# 맨몸 reconfigure 는 그 자리에서 AttributeError 로 회차를 통째로 죽인다(`[235]`).
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import kakao_extract as KE


ROOT = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.environ.get("COUPANG_REPORT_DIR") or os.path.join(ROOT, "reports")
REPORT_JSON = os.path.join(REPORT_DIR, "대표대화_추출.json")
REPORT_MD = os.path.join(REPORT_DIR, "대표대화_추출.md")
# 사람이 "처리했다"고 확인한 지시. 리포트가 아니라 이 파일이 정본이다 — 리포트는 09:50
# 회차가 카톡 원문에서 매번 새로 만들므로, 리포트에서 지우면 다음 회차에 그대로 되살아난다.
ACK_JSON = os.path.join(REPORT_DIR, "대표지시_확인.json")
# daily_brief 가 화면·캡처에서 내리는 낱말이다. 여기서 새 낱말을 지어내면 확인해도 안 내려간다.
ACK_STATE = "완료"

# 사람이 바뀌거나 카카오 표시명이 흔들리면 이 표만 늘린다. 판정 코드에 이름을 흩뿌리지 않는다.
CEO_SENDERS = {
    "유수비 대표 유니버셜리프트": "유수비 대표",
    "유수비 대표": "유수비 대표",
    "유수비": "유수비 대표",
}
COUPANG_ROOM_MARKERS = ("쿠팡",)
ATTACHMENT_ONLY = re.compile(
    r"^(?:사진|동영상|파일|음성메시지)(?:\s*(?:[x×]\s*)?\d+\s*(?:장|개)?)?$", re.I
)
COUNT_KEYS = ("쿠팡 건", "쿠팡 무관", "모름")


def _space(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def sender_role(sender):
    """대표 표시명 정본. 부분 일치는 동명이인을 삼킬 수 있어 쓰지 않는다."""
    return CEO_SENDERS.get(_space(sender), "")


def fields_of(text):
    """공지 필드는 kakao_extract의 라벨·정리 규칙을 그대로 빌린다."""
    fields = {}
    for label, value in KE.RE_FIELD.findall(str(text or "")):
        key = KE.ALIAS.get(label.strip().lower())
        if not key:
            continue
        value = KE.clean(value)
        # 실제 양식은 'GWJ1' 다음에 'GWJ1 M_순천1'을 한 번 더 쓴다. 캠프는 더 구체적인
        # 값을 택하고, 나머지는 원래 추출기와 같이 먼저 나온 값을 근거로 삼는다.
        if key == "캠프명":
            if len(value) > len(fields.get(key, "")):
                fields[key] = value
        elif key not in fields:
            fields[key] = value
    return fields


def _event(room, msg, fields):
    text = str(msg.get("text") or "")
    code_match = KE.RE_CODE.search(text)
    code = code_match.group(0) if code_match else ""
    day = msg["date"].isoformat()
    stamp = "%s %s" % (day, msg.get("time") or "")
    fingerprint = hashlib.sha256(
        (room + "\0" + stamp + "\0" + _space(msg.get("sender")) + "\0" + text)
        .encode("utf-8", errors="replace")
    ).hexdigest()[:16]
    kind = "정기점검접수" if ("정기" in room or "분기" in room) else "돌발AS접수"
    requested = KE.norm_date(fields.get("신청일자"))
    return {
        "날짜": day,                         # 파일명이 아니라 카톡 날짜 구분선에서 온다
        "업무유형": kind,
        "프로젝트NO": code,
        "캠프명": fields.get("캠프명", ""),
        "접수일": requested.isoformat() if requested else "",
        "게시자": sender_role(msg.get("sender")),
        "처리자": "",
        "신청내용": fields.get("신청내용", ""),
        "처리내용": "",
        "상태": "대표 전달 접수",
        "비용": "",
        "근거": f"{stamp} 카카오톡 · {room}",
        "레코드ID": "CEO-" + fingerprint,
        "레코드종류": "ceo",
        "출처파일": str(msg.get("_source") or ""),
        "방": room,
        "카톡일시": stamp,
        "원문": text,
    }


DIRECTIVE_ITEMS = (
    {
        "제목": "AS 부품 사전신청 원칙 교육",
        "요구내용": "필요 부품을 사전에 신청하고 로켓배송 익일 도착을 활용해 AS 일정 차질을 방지. 당일 구매는 긴급 상황만 예외",
        "제출내용": "교육 실시 결과",
    },
    {
        "제목": "긴급 구매 사전승인 절차",
        "요구내용": "부득이한 당일 긴급 구매는 구매부 사전 승인 후 진행하고 승인 없는 임의 구매를 금지",
        "제출내용": "긴급 구매 절차 개선 내용",
    },
    {
        "제목": "상시 재고 운영체계 구축",
        "요구내용": "상시 사용 AS 품목 목록과 적정 재고·사전 구매 기준을 마련하고 쿠팡 관계자 및 김정훈 매니저와 협의",
        "제출내용": "AS 재고 유지 품목 리스트 및 운영방안",
    },
)


def _directive_due(text):
    """지시 원문의 보고기한만 읽는다. 없으면 날짜를 만들어 내지 않는다."""
    m = re.search(
        r"(20\d{2})년\s*(\d{1,2})월\s*(\d{1,2})일(?:\([^)]*\))?\s*"
        r"(?:(오전|오후)\s*)?(\d{1,2})(?::(\d{2}))?\s*까지",
        str(text or ""),
    )
    if not m:
        return ""
    year, month, day, ap, hour, minute = m.groups()
    hour = int(hour)
    if ap == "오후" and hour < 12:
        hour += 12
    elif ap == "오전" and hour == 12:
        hour = 0
    return f"{int(year):04d}-{int(month):02d}-{int(day):02d} {hour:02d}:{int(minute or 0):02d}"


def _directive(room, msg):
    """쿠팡 대표 지시를 AS 접수와 섞지 않고 별도 이행현황으로 만든다."""
    text = str(msg.get("text") or "")
    day = msg["date"].isoformat()
    stamp = "%s %s" % (day, msg.get("time") or "")
    fingerprint = hashlib.sha256(
        (room + "\0" + stamp + "\0" + _space(msg.get("sender")) + "\0" + text)
        .encode("utf-8", errors="replace")
    ).hexdigest()[:16]
    items = [dict(row, 상태="결과 미확인") for row in DIRECTIVE_ITEMS]
    return {
        "날짜": day,
        "제목": "쿠팡 AS 업무 운영 정비",
        "게시자": sender_role(msg.get("sender")),
        "보고기한": _directive_due(text),
        "상태": "경과보고 확인 필요",
        "항목": items,
        "근거": f"{stamp} 카카오톡 · {room}",
        "레코드ID": "CEO-DIR-" + fingerprint,
        "레코드종류": "ceo_directive",
        "출처파일": str(msg.get("_source") or ""),
        "방": room,
        "카톡일시": stamp,
        "원문": text,
    }


def classify(room, msg):
    """대표 메시지 하나를 세 갈래로 판정한다. 대표가 아니면 ``None``이다."""
    if not sender_role(msg.get("sender")):
        return None, None, "대표 아님"
    text = _space(msg.get("text"))
    if not any(marker in room for marker in COUPANG_ROOM_MARKERS):
        return "쿠팡 무관", None, "비쿠팡 대화방"
    if ATTACHMENT_ONLY.fullmatch(text):
        return "쿠팡 무관", None, "첨부만 있는 메시지"
    if "대표이사 업무지시서" in text:
        if "쿠팡" in text and "AS" in text.upper():
            return "쿠팡 건", _directive(room, msg), "쿠팡 대표 지시 확인"
        return "쿠팡 무관", None, "비쿠팡 회사 내부 지시"

    fields = fields_of(text)
    has_case_signal = bool(fields.get("캠프명") or KE.RE_CODE.search(text))
    if not has_case_signal:
        return "모름", None, "캠프명·프로젝트번호 근거 없음"
    return "쿠팡 건", _event(room, msg, fields), "쿠팡 접수 근거 확인"


def load_ack(path=None):
    """사람이 확인한 지시 표. 못 읽으면 빈 표다 — 못 읽은 것을 '확인함'으로 치지 않는다."""
    try:
        raw = json.load(open(path or ACK_JSON, encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return {}
    rows = raw.get("확인") if isinstance(raw, dict) else None
    if not isinstance(rows, dict):
        return {}
    return {str(k): v for k, v in rows.items() if isinstance(v, dict)}


def apply_ack(directives, acks=None):
    """확인된 지시를 완료로 내린다. **지우지 않는다** — 누가 언제 확인했는지를 붙여 남긴다.

    지우면 "그때 정말 그런 지시가 있었나"를 잃고, 확인 근거도 같이 사라진다.
    화면·캡처에서 내리는 일은 ``상태`` 낱말 하나로 충분하다(daily_brief 가 이미 거른다).
    """
    acks = load_ack() if acks is None else acks
    out = []
    for row in directives:
        got = acks.get(str(row.get("레코드ID") or "")) if isinstance(row, dict) else None
        if not isinstance(got, dict):
            out.append(row)
            continue
        row = dict(row)
        row["상태"] = ACK_STATE
        row["확인자"] = str(got.get("who") or "관리자")
        row["확인시각"] = str(got.get("when") or "")
        row["확인근거"] = str(got.get("why") or "사람 확인")
        row["항목"] = [dict(item, 상태=ACK_STATE) for item in row.get("항목") or []]
        out.append(row)
    return out


def save_ack(record_id, who="관리자", why="", path=None, remove=False):
    """지시 하나를 확인 처리하거나 되돌린다. **없는 번호는 조용히 넘기지 않는다.**

    0건을 성공으로 적으면 사람은 내려간 줄 알고 화면을 다시 안 본다(`[169]`).
    """
    path = path or ACK_JSON
    record_id = str(record_id or "").strip()
    if not record_id:
        raise ValueError("레코드ID 가 비어 있습니다")
    known = {str(r.get("레코드ID")) for r in (load_cached().get("directives") or [])
             if isinstance(r, dict)}
    if known and record_id not in known and not remove:
        raise ValueError("리포트에 없는 레코드ID 입니다: %s\n  → python ceo_events.py --list"
                         % record_id)
    try:
        raw = json.load(open(path, encoding="utf-8"))
        rows = raw.get("확인") if isinstance(raw, dict) else None
        rows = dict(rows) if isinstance(rows, dict) else {}
    except (OSError, ValueError, TypeError):
        rows = {}
    if remove:
        if record_id not in rows:
            raise ValueError("확인 기록에 없는 레코드ID 입니다: " + record_id)
        rows.pop(record_id)
    else:
        rows[record_id] = {"who": str(who or "관리자"),
                           "when": datetime.now().isoformat(timespec="seconds"),
                           "why": str(why or "사람 확인")}
    _atomic_write(path, json.dumps({"version": 1, "확인": rows},
                                   ensure_ascii=False, indent=2))
    return rows


def extract(paths=None):
    """카카오톡 원본을 읽어 판정과 접수 이벤트를 만든다. 같은 내보내기 이력은 한 번만 센다."""
    paths = KE.source_paths() if paths is None else list(paths)
    parser = KE._load_reconcile()
    counts = Counter({key: 0 for key in COUNT_KEYS})
    daily = defaultdict(lambda: Counter({key: 0 for key in COUNT_KEYS}))
    reasons, events, directives, seen, directive_seen = Counter(), [], [], set(), set()

    for path in paths:
        room = KE.room_of(path)
        for msg in parser.parse_export(path):
            if not sender_role(msg.get("sender")):
                continue
            text = str(msg.get("text") or "")
            # 전체 대화 내보내기 사본마다 같은 과거 메시지가 반복된다. 파일이 아니라
            # 메시지 지문으로 중복을 막아야 내보내기 횟수가 접수 건수로 부풀지 않는다.
            key = (room, msg["date"].isoformat(), msg.get("time") or "",
                   _space(msg.get("sender")), _space(text))
            if key in seen:
                continue
            seen.add(key)
            msg = dict(msg, _source=os.path.basename(path))
            verdict, event, why = classify(room, msg)
            if verdict is None:
                continue
            # 같은 대표 지시가 돌발·정기점검 방 양쪽에 동시에 공지된다. 방이 다르다는 이유로
            # 두 지시로 세면 이행항목도 캡처에 두 벌 나온다. 접수는 방을 보존하되 지시는
            # 카톡 일시·게시자·본문이 같으면 한 번만 센다.
            if event and event.get("레코드종류") == "ceo_directive":
                directive_key = (msg["date"].isoformat(), msg.get("time") or "",
                                 _space(msg.get("sender")), _space(text))
                if directive_key in directive_seen:
                    continue
                directive_seen.add(directive_key)
            counts[verdict] += 1
            daily[msg["date"].isoformat()][verdict] += 1
            reasons[why] += 1
            if event and event.get("레코드종류") == "ceo_directive":
                directives.append(event)
            elif event:
                events.append(event)

    events.sort(key=lambda row: (row.get("날짜", ""), row.get("카톡일시", ""),
                                 row.get("레코드ID", "")))
    directives = apply_ack(directives)
    directives.sort(key=lambda row: (row.get("보고기한", "9999"), row.get("카톡일시", ""),
                                     row.get("레코드ID", "")))
    return {
        "version": 2,
        "상태": "확인",
        "생성시각": datetime.now().isoformat(timespec="seconds"),
        "원본파일수": len(paths),
        "대표메시지수": sum(counts.values()),
        "판정": {key: counts[key] for key in COUNT_KEYS},
        "날짜별": {day: {key: values[key] for key in COUNT_KEYS}
                 for day, values in sorted(daily.items())},
        "제외이유": dict(sorted(reasons.items())),
        "events": events,
        "directives": directives,
    }


def render(report):
    c = report.get("판정") or {}
    lines = [
        "# 유수비 대표 대화 쿠팡 접수 판정",
        "",
        f"- 생성: {report.get('생성시각') or '-'}",
        f"- 원본: {report.get('원본파일수') or 0}개",
        f"- 쿠팡 건 **{c.get('쿠팡 건', 0)}** · 쿠팡 무관 **{c.get('쿠팡 무관', 0)}** · 모름 **{c.get('모름', 0)}**",
        "- 읽기 전용: 원장·입력 큐·카카오톡 원문은 수정하지 않음",
        "",
        "## 반영 대상",
        "",
        "| 카톡 일시 | 방 | 캠프 | 프로젝트NO | 신청내용 |",
        "|---|---|---|---|---|",
    ]
    for row in report.get("events") or []:
        safe = lambda v: str(v or "").replace("|", "\\|").replace("\n", " ")  # noqa: E731
        lines.append("| %s | %s | %s | %s | %s |" % (
            safe(row.get("카톡일시")), safe(row.get("방")), safe(row.get("캠프명")),
            safe(row.get("프로젝트NO")) or "배정 전", safe(row.get("신청내용"))))
    if not report.get("events"):
        lines.append("| - | - | - | - | 반영 대상 없음 |")
    lines.extend(["", "## 대표 지시 이행현황", ""])
    # 확인된 지시는 화면·캡처에서 내리되 여기서는 아래 칸에 남긴다. 통째로 지우면
    # "그때 정말 그런 지시가 있었나"와 확인 근거를 같이 잃는다.
    done = [r for r in (report.get("directives") or [])
            if isinstance(r, dict) and str(r.get("상태") or "") in ("완료", "확인완료")]
    pending = [r for r in (report.get("directives") or []) if r not in done]
    if pending:
        for row in pending:
            lines.extend([
                f"### {row.get('제목') or '쿠팡 업무지시'}",
                "",
                f"- 게시: {row.get('카톡일시') or '-'} · {row.get('게시자') or '-'}",
                f"- 보고기한: **{row.get('보고기한') or '원문 확인 필요'}**",
                f"- 상태: **{row.get('상태') or '경과보고 확인 필요'}**",
                "",
            ])
            for item in row.get("항목") or []:
                lines.append("- **%s** — %s · 제출: %s · 상태: **%s**" % (
                    item.get("제목") or "항목", item.get("요구내용") or "-",
                    item.get("제출내용") or "-", item.get("상태") or "결과 미확인"))
    elif done:
        lines.append("- 미처리 지시 없음 (아래 %d건은 확인 완료)" % len(done))
    else:
        lines.append("- 확인된 쿠팡 대표 지시 없음")
    if done:
        lines.extend(["", "### 확인 완료 — 화면·캡처에서 내림", ""])
        for row in done:
            lines.append("- **%s** — 게시 %s · 확인 %s %s · 근거 %s · `%s`" % (
                row.get("제목") or "쿠팡 업무지시", row.get("카톡일시") or "-",
                row.get("확인시각") or "-", row.get("확인자") or "-",
                row.get("확인근거") or "-", row.get("레코드ID") or "-"))
        lines.append("")
        lines.append("> 되돌리려면 `python ceo_events.py --unack <레코드ID> --sync`")
    lines.extend(["", "## 날짜별 판정", "", "| 날짜 | 쿠팡 건 | 쿠팡 무관 | 모름 |",
                  "|---|---:|---:|---:|"])
    for day, values in (report.get("날짜별") or {}).items():
        lines.append(f"| {day} | {values.get('쿠팡 건', 0)} | "
                     f"{values.get('쿠팡 무관', 0)} | {values.get('모름', 0)} |")
    return "\n".join(lines) + "\n"


def _atomic_write(path, body):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(body)
    os.replace(tmp, path)


def sync(paths=None):
    report = extract(paths)
    _atomic_write(REPORT_JSON, json.dumps(report, ensure_ascii=False, indent=2))
    _atomic_write(REPORT_MD, render(report))
    return report


def load_cached(path=None):
    try:
        raw = json.load(open(path or REPORT_JSON, encoding="utf-8"))
        return raw if isinstance(raw, dict) else {"상태": "확인못함", "events": [], "directives": []}
    except (OSError, ValueError, TypeError):
        return {"상태": "확인못함", "이유": "대표대화 추출 리포트 없음", "events": [],
                "directives": [], "날짜별": {}}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--sync", action="store_true", help="판정 리포트를 reports/에 갱신")
    parser.add_argument("--list", action="store_true", help="지시 목록과 레코드ID 만 본다")
    parser.add_argument("--ack", metavar="레코드ID", help="처리 완료로 확인 — 화면·캡처에서 내린다")
    parser.add_argument("--unack", metavar="레코드ID", help="확인을 되돌려 다시 올린다")
    parser.add_argument("--who", default="관리자", help="확인한 사람(리포트에 남는다)")
    parser.add_argument("--why", default="사람 확인", help="확인 근거 한 줄")
    args = parser.parse_args()

    if args.list:
        rows = load_cached().get("directives") or []
        for row in rows:
            print("%s  %s  상태=%s  기한=%s" % (
                row.get("레코드ID"), row.get("제목"), row.get("상태"),
                row.get("보고기한") or "-"))
        if not rows:
            print("지시 없음 — 리포트가 없거나 아직 안 돌았습니다 (python ceo_events.py --sync)")
        return
    if args.ack or args.unack:
        save_ack(args.ack or args.unack, who=args.who, why=args.why,
                 remove=bool(args.unack))
        print(("확인 처리: %s — 화면·캡처에서 내려갑니다" if args.ack
               else "확인 취소: %s — 다시 올라옵니다") % (args.ack or args.unack))
        if not args.sync:
            print("→ 리포트에 반영하려면 --sync 를 같이 주세요")

    report = sync() if args.sync else extract()
    c = report["판정"]
    print("대표 대화 판정 — 쿠팡 건 %d · 쿠팡 무관 %d · 모름 %d" %
          (c["쿠팡 건"], c["쿠팡 무관"], c["모름"]))
    if args.sync:
        print("리포트:", REPORT_MD)


if __name__ == "__main__":
    main()
