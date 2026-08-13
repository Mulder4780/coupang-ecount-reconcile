#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""error_book.py — 오류를 **남기고 · 풀어 주고 · 다시 안 나게** 하는 한 곳 (2026-08-13 지시)

사용자 지시: "오류 로그 기록 다 보관하고 그 오류가 발생되었을 때 해결할 수 있게
알고리즘 정리하고 담당자가 업무 처리시 초등학생도 알아볼 수 있게 팝업 띄워주고
팝업이 뜨면 캡처해서 나한테 보내줘서 내가 너한테 보내주고 바로 잡을 수 있는 구조
만들어줘, 로그기록이 있으면 다음에 이 사항이 안발생되게 정리하고"

빠져 있던 것은 '기록하는 기능'이 아니었다. 실측 2026-08-13:
  · 앱은 이미 오류를 적고 있다 — 90일 4,130건 · 최근 3일 692건, 사유도 남는다.
  · 그런데 **그 기록은 90일이 지나면 지워진다**(`ledger_db.ux_add` 의 DELETE).
    그래서 "지난달에도 이랬나"를 물을 수가 없었다.
  · 그리고 **사람에게는 아무 말도 안 했다.** `/api/originals` 권한거부 **222건**이
    최근 3일에 쌓였는데 어느 화면에도 안 떴다 — 담당자는 원본을 못 열고 있고
    아무도 그 사실을 모른다. 조용한 사고의 전형이다([169]).
  · 사람에게 뜬 말도 개발자 말이었다: `HTTP_ERROR:400`, `VersionConflict`.

그래서 이 파일이 하는 일은 넷이다.
  ① **보관** — 오류는 `reports/오류기록/YYYY-MM.jsonl` 에 **덧붙이기만** 한다.
     지우지 않는다. ux 표(90일)는 화면용 요약이고 정본은 여기다.
  ② **사전** — 오류 지문 → 초등학생도 알아볼 말 + 이렇게 하세요 1·2·3.
     ★ **모르는 오류는 지어내지 않는다**(`look_up` 이 None 을 준다).
  ③ **신고문구** — 팝업의 [캡처해서 보내기] 가 만드는 글. 형님이 그대로 붙여넣으면
     AI 가 처음부터 뒤질 일이 없다([181] 과 같은 자리 — 왕복 한 번을 줄인다).
  ④ **재발 방지** — 회차가 매일 세어 세 갈래로 가른다:
     **회귀**(막았다는데 또 났다 · 맨 위) · **사전에 없음**(★새 오류) · 아는 오류.

읽기 전용이다 — 아무것도 안 고치고 큐에도 안 넣고 엑셀은 열지도 않는다
(`typo_watch`·`truth_watch` 와 같은 자리). 무엇이 맞는지는 사람만 안다.
"""
import os
import sys
import re
import json
import glob
import hashlib
from datetime import datetime, timedelta

if hasattr(sys.stdout, "reconfigure"):        # 무인 회차는 pythonw = stdout 이 None [235]
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

ARCHIVE_DIR = os.path.join(ROOT, "reports", "오류기록")
FAIL_MARK = os.path.join(ARCHIVE_DIR, ".보관실패.json")
REPORT_MD = os.path.join(ROOT, "reports", "오류_사전.md")
REPORT_JSON = os.path.join(ROOT, "reports", "오류_사전.json")

# ── 사전 ────────────────────────────────────────────────────────────────────
# `when` 은 (target, detail) 을 이어 붙인 한 줄에서 찾을 조각이다. 정규식이 아니라
# 조각으로 두는 이유: 사람이 규칙을 늘릴 때 정규식을 틀리면 **그 갈래만 조용히
# 안 걸린다**([184] 와 같은 함정).
#
# `막음` 은 "이건 이미 코드로 막았다"는 주장이다. **주장에는 날짜가 붙는다** —
# 그 날 뒤에도 나오면 회귀이고, 회귀는 새 오류보다 무겁다.
BOOK = [
    {
        "이름": "저장 버전 없음",
        "when": ["화면 데이터 버전이 없습니다"],
        "쉬운말": "저장하려는 내용이 '언제 것인지'를 화면이 잃어버렸습니다.",
        "왜": "목록을 연 뒤 시간이 지나 다른 곳에서 같은 건이 바뀌었거나, "
              "목록을 거치지 않고 이 화면이 열렸습니다.",
        "하세요": ["목록으로 돌아가 새로고침을 한 번 누릅니다.",
                 "같은 건을 다시 열어 값을 확인합니다.",
                 "저장을 다시 누릅니다. (입력하신 내용은 이 기기에 그대로 남아 있습니다)"],
        "막음": "[249]", "고친날": "2026-08-13",
    },
    {
        # 2026-08-13 오종현 신고: "PO+숫자 로만 되어야 한다고 계속 알림이 뜹니다."
        # 실측(원장 v595 750행) — PO 칸이 채워진 701건 중 `PO숫자/PR숫자` 633 ·
        # `PO숫자` 45 · `PR숫자/PO숫자` 23. 즉 옛 규칙이 **실재하는 값의 93.6%를
        # 낯설다고 물었다.** 규칙을 실측에 맞춰 넓혔다([252]).
        "이름": "PO번호 모양이 낯설다",
        "when": ["PO+숫자 모양이 아닙니다",            # 옛 문구(아직 옛 화면을 쓰는 분)
                 "PO' 로 시작하는 번호를 찾지 못했습니다"],
        "쉬운말": "적으신 PO번호에서 PO 번호를 찾지 못했습니다. "
                "잘못 누르신 것이 아니고, 저장이 막힌 것도 아닙니다.",
        "왜": "PO 뒤에 숫자가 붙어 있어야 나중에 쿠팡 자료와 맞춰 볼 수 있는데, "
              "그 모양이 안 보였습니다. 숫자 0 을 영문 O 로 적으면 자주 이렇게 됩니다.",
        "하세요": ["적으신 값에 PO 와 숫자가 붙어 있는지 봅니다 — "
                 "PO327948 · PO327948/PR461621 · PR482790/PO343170 셋 다 맞는 모양입니다.",
                 "맞게 적었으면 '그래도 저장'을 눌러 그대로 저장합니다.",
                 "무엇을 적을지 모르겠으면 이 화면을 캡처해서 관리자에게 보냅니다."],
        "막음": "[252]", "고친날": "2026-08-13",
    },
    {
        "이름": "권한 없음",
        "when": ["HTTP_ERROR:403", "PermissionError", "권한이 없"],
        "쉬운말": "이 자료를 볼 수 있는 권한이 지금 계정에 없습니다. 잘못 누르신 것이 아닙니다.",
        "왜": "업무센터마다 열 수 있는 자료가 정해져 있습니다.",
        "하세요": ["이 화면을 캡처해서 관리자에게 보냅니다.",
                 "관리자가 권한을 열어 줄 때까지 다른 업무를 먼저 봅니다."],
        "막음": None, "고친날": None,
    },
    {
        "이름": "로그인 만료",
        "when": ["AUTH_EXPIRED", "HTTP_ERROR:401"],
        "쉬운말": "로그인 시간이 다 됐습니다. 다시 들어오시면 됩니다.",
        "왜": "오래 열어 두면 안전을 위해 자동으로 잠깁니다.",
        "하세요": ["화면에 뜬 칸에 PIN 번호를 다시 넣습니다.",
                 "하시던 화면으로 돌아가 이어서 합니다."],
        "막음": None, "고친날": None,
    },
    {
        "이름": "다른 사람이 먼저 고침",
        "when": ["HTTP_ERROR:409", "VersionConflict", "IdempotencyConflict",
                 "record version conflict"],
        "쉬운말": "같은 건을 다른 분이 조금 전에 먼저 고쳤습니다.",
        "왜": "두 사람이 같은 칸을 동시에 고치면 나중 사람이 앞사람 것을 지웁니다. "
              "그래서 앱이 일부러 멈춰 세웠습니다.",
        "하세요": ["목록을 새로고침합니다.",
                 "지금 적혀 있는 값과 내가 넣으려던 값을 나란히 봅니다.",
                 "맞는 값으로 다시 저장합니다."],
        "막음": None, "고친날": None,
    },
    {
        # 실측 2026-08-13 15:34 — 류지영: "확인필요에 아무것도 안뜰때 뭘 어떻게
        # 해야하는걸까요!!". 그때 서버에는 확인 필요가 **208건** 있었고 화면만 0 이었다.
        # 목록을 나르는 `/api/issues` 가 끊긴 것이다(같은 창에서 15:38 에 502 가 찍혔다).
        # ★ 이 줄이 '권한 없음'·'로그인 만료' **뒤에** 있어야 한다 — 그 둘이 먼저
        #   걸려야 401·403 에 맞는 답이 나간다(사전은 위에서부터 처음 맞는 것을 준다).
        "이름": "확인 필요 목록을 못 불러옴",
        "when": ["/api/issues"],
        "쉬운말": "확인할 일 목록을 가져오지 못했습니다. 잘못 누르신 것이 아닙니다. "
                "화면에 0 이 보여도 '할 일이 없다'는 뜻이 아닙니다.",
        "왜": "목록을 나르는 길이 잠깐 끊겼습니다. 대개 이 컴퓨터의 앱을 다시 "
              "띄우는 동안(약 10초) 그렇습니다.",
        "하세요": ["화면에 뜬 '다시 불러오기' 를 한 번 누릅니다.",
                 "10초쯤 기다렸다가 한 번 더 누릅니다.",
                 "세 번째도 그대로면 이 화면을 캡처해서 관리자에게 보냅니다."],
        # ★ 막음을 붙이지 않는다 — [251] 이 막은 것은 **끊긴 것을 0 건으로 보여 주던
        #   화면**이지 끊김 자체가 아니다. 붙이면 이 길이 끊길 때마다(실측 11일 103건)
        #   매일 '회귀'로 올라와 진짜 회귀를 덮는다([170]).
        "막음": None, "고친날": None,
    },
    {
        "이름": "서버가 잠깐 끊김",
        # ★ 같은 오류가 두 말투로 온다: 지문(`HTTP_ERROR:502`)과 사람 말
        #   (`서버 요청이 실패했습니다 (HTTP 502).`). 한쪽만 적으면 나머지가 매일
        #   '★새 오류'로 올라와 진짜 새 오류를 덮는다 — 실측 40건이 그랬다.
        "when": ["HTTP_ERROR:502", "HTTP_ERROR:503", "HTTP_ERROR:504", "NETWORK_ERROR",
                 "Failed to fetch", "서버에 닿지 못함", "서버에 연결할 수 없습니다",
                 "(HTTP 502", "(HTTP 503", "(HTTP 504"],
        "쉬운말": "앱이 잠깐 끊겼습니다. 대개 몇 초 뒤 저절로 돌아옵니다.",
        "왜": "PC 의 앱 서버를 다시 띄우는 동안(약 9초) 폰은 잠시 답을 못 받습니다.",
        "하세요": ["10초쯤 기다립니다.",
                 "화면에 '전부 다시 시도' 단추가 보이면 그것을 한 번 누릅니다.",
                 "1분이 지나도 그대로면 캡처해서 보냅니다."],
        # ★ 막음을 붙이지 않는다. [197] 은 **끊겨도 잘 다루게** 고친 것이지
        #   **안 끊기게** 한 것이 아니다. 붙였더니 첫판에 76갈래 중 56이 '회귀'로
        #   나왔다 — 경보가 대부분이면 아무도 안 본다([170]).
        "막음": None, "고친날": None,
    },
    {
        "이름": "필수 값이 빠짐",
        "when": ["HTTP_ERROR:400"],
        "쉬운말": "넣으신 내용 중 하나가 앱이 받을 수 있는 모양이 아닙니다.",
        "왜": "빈칸이거나, 숫자 칸에 글자가 들어갔거나, 사유를 적어야 하는 칸입니다.",
        "하세요": ["팝업에 적힌 칸 이름을 봅니다.",
                 "그 칸만 고쳐서 다시 저장합니다.",
                 "칸 이름이 안 보이면 캡처해서 보냅니다."],
        "막음": None, "고친날": None,
    },
    {
        "이름": "아직 저장 전(오프라인 보관)",
        "when": ["queued", "이 기기에 보관"],
        "쉬운말": "인터넷이 없어서 이 휴대폰 안에만 넣어 뒀습니다. 아직 저장된 게 아닙니다.",
        "왜": "서버에 닿지 못하면 입력을 잃지 않으려고 기기에 잠깐 담아 둡니다.",
        "하세요": ["인터넷이 되는 곳으로 갑니다.",
                 "앱을 열어 두면 저절로 올라갑니다.",
                 "화면 위 숫자가 0 이 되면 다 올라간 것입니다."],
        # 이것은 고장이 아니라 **정상 상태**다. 막을 대상이 아니므로 회귀도 아니다.
        "막음": None, "고친날": None,
    },
    {
        "이름": "응답을 못 읽음",
        "when": ["INVALID_JSON", "UNEXPECTED_CONTENT_TYPE", "BODY_READ_ERROR", "EMPTY_BODY",
                 "Failed to execute 'json'", "Unexpected end of JSON"],
        "쉬운말": "앱이 받은 답이 중간에 잘렸습니다.",
        "왜": "연결이 불안하거나 앱 서버가 마침 다시 뜨는 중입니다.",
        "하세요": ["잠시 뒤 다시 눌러 봅니다.",
                 "두 번 더 같으면 캡처해서 보냅니다. (화면에 있던 값은 그대로 둡니다)"],
        # ★ 막음을 붙이지 않는다. [197] 은 **끊겨도 잘 다루게** 고친 것이지
        #   **안 끊기게** 한 것이 아니다. 붙였더니 첫판에 76갈래 중 56이 '회귀'로
        #   나왔다 — 경보가 대부분이면 아무도 안 본다([170]).
        "막음": None, "고친날": None,
    },
    {
        # 실측 2026-08-13 — 사람이 값을 고치려다 막힌 것이지 고장이 아니다.
        # 고장과 섞어 두면 진짜 고장이 안 보인다.
        "이름": "정정 사유가 필요함",
        "when": ["정정 사유를 입력해 주세요"],
        "쉬운말": "이미 적혀 있던 값을 바꾸려면 '왜 바꾸는지'를 한 줄 적어야 합니다.",
        "왜": "금액·PO번호·계산서처럼 돈이 걸린 칸은 나중에 '누가 왜 바꿨나'를 "
              "찾을 수 있어야 해서 사유를 받습니다.",
        "하세요": ["팝업에 적힌 칸 이름을 봅니다(예: PO번호).",
                 "화면 아래 '사유' 칸에 짧게 적습니다 (예: 쿠팡에서 번호 정정 통보).",
                 "저장을 다시 누릅니다."],
        "막음": None, "고친날": None,
    },
    {
        # 2026-08-13 류지영 요청("삭제할수있게 해주세요!!")으로 삭제·청구제외가 생겼다.
        # 사유를 안 적으면 막는데, 그것은 고장이 아니라 **일부러 세운 문**이다.
        "이름": "삭제·청구 제외에 사유가 필요함",
        "when": ["삭제 사유를 적어 주세요", "청구 제외는 사유를 함께 적어야 합니다"],
        "쉬운말": "이 건을 목록에서 빼려면 '왜 빼는지'를 한 줄 적어야 합니다. "
                "잘못 누르신 것이 아닙니다.",
        "왜": "몇 달 뒤 '이건 왜 청구가 안 나갔나'를 묻는 사람이 반드시 있습니다. "
              "그때 답할 수 있는 것은 지금 적어 두는 이 한 줄뿐입니다.",
        "하세요": ["팝업의 사유 칸에 짧게 적습니다 "
                 "(예: 정기점검과 동시 진행 — 돌발AS로 청구 안 함).",
                 "다시 누릅니다."],
        "막음": None, "고친날": "2026-08-13",
    },
    {
        # 삭제와 청구제외를 헷갈려 물어 오는 것이 실제로 잦을 자리다.
        "이름": "삭제와 청구 제외 중 무엇을 고를까",
        "when": ["삭제와 제외", "청구 제외", "어떤 걸 눌러야"],
        "쉬운말": "둘은 다른 일입니다. "
                "다녀오셨으면 [청구 제외], 애초에 잘못 등록된 건이면 [삭제]입니다.",
        "왜": "청구 제외는 '갔다 왔지만 이 건으로는 돈을 안 받는다'는 뜻이라 "
              "기사 실적·현장 기록이 그대로 남습니다. 삭제는 '이 건은 없던 것'이라 "
              "목록에서 통째로 빠집니다(되살릴 수는 있습니다).",
        "하세요": ["정기점검과 같이 가서 돌발AS로만 청구를 안 하는 것이면 [청구 제외].",
                 "같은 건이 두 번 등록됐거나 잘못 만든 건이면 [삭제].",
                 "잘못 눌렀으면 [삭제된 건 보기] → [되살리기] 로 되돌립니다."],
        "막음": None, "고친날": "2026-08-13",
    },
]

UNKNOWN = {
    "이름": None,
    "쉬운말": "앱이 처음 보는 오류입니다. 그래서 무엇 때문인지 아직 말씀드릴 수 없습니다.",
    "왜": "이 오류는 사전에 아직 없습니다.",
    "하세요": ["이 화면을 그대로 캡처합니다.",
             "관리자에게 캡처를 보냅니다.",
             "고쳐질 때까지는 다른 업무를 먼저 봅니다."],
    "막음": None, "고친날": None,
}


def _sig_text(target, detail):
    return (str(target or "") + " · " + str(detail or "")).strip()


def look_up(target, detail=""):
    """오류 지문 → 사전 항목. **모르면 None** — 지어내지 않는다."""
    line = _sig_text(target, detail)
    if not line:
        return None
    for ent in BOOK:
        for frag in ent["when"]:
            if frag in line:
                return ent
    return None


def help_for(target, detail=""):
    """앱이 그대로 그릴 수 있는 모양. 모르는 오류도 **빈손으로 안 보낸다**."""
    ent = look_up(target, detail)
    known = ent is not None
    ent = ent or UNKNOWN
    return {
        "앎": known,
        "이름": ent["이름"] or "처음 보는 오류",
        "쉬운말": ent["쉬운말"],
        "왜": ent["왜"],
        "하세요": list(ent["하세요"]),
        "이미막음": ent["막음"],
        "지문": signature(target, detail),
        "신고문구": report_text(target, detail),
    }


def signature(target, detail=""):
    """같은 오류를 같은 것으로 세기 위한 열쇠. 숫자·아이디는 지운다 —
    안 지우면 `AS-2601-574` 하나하나가 다른 오류가 되어 **아무것도 안 모인다**."""
    line = _sig_text(target, detail)
    line = re.sub(r"\d{2,}", "#", line)
    line = re.sub(r"\s+", " ", line).strip()[:160]
    return line or "(빈 지문)"


def report_text(target, detail="", extra=None):
    """형님이 그대로 붙여넣을 글. 여기에 **이미 아는 것**을 다 담아 두면
    AI 가 처음부터 뒤질 일이 없다([181] 과 같은 절약)."""
    ent = look_up(target, detail)
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    lines = [
        "[앱 오류 신고]",
        f"때: {now}",
        f"어디: {target or '(모름)'}",
        f"무엇: {detail or '(사유 없음)'}",
        f"지문: {signature(target, detail)}",
        f"사전: {ent['이름'] if ent else '★ 사전에 없음 — 새 오류'}",
    ]
    if ent and ent.get("막음"):
        lines.append(f"※ 이 오류는 {ent['고친날']} 에 {ent['막음']} 로 막았다고 적혀 있습니다"
                     " — 또 났다면 회귀입니다.")
    for k, v in (extra or {}).items():
        lines.append(f"{k}: {v}")
    return "\n".join(lines)


# ── ① 보관 — 덧붙이기만 한다 ────────────────────────────────────────────────
def archive(events):
    """오류 기록을 달마다 한 파일에 덧붙인다. **지우지 않는다.**
    실패하면 조용히 넘어가지 않고 자국을 남긴다 — 보관이 안 되고 있다는 사실
    자체가 조용하면, 나중에 '기록이 없다'가 '오류가 없었다'로 읽힌다([169])."""
    rows = [e for e in (events or []) if str((e or {}).get("kind") or "") == "error"]
    if not rows:
        return 0
    try:
        os.makedirs(ARCHIVE_DIR, exist_ok=True)
        path = os.path.join(ARCHIVE_DIR, datetime.now().strftime("%Y-%m") + ".jsonl")
        with open(path, "a", encoding="utf-8") as f:
            for e in rows:
                d = dict(e or {})
                rec = {
                    "ts": d.get("ts") or datetime.now().isoformat(timespec="seconds"),
                    "target": str(d.get("target") or "")[:160],
                    "detail": str(d.get("detail") or "")[:400],
                    "who": str(d.get("who") or "")[:40],
                }
                rec["sig"] = signature(rec["target"], rec["detail"])
                f.write(json.dumps(rec, ensure_ascii=False) + "\n")
        if os.path.exists(FAIL_MARK):
            os.remove(FAIL_MARK)                 # 고쳐졌으면 옛 자국을 남기지 않는다
        return len(rows)
    except Exception as exc:                     # noqa: BLE001 — 여기서 죽으면 UX 수집이 통째로 막힌다
        try:
            os.makedirs(ARCHIVE_DIR, exist_ok=True)
            with open(FAIL_MARK, "w", encoding="utf-8") as f:
                json.dump({"때": datetime.now().isoformat(timespec="seconds"),
                           "왜": str(exc)[:300], "못담은건수": len(rows)}, f, ensure_ascii=False)
        except Exception:
            pass
        return -1                                # 0 이 아니다 — '못 담았다'는 다른 말이다


def _read_archive(days):
    """보관본을 읽는다. 못 읽으면 **0 이 아니라 '못 읽음'** 으로 돌려준다."""
    since = datetime.now() - timedelta(days=days)
    rows, unreadable = [], []
    for path in sorted(glob.glob(os.path.join(ARCHIVE_DIR, "*.jsonl"))):
        try:
            with open(path, encoding="utf-8") as f:
                for ln in f:
                    ln = ln.strip()
                    if not ln:
                        continue
                    try:
                        r = json.loads(ln)
                    except Exception:
                        continue
                    try:
                        if datetime.fromisoformat(str(r.get("ts"))[:19]) < since:
                            continue
                    except Exception:
                        pass
                    rows.append(r)
        except Exception as exc:                 # noqa: BLE001
            unreadable.append(f"{os.path.basename(path)} — {str(exc)[:80]}")
    return rows, unreadable


# ── ④ 재발 방지 — 세 갈래로 가른다 ──────────────────────────────────────────
def rollup(days=7):
    rows, unreadable = _read_archive(days)
    보관없음 = not rows

    # 보관본이 아직 얕다(오늘 만들었다). 그동안의 근거는 ux 표에 있다 — 같이 센다.
    ux_rows, ux_err = [], None
    try:
        import ledger_db
        for t, d, c in ledger_db.ux_summary(days=days, limit=400)["오류"]:
            ux_rows.append({"target": t, "detail": d, "n": int(c)})
    except Exception as exc:                     # noqa: BLE001
        ux_err = str(exc)[:160]

    counts, samples, last_ts = {}, {}, {}
    날짜모름 = set()
    for r in rows:
        sig = r.get("sig") or signature(r.get("target"), r.get("detail"))
        counts[sig] = counts.get(sig, 0) + 1
        samples.setdefault(sig, r)
        ts = str(r.get("ts") or "")[:19]
        if ts and ts > last_ts.get(sig, ""):
            last_ts[sig] = ts
    for r in ux_rows:
        sig = signature(r["target"], r["detail"])
        counts[sig] = counts.get(sig, 0) + r["n"]
        samples.setdefault(sig, {"target": r["target"], "detail": r["detail"], "ts": ""})
        날짜모름.add(sig)                          # ux 요약은 날짜를 안 준다

    회귀, 새오류, 아는것, 못가름 = [], [], [], []
    for sig, n in sorted(counts.items(), key=lambda kv: -kv[1]):
        s = samples.get(sig, {})
        ent = look_up(s.get("target"), s.get("detail"))
        item = {"지문": sig, "건수": n, "어디": s.get("target", ""),
                "무엇": s.get("detail", ""), "사전": (ent or {}).get("이름"),
                "마지막": last_ts.get(sig, "")}
        if ent is None:
            새오류.append(item)
            continue
        if not ent.get("막음"):
            아는것.append(item)
            continue
        # ★ 회귀는 '막음이 붙어 있다'가 아니라 **고친 날 뒤에도 났다**는 뜻이다.
        #   날짜를 안 보면 반년 전 기록까지 회귀가 되어 경보가 통째로 죽는다([170]).
        item["막음"] = ent["막음"]
        item["고친날"] = ent["고친날"]
        seen = last_ts.get(sig, "")
        if seen and seen[:10] >= str(ent["고친날"]):
            회귀.append(item)
        elif sig in 날짜모름 and not seen:
            # 날짜를 모르는 기록뿐이다 — 회귀라고도, 아니라고도 못 한다([169]).
            못가름.append(item)
        else:
            아는것.append(item)

    못본것 = []
    for x in 못가름:
        못본것.append(f"'{x['사전']}' {x['건수']}건이 날짜 없는 기록이라 "
                     f"{x['고친날']}({x['막음']}) 뒤인지 못 가렸다")
    if unreadable:
        못본것.append(f"보관본 {len(unreadable)}개를 못 읽었다: " + " · ".join(unreadable[:3]))
    if ux_err:
        못본것.append(f"ux 표를 못 읽었다 — {ux_err}")
    if os.path.exists(FAIL_MARK):
        try:
            with open(FAIL_MARK, encoding="utf-8") as f:
                m = json.load(f)
            못본것.append(f"보관이 실패한 적이 있다({m.get('때')}) — {m.get('왜','')[:80]}")
        except Exception:
            못본것.append("보관 실패 자국이 있는데 그 자국도 못 읽었다")
    if 보관없음 and not ux_rows and not 못본것:
        못본것.append("오류가 0건인지 아직 아무것도 안 담긴 것인지 가릴 근거가 없다"
                     " — 보관본이 비어 있고 ux 표도 비었다")

    return {"기간": f"최근 {days}일", "잰때": datetime.now().isoformat(timespec="seconds"),
            "합계": sum(counts.values()), "갈래": len(counts),
            "회귀": 회귀, "새오류": 새오류, "아는것": 아는것, "못본것": 못본것,
            "보관본건수": len(rows), "ux건수": sum(r["n"] for r in ux_rows)}


def _md(res):
    L = ["# 오류 사전 · 재발 감시", "",
         f"- 잰 때: {res['잰때']} · {res['기간']}",
         f"- 오류 {res['합계']}건 · {res['갈래']}갈래"
         f" (보관본 {res['보관본건수']} · ux표 {res['ux건수']})", ""]
    if res["못본것"]:
        L += ["## ⚠ 못 본 것 — '이상 없음'이 아니다", ""]
        L += [f"- {x}" for x in res["못본것"]] + [""]
    if res["회귀"]:
        L += ["## ★ 회귀 — 막았다고 적어 뒀는데 또 났다", "",
              "고쳤다는 기록이 있는데 다시 나온 것이다. 새 오류보다 무겁다.", ""]
        for x in res["회귀"][:20]:
            L.append(f"- **{x['건수']}건** · {x['사전']} (막음 {x['막음']} · {x['고친날']})"
                     f"\n  - `{x['지문']}`")
        L.append("")
    if res["새오류"]:
        L += ["## ★신규 — 사전에 없는 오류", "",
              "사람에게 '무엇 때문인지' 말해 줄 말이 아직 없다."
              " 규칙을 `error_book.BOOK` 에 한 줄 더한다.", ""]
        for x in res["새오류"][:20]:
            L.append(f"- **{x['건수']}건** · `{x['어디']}` · {x['무엇'] or '(사유 없음)'}"
                     f"\n  - 지문 `{x['지문']}`")
        L.append("")
    if res["아는것"]:
        L += ["## 아는 오류 (사람에게 풀어서 말해 주고 있다)", ""]
        for x in res["아는것"][:20]:
            L.append(f"- {x['건수']}건 · {x['사전']} — `{x['지문']}`")
        L.append("")
    if not (res["회귀"] or res["새오류"] or res["아는것"]):
        L += ["오류 기록이 없다. — **없는 것인지 안 본 것인지**는 위 '못 본 것'이 말한다.", ""]
    return "\n".join(L)


def write_report(days=7):
    res = rollup(days)
    os.makedirs(os.path.dirname(REPORT_MD), exist_ok=True)
    with open(REPORT_MD, "w", encoding="utf-8") as f:
        f.write(_md(res))
    with open(REPORT_JSON, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)
    return res


def handoff_lines(days=7):
    """인계 문서 '먼저 처리할 것' 이 읽어 가는 줄. **다시 세지 않는다** —
    회차가 써 둔 리포트를 읽는다([168])."""
    try:
        with open(REPORT_JSON, encoding="utf-8") as f:
            res = json.load(f)
    except Exception:
        return []
    out = []
    for x in res.get("회귀", [])[:3]:
        out.append(f"★회귀 오류 {x['건수']}건 — {x['사전']} (막음 {x.get('막음')}) "
                   f"→ python error_book.py --print")
    for x in res.get("새오류", [])[:3]:
        out.append(f"★새 오류 {x['건수']}건 — {x['어디']} · {x['무엇'] or '사유 없음'} "
                   f"→ python error_book.py --print")
    for x in res.get("못본것", [])[:2]:
        out.append(f"오류 감시가 못 본 것 — {x}")
    return out


def main():
    args = sys.argv[1:]
    days = 7
    if "--days" in args:
        try:
            days = int(args[args.index("--days") + 1])
        except Exception:
            pass
    if "--help-for" in args:
        i = args.index("--help-for")
        tgt = args[i + 1] if len(args) > i + 1 else ""
        det = args[i + 2] if len(args) > i + 2 else ""
        print(json.dumps(help_for(tgt, det), ensure_ascii=False, indent=1))
        return 0
    res = write_report(days)
    if "--json" in args:
        print(json.dumps(res, ensure_ascii=False, indent=1))
    elif "--print" in args:
        print(_md(res))
    else:
        print(f"오류 {res['합계']}건 · {res['갈래']}갈래 — "
              f"회귀 {len(res['회귀'])} · ★새 오류 {len(res['새오류'])} · "
              f"아는 것 {len(res['아는것'])} · 못 본 것 {len(res['못본것'])}"
              f"  → {os.path.relpath(REPORT_MD, ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
