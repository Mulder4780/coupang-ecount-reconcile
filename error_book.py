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
  ④ **재발 방지** — 회차가 매일 세어 네 갈래로 가른다:
     **회귀**(막았다는데 또 났다 · 맨 위) · **사전에 없음**(★새 오류) · 아는 오류 ·
     **고친 뒤 재발 없음**(마지막 발생이 고침보다 앞선 것 — 경보에서 내린다, `[288]`).

읽기 전용이다 — 아무것도 안 고치고 큐에도 안 넣고 엑셀은 열지도 않는다
(`typo_watch`·`truth_watch` 와 같은 자리). 무엇이 맞는지는 사람만 안다.
"""
import os
import sys
import re
import json
import glob
import hashlib
from datetime import datetime, timedelta, timezone

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
        # ★ **'권한 없음' 보다 먼저 온다.** 더 좁은 규칙이 앞이다(분개장 판정이 '적요'
        #   판정보다 먼저 오는 것과 같은 이유). 뒤에 두면 관리자 본인이 인증만 풀린
        #   것도 '권한이 없습니다' 로 읽혀 **고칠 길이 화면 어디에도 없어진다.**
        # [290] 이 두 문구를 갈랐다 — 담당자 화면(ADMIN_ONLY)이냐 인증 만료
        #   (ADMIN_AUTH_NEEDED)냐. 옛 문구도 같이 적어 둔다(옛 화면을 쓰는 분).
        "이름": "관리자만 되는 일",
        "when": ["관리자 전용 기능입니다", "관리자 인증이 필요합니다"],
        "쉬운말": "이 단추는 관리자만 누를 수 있습니다. 잘못 누르신 것이 아니고, "
                "앱이 고장 난 것도 아닙니다.",
        "왜": "실행·정책·전체 원장처럼 되돌리기 어려운 일은 관리자 화면에서만 합니다. "
              "관리자인데도 이 말이 나오면 기기 인증이 만료된 것입니다.",
        "하세요": ["담당자 업무센터로 여신 것이면 관리자에게 부탁합니다.",
                 "관리자이시면 [설정] → 관리자 PIN 으로 다시 인증합니다.",
                 "다시 인증해도 같으면 이 화면을 캡처해서 보냅니다."],
        "막음": "[290]", "고친날": "2026-08-16",
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
                 "(HTTP 502", "(HTTP 503", "(HTTP 504",
                 # ★ 세 번째 말투 — [197] 이 **다시 걸어 보다 끝내 못 받았을 때** 화면에
                 #   내는 문구다(`promise · 캘린더/업무센터 최신 자료를 …`). 위 주석이
                 #   말한 그대로 한쪽만 적어 뒀더니 이것만 매일 '★새 오류'로 올라왔다.
                 "최신 자료를 불러오지 못했습니다"],
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
    {
        # 2026-08-17 실측: `/api/flow · HTTP_ERROR:500` 15건인데 **사유가 한 글자도
        # 없었다.** 전역 가드가 `{"error": str(e)}` 를 보내는데 `str(e)` 는 비어 있을
        # 수 있어서다(`Exception()` 처럼 인자 없이 오른 예외). 그러면 화면에는
        # `HTTP_ERROR:500` 다섯 글자만 남고 **왜인지 영영 알 수 없다** — [109] 와 같은
        # 자리다. 이제 갈래 이름과 터진 자리를 반드시 싣는다.
        #
        # ★ `막음 [292]` 의 뜻은 **"500 이 안 난다"가 아니라 "문구 없는 500 이 안
        #   나간다"** 이다. 그러므로 이 뒤에 500 이 또 나면 그것은 회귀가 맞고,
        #   이번에는 **사유를 달고** 오므로 사람이 고칠 자리를 찾을 수 있다.
        "이름": "서버가 처리하다 멈춤",
        "when": ["HTTP_ERROR:500", "(HTTP 500"],
        "쉬운말": "앱 서버가 그 요청을 처리하다 멈췄습니다. "
                "잘못 누르신 것이 아니고, 적으신 내용도 그대로 남아 있습니다.",
        "왜": "서버 쪽에서 예상 못 한 일이 생겼습니다. 끊긴 것과 달리 "
              "다시 눌러도 같은 자리에서 또 멈추는 것이 보통입니다.",
        "하세요": ["같은 단추를 여러 번 누르지 않습니다 — 같은 자리에서 또 멈춥니다.",
                 "이 화면을 캡처해서 관리자에게 보냅니다 "
                 "(오류 문구에 적힌 파일·줄 번호가 고칠 자리를 가리킵니다).",
                 "급하시면 다른 화면으로 먼저 업무를 이어 갑니다."],
        "막음": "[292]", "고친날": "2026-08-17",
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
        ux = ledger_db.ux_summary(days=days, limit=400)
        # 시각을 주는 키를 먼저 본다. 옛 ledger_db 면 예전처럼 날짜 없이 세고
        # 그 사실을 '못 본 것'이 말한다 — 못 읽었다고 0 건으로 만들지 않는다([169]).
        for row in (ux.get("오류최근") or ux.get("오류") or []):
            # ★ 여기서 **한 시계로 모은다**([288]). 예전에는 `[:19]` 로 잘라 UTC 를
            #   그대로 담았고, 보관본(이 PC 시각)과 문자열로 비교돼 9시간이 어긋났다.
            ux_rows.append({"target": row[0], "detail": row[1], "n": int(row[2]),
                            "ts": to_local(row[3]) if len(row) > 3 else ""})
    except Exception as exc:                     # noqa: BLE001
        ux_err = str(exc)[:160]

    counts, samples, last_ts = {}, {}, {}
    날짜모름 = set()
    for r in rows:
        sig = r.get("sig") or signature(r.get("target"), r.get("detail"))
        counts[sig] = counts.get(sig, 0) + 1
        samples.setdefault(sig, r)
        ts = to_local(r.get("ts"))               # 보관본도 같은 시계로([288])
        if ts and ts > last_ts.get(sig, ""):
            last_ts[sig] = ts
    for r in ux_rows:
        sig = signature(r["target"], r["detail"])
        counts[sig] = counts.get(sig, 0) + r["n"]
        samples.setdefault(sig, {"target": r["target"], "detail": r["detail"], "ts": ""})
        ts = r.get("ts") or ""
        if ts:
            if ts > last_ts.get(sig, ""):
                last_ts[sig] = ts
        else:
            날짜모름.add(sig)                      # 그 줄만 날짜가 없다

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

    # ── 고친 뒤에도 났나 — **경보를 내리는 유일한 근거**([288]) ────────────────
    # ★ `[108]` 실측: 회귀·새 오류 넷이 매일 P1 으로 올라왔는데 넷 다 **고침보다
    #   앞선 기록**이었다. 경보가 대부분 가짜면 나머지도 아무도 안 본다([170]).
    # ★ 그러나 '안 났다'와 '아무도 그 화면을 안 열었다'는 다른 말이다([169]) —
    #   그래서 고침 뒤 화면을 연 횟수를 같이 세고, 0 이면 **경보에서 내리되
    #   '못 본 것'에 올려** 사라지지 않게 한다.
    고쳐짐 = []
    for 통 in (회귀, 새오류):
        for x in list(통):
            seen = x.get("마지막") or ""
            if not seen:                          # 때를 모르면 판정 자체를 안 한다
                continue
            ev = fix_evidence(x)
            when = ev.get("고친때")
            if not when or seen >= when:          # 근거가 없거나 고친 뒤에도 났다 → 그대로
                x["고침판정"] = ev.get("왜못함") or ("고친 뒤에도 났다 (고침 %s)" % when)
                continue
            n, 왜 = views_since(when)
            x["고친때"], x["고침근거"], x["방문"] = when, ev.get("근거", ""), n
            x["확인됨"] = bool(n)
            x["못센이유"] = 왜
            고쳐짐.append(x)
            통.remove(x)
    고쳐짐.sort(key=lambda x: -x["건수"])

    못본것 = []
    for x in 고쳐짐:
        if x.get("확인됨"):
            continue
        # 재발이 없는 것인지 아무도 안 본 것인지 **가릴 근거가 없다** — 남긴다.
        이름 = x.get("사전") or (x.get("무엇") or x.get("지문"))
        못본것.append(f"'{이름}' {x['건수']}건은 {x['고친때'][:16].replace('T',' ')} 고침 뒤"
                     f" 기록이 없는데 그 뒤 화면을 연 기록도 "
                     f"{x.get('못센이유') or '0회'} — 재발 없음을 확언 못 한다")
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
            "고쳐짐": 고쳐짐,
            "보관본건수": len(rows), "ux건수": sum(r["n"] for r in ux_rows)}


def _ago(ts):
    """'언제 났나' 를 사람 말로. **모르면 모른다고 한다** — 빈 값을 '방금' 이라 적으면
    이틀 전에 끝난 고장이 새 고장으로 읽힌다([169]).
    ※ 시각은 `to_local()` 이 이미 한 시계로 모아 둔다([288]) — 예전처럼 UTC 와
      이 PC 시각이 섞이지 않으므로 분 단위로 말해도 된다."""
    ts = str(ts or "")[:19]
    if not ts:
        return "때 모름"
    보임 = ts[:16].replace("T", " ")
    try:
        t = datetime.fromisoformat(ts)
    except Exception:                            # noqa: BLE001
        return 보임
    d = (datetime.now() - t).days
    return f"{보임} · 오늘" if d <= 0 else f"{보임} · {d}일 전"


# ── 시각은 한 시계로 모은다 ─────────────────────────────────────────────────
def to_local(ts):
    """어느 시계로 적힌 시각이든 **이 PC 시각**으로 옮긴다. 못 읽으면 원본 그대로.

    ★ 실측 2026-08-16: ux 표 42,339행 중 **42,322행이 `…Z`(UTC)** 다(99.96%).
      `ux_add` 가 화면이 보낸 `ts` 를 그대로 넣기 때문이고, 보관본(`archive`)은
      `datetime.now()` 라 **이 PC 시각**이다. 지금까지 둘을 **문자열로 그냥
      비교**해 왔는데, `_ago` 가 '며칠 전' 까지만 말했으므로 9시간이 어긋난 채로도
      아무 티가 안 났다.
    ★ 그런데 '**고친 뒤에도 났나**'를 물으려면 분 단위가 맞아야 한다. 안 맞추면
      고침 직후 **9시간이 통째로 눈먼 창**이 되어, 고친 뒤에 난 오류를 '고쳐졌다'고
      말한다 — `[169]` 중에서도 제일 나쁜 쪽(안심시키는 거짓)이다.
      실측이 그 크기를 보여 준다: `fmtDateTime` 마지막 발생 `11:41:10Z` 는
      이 PC 시각으로 **20:41** 이고 고침 커밋은 **20:51** — 실제 간격은 9시간이
      아니라 **10분**이었다. 안 맞추면 그 10분이 9시간으로 보인다.
    """
    s = str(ts or "").strip()
    if not s:
        return ""
    try:
        t = datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:                            # noqa: BLE001
        return s[:19]
    if t.tzinfo is not None:                     # UTC·오프셋이 붙어 왔다 → 이 PC 시각으로
        t = t.astimezone().replace(tzinfo=None)
    return t.isoformat(timespec="seconds")


def _to_utc(local_iso):
    """이 PC 시각 → ux 표가 쓰는 `…Z` 꼴. 못 읽으면 None."""
    try:
        t = datetime.fromisoformat(str(local_iso)[:19]).astimezone()
    except Exception:                            # noqa: BLE001
        return None
    return t.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")


# ── '언제 고쳤나' 는 지어내지 않는다 — 저장소가 말하게 한다 ──────────────────
#: 오류 문구에서 '아직 없는 이름' 을 뽑는 말투. 정규식을 넓히면 엉뚱한 낱말을
#: 심볼로 읽어 **없는 고침을 찾아낸다** — 좁게 둔다([172]).
_UNDEF = re.compile(r"\b([A-Za-z_$][\w$]*) is not defined")
#: 그 이름이 지금 **정의돼 있나** 를 볼 파일. 살아 있는 앱 화면 하나다.
_DEF_FILES = ("webapp/index.html",)
#: 검증번호가 사는 곳 — 이 프로젝트는 고침에 번호를 붙인다(`[NNN]`).
_PROOF_FILE = "tests/synthetic_check.py"


def _git_when(text, path):
    """그 글자가 그 파일에서 **마지막으로 달라진** 커밋 시각(이 PC 시각). 못 찾으면 None.

    `git log -S` 는 그 글자의 **개수가 바뀐** 커밋만 고른다 — 지금 그 글자가 있다면
    그 커밋이 곧 '들어온 때' 다. 짐작이 아니라 저장소가 답한다.
    """
    try:
        import proc_guard                        # 창 없이 · 나무째 죽인다([179][272])
        res = proc_guard.run_tree(
            ["git", "log", "-1", "--format=%cI", "-S", text, "--", path],
            cwd=ROOT, timeout=25)
    except Exception:                            # noqa: BLE001
        return None
    if getattr(res, "timed_out", False) or res.returncode != 0:
        return None
    out = (res.stdout or "").strip().splitlines()
    return to_local(out[0]) if out else None


def fix_evidence(item):
    """이 오류를 **언제 고쳤나** → `{"고친때", "근거"}` · 모르면 `{"왜못함"}`.

    근거는 두 갈래뿐이고 **둘 다 저장소에 실재하는 글자**다:
      ① 사전에 `막음 [NNN]` 이 붙어 있다 → 검증 `def tNNN(` 이 들어온 커밋
      ② 문구가 `X is not defined` 다 → 그 이름이 **지금 정의돼 있고** 그 정의가
         들어온 커밋. 아직 정의가 없으면 **근거 없음**이다(고쳐지지 않았다).
    ★ 못 찾으면 `None` 이다 — '아마 고쳤겠지' 로 넘어가지 않는다. 근거가 없으면
      경보는 그 자리에 그대로 남는다(안전한 쪽).
    """
    막음 = str(item.get("막음") or "")
    when, why = None, ""
    for n in re.findall(r"\[(\d{2,4})\]", 막음):
        # ⚠ 검증 함수 이름은 `def t249(` 가 아니라 `def t249_entry_save_never_silent():`
        #   처럼 **뒤에 설명이 붙는다.** 괄호까지 붙여 찾으면 한 건도 안 걸리면서
        #   오류도 안 난다 — 그러면 그 갈래는 영영 '근거 없음'으로 남는다([165] 모양).
        t = _git_when("def t%s" % n, _PROOF_FILE)
        if t and (when is None or t > when):
            when, why = t, "검증 [%s] 이 들어온 커밋" % n
    if when:
        return {"고친때": when, "근거": why}

    m = _UNDEF.search(str(item.get("무엇") or ""))
    if m:
        name = m.group(1)
        for rel in _DEF_FILES:
            path = os.path.join(ROOT, *rel.split("/"))
            try:
                src = open(path, encoding="utf-8", errors="replace").read()
            except OSError:
                continue
            for kind in ("function", "const", "let", "var"):
                decl = "%s %s" % (kind, name)
                if re.search(r"\b%s\s+%s\b" % (kind, re.escape(name)), src):
                    t = _git_when(decl, rel)
                    if t:
                        return {"고친때": t, "근거": "`%s` 정의가 %s 에 들어온 커밋"
                                                   % (name, rel)}
                    return {"왜못함": "`%s` 는 %s 에 정의돼 있는데 그 커밋을 못 찾았다"
                                     % (name, rel)}
        return {"왜못함": "`%s` 가 아직 어디에도 정의돼 있지 않다" % name}
    return {"왜못함": "무엇을 고쳤다는 근거(검증번호·정의)가 문구에 없다"}


def views_since(local_iso):
    """고친 **뒤에** 사람이 화면을 몇 번 열었나 → `(건수, 못센이유)`.

    ★ 0 이면 '안 났다'가 아니라 **'아무도 안 봤다'** 다([169]). 그 둘을 뭉치면
      아무도 안 쓴 화면이 '고쳐진 화면' 으로 둔갑한다.
    """
    utc = _to_utc(local_iso)
    if not utc:
        return None, "고친 때를 시각으로 못 읽었다"
    try:
        import ledger_db
        with ledger_db.conn() as c:              # 읽기만 한다
            n = c.execute("SELECT COUNT(*) FROM ux WHERE kind='view' AND ts>=?",
                          (utc,)).fetchone()[0]
        return int(n), ""
    except Exception as exc:                     # noqa: BLE001
        return None, "ux 표를 못 읽었다 — %s" % str(exc)[:100]


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
                     f"\n  - 마지막 {_ago(x.get('마지막'))} · 지문 `{x['지문']}`")
        L.append("")
    if res["아는것"]:
        L += ["## 아는 오류 (사람에게 풀어서 말해 주고 있다)", ""]
        for x in res["아는것"][:20]:
            L.append(f"- {x['건수']}건 · {x['사전']} — `{x['지문']}` (마지막 {_ago(x.get('마지막'))})")
        L.append("")
    if res.get("고쳐짐"):
        L += ["## 고친 뒤 재발 없음 — 경보에서 내렸다", "",
              "마지막 발생이 **고침보다 앞선** 것들이다. 근거는 저장소가 말한 커밋 시각이고,"
              " 고친 뒤 화면을 연 횟수를 같이 적는다 — 0 이면 위 '못 본 것'에도 올라간다.", ""]
        for x in res["고쳐짐"][:20]:
            방문 = x.get("방문")
            말 = (f"그 뒤 화면 {방문}회 열림" if 방문
                 else f"그 뒤 화면을 연 기록 없음 — {x.get('못센이유') or '0회'}")
            L.append(f"- {x['건수']}건 · {x.get('사전') or x.get('무엇') or x['지문']}"
                     f"\n  - 마지막 {_ago(x.get('마지막'))} · 고침 "
                     f"{str(x.get('고친때'))[:16].replace('T', ' ')}"
                     f" ({x.get('고침근거')}) · {말}")
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
        # 막은 날 **뒤에도 났다**는 것이 회귀의 근거다 — 그 날짜를 같이 적는다.
        out.append(f"★회귀 오류 {x['건수']}건 — {x['사전']} "
                   f"(막음 {x.get('막음')} {x.get('고친날')} · 마지막 {_ago(x.get('마지막'))}) "
                   f"→ python error_book.py --print")
    # ★ '먼저 처리할 것' 은 **지금 할 일**의 목록이라 최근 것부터 싣는다.
    #   리포트는 건수 순 그대로다 — 거기서는 무엇이 잦은가가 알고 싶은 것이다.
    #   건수 순으로만 실으면 이틀 전에 끝난 16건이 오늘 난 1건을 맨 위에서 덮는다.
    #   때를 모르는 것은 뒤로 가되 **빠지지는 않는다**(사전에는 여전히 없다, [169]).
    새 = sorted(res.get("새오류", []), key=lambda x: str(x.get("마지막") or ""), reverse=True)
    for x in 새[:3]:
        out.append(f"★새 오류 {x['건수']}건 — {x['어디']} · {x['무엇'] or '사유 없음'} "
                   f"(마지막 {_ago(x.get('마지막'))}) → python error_book.py --print")
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
              f"아는 것 {len(res['아는것'])} · 고친 뒤 재발없음 {len(res.get('고쳐짐', []))} · "
              f"못 본 것 {len(res['못본것'])}"
              f"  → {os.path.relpath(REPORT_MD, ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
