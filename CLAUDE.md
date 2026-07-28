# 쿠팡 통합업무 자동화 프로젝트 — AI 작업 규칙

**어떤 AI(Claude Code·Codex·기타)든 이 프로젝트에서 작업을 시작하면 반드시 아래 순서를 따른다.**

## 시작 체크리스트 (매 세션)
1. `ecount/AGENTS.md` 전체 읽기 — 프로젝트 전모·절대규칙 7항·현재 상태·대기 항목
2. 관리대장 최신본(vN)의 `19_AI작업인수인계` 시트 하단 행들 읽기 — 사람·AI 공용 인수인계 원장
3. `python ecount/tests/synthetic_check.py` 실행 → **ALL GREEN 확인 후에만** 실데이터 작업
4. `python ecount/coupang_workbench.py --status` 로 시스템 상태 파악

## 종료 체크리스트 (작업을 마칠 때)
1. 의미 있는 변경이면 `ecount/AGENTS.md`의 "현재 상태"·"대기 항목" 갱신
2. `python ecount/workbook_patch.py --b "제목" --c "상세"` 로 19시트 인수인계 행 추가(vN+1)
3. git commit + push (푸시 전 `git grep --cached`로 비밀키 스캔 — AGENTS.md 규칙 1)

## Claude Code ↔ Codex **동시 작업** 규칙 (2026-07-28 지시)
- 시작 전 `python ecount/ai_claim.py --who claude --take <ledger|band|publish> --why "이유"` 로 잡는다.
  이미 상대가 잡았으면 **배타 작업은 하지 않는다** — 조회·분석으로 돌리거나 다른 일을 한다.
- **관리대장 쓰기는 절대 동시에 하지 않는다.** 각자 vN+1을 만들면 한쪽이 통째로 묻힌다.
- 우선순위: ① 사용자와 지금 대화 중인 AI ② 원장 쓰기 > 밴드/게시 > 리포트 ③ 먼저 잡은 쪽.
- push 전 `git pull --rebase`. 상대 작업을 지우지 말고 합친다.
- 끝나면 `--free`, 세션 종료 시 `--free-all`.

## Claude Code ↔ Codex 교대 규칙
- 진실의 원천은 **파일**이다: AGENTS.md(요약) + 19시트(원장) + git 이력. 특정 AI의 대화 기록·메모리에 의존하는 정보를 남기지 말 것.
- 교대 시 추가 인수인계 절차는 없다 — 위 시작 체크리스트가 곧 인수인계다.
- Claude 크레딧 소진 → Codex가 이어받고, Codex 소진 → Claude가 이어받는다. 어느 쪽이든 동일.

## 토큰 절약 상시 규칙 (별도 지시 없이 항상 적용)
- 명령 출력은 항상 좁힌다: `| tail -N`·`| head -N`·`grep`으로 필요한 줄만. 경고·진행률(`UserWarning`, 다운로드 바)은 걸러낸다.
- 파일은 부분만 읽는다(Read offset/limit, Grep). 방금 편집한 파일을 재확인용으로 다시 읽지 않는다(Edit이 실패하면 알려준다).
- 검증은 한 번에: 여러 확인을 python -c 한 줄로 묶어 실행한다. 같은 사실을 두 번 확인하지 않는다.
- 에이전트·스크립트 출력은 요약형으로 만든다(상세는 reports/ 파일에, 콘솔에는 집계 한 줄).
- 무관한 작업으로 넘어갈 때는 `/clear`를 권한다(`/compact`는 그 자체로 큰 요청이라 더 비싸다).
- 이 절약 규칙 때문에 검증을 생략하지는 않는다 — 합성검증·비밀스캔은 그대로 수행한다.

## 핵심 금지사항 (상세는 ecount/AGENTS.md)
- 관리대장을 openpyxl로 열어 save() 금지 (차트·도형 파괴) — 수정은 zip 패치 도구만
- 비밀키(config/*.json)를 커밋·채팅·엑셀에 넣기 금지
- 이카운트 무차별 API 탐침 금지(트래픽 제한으로 ERP 차단 위험)
- 실데이터 작업 전 합성검증 생략 금지
