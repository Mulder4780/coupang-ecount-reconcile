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

## Claude Code ↔ Codex 교대 규칙
- 진실의 원천은 **파일**이다: AGENTS.md(요약) + 19시트(원장) + git 이력. 특정 AI의 대화 기록·메모리에 의존하는 정보를 남기지 말 것.
- 교대 시 추가 인수인계 절차는 없다 — 위 시작 체크리스트가 곧 인수인계다.
- Claude 크레딧 소진 → Codex가 이어받고, Codex 소진 → Claude가 이어받는다. 어느 쪽이든 동일.

## 핵심 금지사항 (상세는 ecount/AGENTS.md)
- 관리대장을 openpyxl로 열어 save() 금지 (차트·도형 파괴) — 수정은 zip 패치 도구만
- 비밀키(config/*.json)를 커밋·채팅·엑셀에 넣기 금지
- 이카운트 무차별 API 탐침 금지(트래픽 제한으로 ERP 차단 위험)
- 실데이터 작업 전 합성검증 생략 금지
