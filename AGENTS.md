# AGENTS.md — AI 작업자 승계 문서 (Claude · Codex · 기타 AI 공용)

> 이 저장소에서 작업하는 모든 AI는 **이 파일을 먼저 읽고 시작할 것.**
> 관리대장 엑셀의 `19_AI작업인수인계` 시트가 원본 인수인계 대장이며, 이 파일은 코드 저장소 측 요약이다.
> 사용자(유현민)의 Claude 크레딧이 끊기면 이 문서만으로 Codex 등 다른 AI가 이어서 작업한다.

## 프로젝트 개요
(주)유니버셜리프트앤히타치코리아의 쿠팡 업무(돌발AS·정기점검·신규납품설치)를 관리하는
**관리대장 엑셀** ↔ **이카운트(ECOUNT) ERP** ↔ **네이버 밴드** 3자 자동 대조·자동화 시스템.
매일 08:30 유수비 대표 보고 체계의 데이터 무결성을 보장하는 것이 목적.

- 관리대장: `Z:\2. Cost\★★★쿠팡 업무 폴더★★★\♣ 1000. 쿠팡 통합업무관리 전산화 프로젝트\00. 대시보드 (프로젝트 일정, 담당자, 진행현황, 문제사항)\00. 쿠팡 통합업무 일일보고 관리대장\쿠팡_통합업무_일일보고_관리대장_v{N}.xlsx` (최신 v 자동탐지 지원, 구버전은 OLD/)
- 이 저장소: https://github.com/Mulder4780/coupang-ecount-reconcile (공개)
- Python: `C:\Users\hueng\AppData\Local\Programs\Python\Python312\python.exe` (PATH의 python은 Windows 스토어 스텁 — 사용 금지)

## ★ 절대 규칙 (위반 시 실사고)
1. **비밀정보**(이카운트 API키·밴드 토큰·비밀번호)는 `config/ecount_config.json`·`band/.band_token.json` 등 gitignore된 로컬 파일에만. 커밋·채팅·엑셀 입력 절대 금지. push 전 `git grep --cached`로 키 문자열 스캔.
2. **관리대장을 openpyxl로 열어 save() 금지** — 차트·도형버튼·x14 검증이 파괴된다. 읽기는 `read_only=True`만. 수정은 `workbook_patch.py`(zip 단일파트 패치, 3중 검증 내장)로 vN+1 새 파일 생성. 원본 덮어쓰기 금지.
3. **이카운트 트래픽**: 조회 1회/1초, 저장 1회/10초, 일 5,000건, 연속오류 30건/시(초과 시 ERP 자체가 제한될 수 있음 — 사용자 경고). 존재하지 않는 경로 무차별 탐침 금지. 판매·세금계산서·수금 "조회" API는 **존재하지 않음**(2026-07 삼중 확인) — 다시 찾으려 하지 말 것.
4. **실데이터·실서버 작업 전 반드시 `python tests/synthetic_check.py` 실행** — ALL GREEN이어야 진행 (사용자 상시 지시 2026-07-24).
5. ERP에 전표를 쓰는 작업(`ecount_upload.py --post`)은 **사용자 명시 승인 후에만**. 이중계상 가드(명세서번호 보유 행 제외)를 끄지 말 것(--include-billed는 사용자 지시 시에만).
6. 관리대장 공통 상수 HDR_ROW=4, FIRST=5 변경 금지. 열 삽입·삭제 금지. STANDING ORDER(19시트 60행): 작업 시 18·20 매뉴얼, 19 인수인계, 21·22, A2 사용법 현행화. AS 기사 기준 4인: 김준형·권오철·김필우·차동호.
7. 확인 안 된 쿠팡 정책 단정 금지(특히 "PO 없으면 세금계산서 불가" 단정 금지). 원자료에 없는 값 임의 채움 금지.

## 시스템 구성 (전부 표준 라이브러리+openpyxl, 독립 실행)
| 파일 | 역할 | 실행 |
|---|---|---|
| `ecount_client.py` | 이카운트 OAPI 인증(Zone→로그인→세션캐시) | 라이브러리 |
| `ecount_reconcile.py` | 원장 로드(read-only) + inbox(엑셀 내보내기) 판매·세금계산서 대조 | `python ecount_reconcile.py` |
| `erp_ledger_check.py` | ERP 거래처별계정별원장 ↔ 원장 4유형 대조(A: ERP에만/B: 원장에만/C: 회계O·계산서X/D: 금액차) | `python erp_ledger_check.py` |
| `ecount_upload.py` | 원장→ERP 매출전표 자동등록(SaveInvoiceAuto). 기본 dry-run | `python ecount_upload.py [--post --limit N]` |
| `workbook_patch.py` | 관리대장 vN→vN+1 안전 버전업(인수인계 행 추가) | `python workbook_patch.py --b 제목 --c 내용` |
| `band/band_auth.py → band_sync.py → band_reconcile.py` | 밴드 게시글 수집·원장 대조 | 순서대로 |
| `kakao/kakao_reconcile.py` | 카톡 [대화 내보내기] .txt ↔ 원장 대조 (읽기 API 없음 — 내보내기가 공식 경로, PC 클라이언트 후킹은 약관위반이라 금지) | `kakao/inbox/*.txt` 투입 후 실행 |
| `daily_run.py` (+`daily_run.bat`) | 일일 에이전트 오케스트레이터: 합성검증→전체 대조→종합리포트 1장. ERP 쓰기는 절대 자동 실행 안 함 | 작업 스케줄러 등록(bat 상단 명령) |
| `tests/synthetic_check.py` | 합성데이터 상시 검증(실서버 접촉 0) | 작업 전 필수 |

- 데이터 흐름: 이카운트 화면 엑셀 내보내기 → `inbox/` → 대조기 → `reports/` (md·csv·xlsx). 관리대장은 절대 직접 수정하지 않음.
- 회계 확정값: CUST=2548801036(쿠팡로지스틱스서비스 유한회사), CR_CODE=4049(제품매출, 변재선 과장 확인), TAX_GUBUN=11(가정—첫 실전송에서 검증).
- 전표 매칭 키: 원장 06시트 거래명세서번호("2026/07/01-4") = ERP 일자-No.("2026/07/01 -1", 공백 정규화).

## 현재 상태 (2026-07-24 기준)
- 이카운트 로그인 작동(COM_CODE 664540 / USER_ID 유현민 / ZONE AB / 허용IP 220.75.164.157). 품목조회 실데이터 7,281건 수신 확인.
- 전표 자동등록: 유상 62건 중 61건은 ERP 기등록 추정(명세서번호 보유)으로 제외, **JS-2607-041(울산2캠프, 1,472,500원) 1건 전송 대기** — 권한 게이트로 사용자 Run 승인 대기 중.
- 실사례 문제(동기): ERP 전표 2026/03/25-1 인천8MB 26,690,000원 — 회계반영O·세금계산서X·설치여부 미확인(변재선 과장 발견). `erp_ledger_check.py` 유형 A+C가 이런 건을 자동 검출.
- 밴드: 앱 "Coupang Report" 심사중(~2026-07-31 예상). 승인 후 client_id/secret → band_auth.
- 대상 밴드: band.us/band/90610953, band.us/band/84789192 (config band.대상밴드_URL).

## 사용자 대기 항목
1. `--post --limit 1` Run 승인(JS-2607-041 시험 전표) → ERP 화면 확인
2. 거래처별계정별원장(쿠팡로지스틱스, 4049, 해당월) 엑셀 → `inbox/` (파일명에 '원장' 포함) → `erp_ledger_check.py`
3. 밴드 앱 심사 결과(client_id/secret)
4. 3월 인천8MB 26,690,000원 건 현장 확인 결과(김필우·김준형 기사)

## 크레딧 중단 시 승계 절차 (Codex 등)
1. 이 파일 + 관리대장 `19_AI작업인수인계` 시트(최신 행부터)를 읽는다.
2. `python tests/synthetic_check.py` 로 환경 검증(ALL GREEN 확인).
3. "사용자 대기 항목"부터 이어서 진행. 새 작업 완료 시:
   - `workbook_patch.py`로 관리대장에 인수인계 행 추가(vN+1)
   - 이 파일의 "현재 상태" 갱신 → commit·push (비밀 스캔 후)
4. 모든 판단 근거·시행착오는 이 문서와 19시트에 남긴다 — 다음 작업자가 같은 실수를 반복하지 않게.
