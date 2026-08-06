# 문서 스캔(OCR) 도구 — 무엇이 제일 좋은가, 왜 그것을 쓰는가

2026-08-06 조사. 지시: **"문서 스캔해서 텍스트 확인하는 최고의 무료 도구 확인 해서
비교 대조 할 수 있는 알고리즘 추가해"**

## 이 프로젝트의 조건이 후보를 먼저 좁힌다

| 조건 | 왜 |
|---|---|
| **무료** | 지시 그대로 |
| **PC 안에서만 실행 (업로드 금지)** | 다루는 것이 거래명세서·세금계산서다. 사업자번호·금액·거래처가 들어 있어 외부 API 로 보내지 않는다 |
| **한글 + 숫자** | 한글 항목명과 금액·번호를 같이 읽어야 한다 |
| **표·칸이 있는 서식** | 명세서는 표다. 줄글 OCR 은 옆 칸 숫자와 들러붙는다 |
| **CPU 로도 돌아갈 것** | 이 PC 에 전용 GPU 가 없다 |

이 조건 때문에 성능이 좋아도 **클라우드 OCR(구글·네이버·MS Azure·Upstage 등)은 후보에서 빠진다.**
무료 구간이 있어도 문서가 PC 밖으로 나가기 때문이다.

## 비교 (2026-08 시점 조사)

| 도구 | 한국어 | 표/레이아웃 | 속도(CPU) | 설치 | 판정 |
|---|---|---|---|---|---|
| **PaddleOCR** (한국어 모델, PP-Structure) | **최상위** | **강함** | 보통 | 무겁다(별도 런타임) | ★ **주 엔진으로 유지** |
| **Windows.Media.Ocr** (Win10+ 내장) | 보통 | 약함 | 빠름 | **불필요** | ★ **둘째 의견으로 채택** |
| Surya | 상위 | 매우 강함 | 느림(GPU 권장) | 보통 | 대기 — GPU 붙으면 재검토 |
| Tesseract 5 (kor) | 낮음 | 약함 | 빠름 | 쉬움 | 선택 — 있으면 셋째 표로 |
| docTR / RapidOCR | 보통 | 보통 | 보통 | 보통 | PaddleOCR 대비 이점 없음 |
| VLM 계열(dots.ocr 등) | 상위 | 최상위 | **매우 느림** | GPU 필수 | 이 PC 조건에 안 맞음 |

근거: [LlamaIndex — Best OCR Libraries 2026](https://www.llamaindex.ai/blog/best-ocr-libraries-for-developers) ·
[Unstract — Best Open Source OCR Tools](https://unstract.com/blog/best-opensource-ocr-tools/) ·
[Koncile — PaddleOCR vs Tesseract](https://www.koncile.ai/en/ressources/paddleocr-analyse-avantages-alternatives-open-source) ·
[InvoiceDataExtraction — 6 Python OCR engines compared](https://invoicedataextraction.com/blog/python-ocr-library-comparison-invoices) ·
[GIGAGPU — Best OCR Models 2026](https://gigagpu.com/best-ocr-models-2026/)

## 결론 — 엔진을 바꾸지 않고, **두 번 읽는다**

조사 결과 **지금 쓰는 PaddleOCR 한국어 모델이 무료·로컬 조건에서 여전히 최상위**였다.
엔진을 갈아 끼울 이유가 없다. 그런데 정확도를 올릴 여지는 다른 곳에 있었다:

> **한 엔진의 답만 보면, 그 답이 틀렸다는 사실을 알 방법이 없다.**

그래서 `band/ocr_crosscheck.py` 를 만들었다. 원리는 하나다 —
**서로 다른 엔진 둘에게 같은 사진을 읽히고, 값이 겹칠 때만 원장에 넣는다.**

| 교차 판정 | 뜻 | 원장 입력 |
|---|---|---|
| **합치** | 두 엔진 이상이 같은 값 | 허용 (빈칸일 때만) |
| **단독** | 한 엔진만 값을 냄 | 금지 — 제안만 |
| **충돌** | 엔진마다 값이 다름 | 금지 — 사람에게 |
| 없음 | 아무도 못 읽음 | — |

Windows 내장 OCR 은 PaddleOCR 보다 정확도가 낮다. 그래도 둘째 의견으로 쓸모가 있는 이유는
**정확해서가 아니라 다른 방식으로 틀리기 때문**이다. 같은 오답을 낼 확률이 낮으므로,
둘이 같은 값을 내면 그 값이 맞을 확률이 크게 올라간다.

### 전량을 두 번 읽지는 않는다

사진이 1,816장이다. 매일 두 엔진으로 전량을 읽으면 daily_run 이 무너진다.
`needs_second_opinion()` 이 **재검할 것만** 고른다:

- 금액 정합성이 깨진 건 (공급가액 + 세액 ≠ 합계)
- 프로젝트NO · 공급가액 · 발행일 · 문서번호 중 하나라도 못 읽은 건
- 신뢰도가 '높음'이 아닌 건
- **원장에 값을 쓰려는 건 전부** — 쓰는 순간이 되돌리기 제일 비싼 지점이다

나머지는 기본 엔진 결과를 그대로 쓴다. 결과는 엔진별로 캐시되므로 두 번째 실행부터는 즉시 끝난다.

## 두 번째 대조 — 문서 ↔ 원장을 **항목마다**

기존 `doc_ocr.match()` 는 프로젝트NO 로 행을 찾고 **공급가액 하나만** 봤다.
발행일이 틀려도, 명세서번호가 달라도 "일치"로 나갔다.
`compare_ledger()` 는 유형에 따라 아래를 전부 견준다.

| 문서 | 원장 열 |
|---|---|
| 발행일 | 거래명세서발행일 / 세금계산서실제발행일·세금계산서발행일 |
| 명세서번호 · 승인번호 | 거래명세서번호 / 세금계산서승인번호 |
| 공급가액 | 실제작업공급가액 |
| 세액 | 실제작업부가세 |
| 합계 | 거래명세서합계 / 세금계산서합계 |

판정은 항목마다 `일치 · 불일치 · 원장 빈칸 · 문서에서 못 읽음 · 양쪽 빈칸` 이고,
**한 항목이라도 불일치면 그 건은 빈칸도 채우지 않는다** — 사진과 원장이 다른 건을
가리키고 있을 수 있기 때문이다.

세액·합계는 읽어서 **대조만** 하고 원장에 쓰지 않는다(수식·집계 열이다).

## 쓰는 법

```bash
python band/ocr_crosscheck.py --status          # 이 PC 에서 쓸 수 있는 엔진
python band/ocr_crosscheck.py --scan            # 재검 대상만 교차검증 → 리포트 2종
python band/ocr_crosscheck.py --scan --all      # 전량 두 번 읽기(느리다)
python band/ocr_crosscheck.py --scan --apply    # 합치 + 원장 빈칸만 입력 큐로
```

리포트: `reports/문서OCR교차검증_*.csv`(문서 1건 = 1행) ·
`reports/문서OCR항목대조_*.csv`(항목 1개 = 1행, 어디가 왜 틀렸는지) ·
`reports/문서OCR교차검증_최근.json`(앱·워치독이 읽는 요약)

자동 실행: **daily_run 09:50** 의 "문서 OCR 교차검증" 단계.

## 엔진을 하나 더 붙이려면

`ocr_crosscheck.py` 의 `ENGINES` 에 한 줄 넣으면 된다 —
`(설명, 사용가능 판정 함수, 실행 함수, 등급)`. 실행 함수는 `{경로: 텍스트}` 만 돌려주면 된다.

- **Tesseract 5**: `winget install UB-Mannheim.TesseractOCR` 후 한국어 데이터(`kor.traineddata`)
  를 넣으면 `--status` 에 자동으로 O 가 뜬다. 코드 수정 불필요.
- **Surya**: GPU 를 붙인 뒤에 검토한다. PaddleOCR 처럼 별도 파이썬 런타임 + 워커 스크립트
  방식으로 붙이는 것이 안전하다(`paddle_ocr_worker.py` 가 본보기).

## 하지 않는 것

- 외부 OCR API·업로드 (무료 구간이 있어도 하지 않는다)
- 엔진 자동 설치 (있는 것만 골라 쓴다)
- 교차검증을 통과하지 못한 값의 원장 입력
- 이미 값이 있는 칸 덮어쓰기
