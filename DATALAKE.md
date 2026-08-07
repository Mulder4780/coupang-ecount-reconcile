# 전 자료 DB화 설계 — `datalake` (2026-08-07 지시)

사용자 지시: **"모든 데이터는 DB화 해서 별도 보관하고, 앞으로 들어오는 모든 데이터 포함
변경 및 로그 기록까지 같이 정리해. 그리고 인덱스 필터링 가능한 구조,
플로우 차트 연계 가능한 구조로 코딩하는 알고리즘 구성해."**

이 문서는 **구현 전 설계 정본**이다. 코드보다 이 문서가 먼저 있어야 하는 이유는,
같은 일을 세 번 다시 설명하지 않기 위해서다(대화는 사라지고 파일만 산다).

---

## 왜 지금 필요한가 — 오늘 실제로 겪은 세 가지

1. **전수 재주사 2시간.** `source_index.py` 는 `reports/원본색인.json` 을 매번 **통째로**
   다시 만든다. Z: 의 파일이 3천 건을 넘어가면서 한 번 돌리는 데 두 시간이 걸렸고,
   그동안 아무 판단도 할 수 없었다.
2. **색인이 낡으면 거짓 경보가 난다.** `erp_grab.py` 가 "밀림 6종"을 띄웠는데
   실제로는 1종(홈택스)이었다. 자료가 없어서가 아니라 **색인이 그 자료를 아직 몰라서**였다.
   두 시간을 기다린 뒤에야 6종→1종으로 바뀌었다.
3. **변경 이력이 없다.** 같은 경로의 파일이 새 내용으로 덮이면 예전 값이 사라진다.
   "언제부터 이 값이었나"를 물으면 답할 근거가 없다.

증분 갱신·이력 보존·필터 질의는 JSON 한 장으로는 안 된다. 그래서 DB다.

---

## 자리 — **별도 DB** 하나 (`ecount/db/datalake.db`)

기존 `ecount/db/ledger_queue.db` 에 **넣지 않는다.**

- 큐 DB 는 11:00·15:00 반영이 쓰기 트랜잭션으로 잠근다. 색인 흡수는 수만 행을 쓴다.
  한 파일에 두면 **반영이 색인에 막혀 회차를 놓친다**(SQLite 는 파일 단위 쓰기 잠금).
- 큐 DB 는 "곧 엑셀로 나갈 것"만 담는 **임시 통로**고, datalake 는 **영구 보관소**다.
  수명이 다르면 백업·정리 주기도 달라야 한다.
- 두 DB 를 잇는 것은 `ATTACH DATABASE` 로 충분하다(읽기 조인).

`journal_mode=WAL`, `synchronous=NORMAL`. 워크트리에서는 **링크하지 않고 본체 경로를
집는다**(CLAUDE.md 워크트리 규칙 — `-wal` 사이드카가 갈리면 DB 가 깨진다).

---

## 표 구조

### ① `asset` — 원본 파일 한 건 = 한 행

```sql
CREATE TABLE asset(
  id        INTEGER PRIMARY KEY AUTOINCREMENT,
  path      TEXT NOT NULL UNIQUE,   -- Z: 기준 상대경로. 정본 키
  kind      TEXT NOT NULL,          -- band | erp:stmt | erp:sales | kakao | po | deposit | ...
  bucket    TEXT DEFAULT '',        -- 원본 폴더(1. ERP 내보내기 …) — 필터용
  mtime     REAL NOT NULL,          -- 원본 수정시각(증분 판정 1차)
  size      INTEGER NOT NULL,       --                      (증분 판정 1차)
  sha1      TEXT,                   -- 내용 지문(2차). mtime/size 가 같으면 계산하지 않는다
  biz_date  TEXT,                   -- 자료가 말하는 업무일(파일명·본문에서). 밀림 판정은 이것으로
  first_seen TEXT NOT NULL,
  last_seen  TEXT NOT NULL,
  gone_at    TEXT                   -- 사라진 것도 지우지 않는다. 지우면 '있었다'는 사실을 잃는다
);
CREATE INDEX ix_asset_kind_date ON asset(kind, biz_date DESC);
CREATE INDEX ix_asset_seen      ON asset(last_seen DESC);
```

**밀림 판정은 `mtime` 이 아니라 `biz_date` 로 한다.** 오늘 받은 8/4 자 자료는
"오늘 자료"가 아니다. 오늘 `erp_grab` 이 헷갈린 것이 정확히 이 지점이다.

### ② `asset_rev` — 같은 경로가 바뀔 때마다 한 행 (변경 이력)

```sql
CREATE TABLE asset_rev(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id INTEGER NOT NULL REFERENCES asset(id),
  at TEXT NOT NULL, sha1 TEXT, size INTEGER, mtime REAL,
  note TEXT DEFAULT ''
);
```

`sha1` 이 달라졌을 때만 쌓는다. 같은 파일을 백 번 봐도 한 행도 늘지 않는다.

### ③ `record` — 파일에서 **뽑아낸 업무 레코드**

파일 한 장에 밴드 글 한 개일 수도, ERP 행 500 개일 수도 있다. 그래서 층을 나눈다.

```sql
CREATE TABLE record(
  id       INTEGER PRIMARY KEY AUTOINCREMENT,
  kind     TEXT NOT NULL,        -- band_post | erp_slip | erp_stmt | kakao_msg | po | deposit
  natural_key TEXT NOT NULL,     -- 밴드 '90610953/5437', ERP 전표번호 등. 재수집해도 같아야 한다
  asset_id INTEGER REFERENCES asset(id),
  biz_date TEXT,
  party    TEXT DEFAULT '',      -- 거래처/현장 — 가장 자주 거는 필터
  amount   INTEGER,
  status   TEXT DEFAULT '',
  payload  TEXT NOT NULL,        -- 종류마다 다른 나머지 전부(JSON). 표를 늘리지 않는 자리
  hash     TEXT NOT NULL,        -- payload 지문 — 바뀌었나 판정
  created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
  UNIQUE(kind, natural_key)
);
CREATE INDEX ix_rec_date  ON record(biz_date DESC);
CREATE INDEX ix_rec_party ON record(party, biz_date DESC);
CREATE INDEX ix_rec_kind  ON record(kind, biz_date DESC);
```

### ④ `record_rev` — 값이 바뀐 이력 (누가·언제·무엇을→무엇으로)

```sql
CREATE TABLE record_rev(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  record_id INTEGER NOT NULL REFERENCES record(id),
  at TEXT NOT NULL, who TEXT NOT NULL,     -- app | claude:<sid> | codex | tool:<이름>
  field TEXT NOT NULL, old TEXT, new TEXT,
  why TEXT DEFAULT ''
);
```

### ⑤ `event` — **모든 로그가 여기 한 곳으로** (append-only)

수집·분류·대조·반영·오류·사람 조작이 지금은 `reports/*.json` 여기저기 흩어져 있다.
그래서 "8/5 돌발AS 가 왜 1건이었나"를 되짚는 데 반나절이 든다.

```sql
CREATE TABLE event(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  at   TEXT NOT NULL,
  who  TEXT NOT NULL,          -- 세션·도구 식별자
  area TEXT NOT NULL,          -- collect | intake | match | ledger | app | error
  action TEXT NOT NULL,        -- band.grab | erp.excel | queue.apply | ...
  ok   INTEGER NOT NULL DEFAULT 1,
  ref_kind TEXT, ref_id INTEGER,   -- asset/record 를 가리킬 때
  detail TEXT DEFAULT ''           -- JSON
);
CREATE INDEX ix_ev_at   ON event(at DESC);
CREATE INDEX ix_ev_area ON event(area, at DESC);
```

**UPDATE·DELETE 를 트리거로 막는다.** 로그가 고쳐질 수 있으면 근거가 아니다.

```sql
CREATE TRIGGER ev_no_update BEFORE UPDATE ON event
BEGIN SELECT RAISE(ABORT,'event 는 고칠 수 없다'); END;
CREATE TRIGGER ev_no_delete BEFORE DELETE ON event
BEGIN SELECT RAISE(ABORT,'event 는 지울 수 없다'); END;
```

### ⑥ `link` — 레코드끼리 잇기 = **플로우차트의 간선**

```sql
CREATE TABLE link(
  src INTEGER NOT NULL REFERENCES record(id),
  dst INTEGER NOT NULL REFERENCES record(id),
  rel TEXT NOT NULL,           -- po→as | as→slip | slip→invoice | invoice→deposit
  conf REAL DEFAULT 1.0,       -- 대조 확신도(자동 매칭이면 1 미만)
  by   TEXT DEFAULT '', at TEXT NOT NULL,
  PRIMARY KEY(src, dst, rel)
);
CREATE INDEX ix_link_dst ON link(dst);
```

이 표 하나가 "PO → 밴드 AS → ERP 전표 → 세금계산서 → 입금" 전 구간을 담는다.
지금 `db/ledger_queue.db` 의 `flow_step`(12행, AS 처리 단계 정의)은 **틀**이고,
`link` 는 **실제로 흐른 자국**이다. 둘을 맞대면 "어느 단계에서 멈췄나"가 나온다.

---

## 인덱스 필터링 — 사람이 쓰는 모양

```
python ecount/datalake.py --find kind=band_post since=2026-08-01 party=현대 q="정기점검"
python ecount/datalake.py --find kind=erp_stmt since=2026-08-01 amount'>'1000000
python ecount/datalake.py --log area=collect since=2026-08-07 --fail-only
```

- 자유문 검색은 **FTS5 가상표** `record_fts(natural_key, party, body)` 로 받는다.
  `payload` 안의 본문을 여기에 복제하며, 갱신은 트리거가 맡는다.
- 서버 `/api/find` 가 같은 함수를 부른다 — CLI 와 앱이 **같은 질의 한 곳**을 쓴다
  (두 벌로 만들면 결과가 갈리고, 갈린 것을 알아채는 데 또 며칠이 든다).

## 플로우차트 연계

```
python ecount/datalake.py --flow record:1234 --format mermaid
```

`link` 를 너비우선으로 따라가 mermaid `graph LR` 을 뱉는다. 그대로
앱 화면·문서·Artifact 에 붙는다(별도 라이브러리 필요 없음). 노드 색은 `record.status`,
끊긴 자리는 점선 — **어디서 멈췄는지가 그림에서 바로 보이는 것**이 목적이다.

---

## 흡수 경로 — 기존 도구를 갈아엎지 않는다

새 파이프라인을 따로 만들면 둘이 어긋난다. 그래서 **기존 도구 끝에 한 줄만 건다.**

| 기존 도구 | 거는 곳 |
|---|---|
| `upload_intake.py --apply` | 분류 확정 직후 `datalake.ingest_asset()` |
| `download_intake.py --apply` | 이동 완료 직후 같은 함수 |
| `band/convert_dump.py` | 캐시 반영 직후 `ingest_records(kind='band_post')` |
| `erp_grab` 계열 | Excel 저장 직후 `ingest_asset()` |
| `ledger_db.py --apply` | 회차 끝에 `event(area='ledger')` |
| 모든 도구의 실패 경로 | `event(ok=0)` |

`source_index.py` 는 **바로 없애지 않는다.** datalake 가 같은 답을 낸다는 것을
한동안 나란히 돌려 확인한 뒤에 물러난다(그 사이 `--verify-against-json` 이 둘을 대조).

## 증분 — 2시간을 몇 초로

1. `os.scandir` 로 (path, mtime, size) 만 훑는다.
2. DB 의 같은 값과 비교해 **다른 것만** 골라낸다.
3. 고른 것만 sha1·내용 파싱. 안 바뀐 파일은 열지 않는다.
4. 이번 주사에서 못 본 경로는 `gone_at` 만 찍는다(지우지 않는다).

## 지켜야 할 것

- **엑셀은 이 설계에 등장하지 않는다.** CLAUDE.md 의 최종 방향(엑셀 저장 안 함)과 같은 방향이다.
  엑셀 반영은 지금처럼 11:00·15:00 큐 경로만 쓴다 — datalake 는 엑셀을 열지 않는다.
- **삭제하지 않는다.** `gone_at`·`rev`·`event` 로 남긴다. 오늘 밴드 사고가 가르쳐 준 것:
  실패는 삭제의 증거가 아니다.
- **`who` 는 세션 식별자까지 적는다**(`claude:<sid>`). 창이 여러 개인 것이 기본이다.
- 수집 세션/코딩 세션 분리(2026-08-07 지시)는 그대로다 — datalake 는 **읽기만 하는 세션도
  안전하게** 쓸 수 있어야 한다(WAL 이 그것을 준다).

## 만드는 순서 (각 단계가 끝나면 그 자체로 쓸모가 있어야 한다)

1. `datalake.py` 스키마 + `ingest_asset()` + 증분 주사 → **`source_index` 대체 후보**
2. `--find` / FTS5 → 사람이 바로 쓸 수 있는 검색
3. `event()` 를 기존 도구 실패 경로에 배선 → 사고 추적이 한 곳에서
4. `record` / `record_rev` — 밴드 글부터(가장 잘 정돈돼 있다), 다음 ERP
5. `link` + `--flow mermaid` → 앱 연계
6. `synthetic_check.py` 검증 항목 추가(다음 번호를 받는다): 스키마 · append-only 트리거 ·
   증분 정확성 · 왕복 무손실 · 워크트리에서 본체 DB 를 집는가

각 단계는 **합성데이터로 먼저** 검증한다(절대규칙). 실 Z: 주사는 그다음이다.
