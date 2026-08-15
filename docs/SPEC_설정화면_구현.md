# 설정 화면 고도화 — **구현 지시서** (2026-08-15)

설계·근거는 `SPEC_설정화면.md`. 이 문서는 **어디를 어떻게 고치는지**만 적는다.
조사는 2026-08-15 21:5x 에 끝났고 **코드 편집은 `code` 점유가 옆 세션에 있어 못 했다.**
점유가 풀리면 이 문서대로 그대로 적용하면 된다 — 다시 조사할 것이 없다.

## 0. 결정된 것 — [103] 과의 충돌은 **절충안**으로 간다

형님 2026-08-15 "고도화 작업 진행해" 로 진행 승인. 고른 길:

> **허브 타일을 '바로가기'가 아니라 '상태판'으로 뜻을 바꾼다.**

- 형님 지시("설정이랑 관계없는 건 넣지 마라")를 지킨다 — 업무 화면으로 가는 길은
  사이드바가 이미 한다. 설정이 메뉴를 복제하지 않는다.
- 분담판 **[103]**(`/api/system-audit` 상태 스트립)은 **살아난다** — 오히려 이
  형태가 [103] 의 본뜻에 더 맞는다. `dff263d` 로 만든 허브 그리드도 안 버린다.
- 즉 **타일 수를 줄이는 것이 아니라 타일이 하는 말을 바꾸는 것**이다.
  "정산 화면으로 감" → "정산: 이상 없음 / 확인 3건".

## 1. 확인된 좌표 (2026-08-15 실측 · webapp/index.html)

| 무엇 | 줄 | 비고 |
|---|---:|---|
| `#v-settings` 마크업 | 4513~4565 | 카드 셋 |
| 허브 타일 16개 | 4541~4562 | `routeNav(` 15 + `toggleTheme(` 1 |
| `renderAdminSettings` | 6895 | `adminHubState` 는 6912 |
| `loadAdminSettings(force)` | 6915 | `/api/auth/session` |
| `adminModeLogout` | 6926 | |
| `openPinDialog` | 6858 | |
| 글꼴 카드(통째) | 3898~3925 | `#v-run` 안 · **표식 포함** |
| `setFontPreset(key)` | 5145 | `csos_font_preset` |
| `THEME_MODES` / `THEME_LABEL` | 5621 / 5622 | `['light','dark','system']` |
| `themeMode()` | 5625 | 지금 값 |
| `applyTheme(t)` `toggleTheme()` | 5639 / 5689 | 토글은 **순환**이다 |
| `installState()` | 6041 | `{ok,code,msg,fix,alt}` |
| `routeNav(v)` | 8590 | 지금 `settings` 갈래 **없음** |
| `openSystemAudit()` | 8601 | 실행 탭 진단 카드로 |

★ **글꼴 카드는 통째로 옮겨도 된다** — `font_switch.py` 는 `FONT-CARDS:BEGIN/END`
**표식으로** 찾는다(149·166줄). 위치를 안 본다. 카드 `<div class="card">` 부터
`</div>` 까지(표식 두 줄 포함) 그대로 잘라 설정으로 옮기고 `#v-run` 에서는 지운다.
옮긴 뒤 `python webapp/font_switch.py --sync` 로 확인한다.

## 2. 고칠 것 — 순서대로

### ① `#v-settings` 마크업 재구성 (4513~4565 교체)

구역 넷. `settings-grid`·`settings-card` 클래스는 **그대로 쓴다**(CSS 4470~4511 유지).

**A. 이 기기** (`wide`) — 서버에 안 간다
- 글꼴: 3898~3925 카드 본문을 여기로 (표식·`id="fontCards"`·`id="fontNow"` 유지)
- 화면 밝기: 라디오 3개 — **`toggleTheme()` 을 부르지 않는다**(그건 순환이라 원하는
  값을 못 고른다). `setThemeMode(m)` 를 새로 만들되 안은 세 줄:
  `localStorage.setItem('cw_theme',m); applyTheme(m); forceRestyle();`
  ★ `applyTheme` 를 다시 쓰지 않는다 — 판정은 한 곳이다(`[162]`)
- 내 화면 구성: **편집은 여기로 안 옮긴다**(카드 순서는 그 화면을 보며 고치는 것이
  자연스럽다). `초기화` 단추 하나만 — `DASH_LAYOUT_KEY` 삭제 후 다시 그림
- 이 기기 설정 초기화: `csos_font_preset`·`cw_theme`·`DASH_LAYOUT_KEY`·
  `ORG_ORDER_KEY`·`HELP_FOLD_KEY`·`SEC_FOLD_KEY` 만 지운다.
  ★ **`cw_pin`·`cw_dev_auth` 는 안 지운다** — 그건 로그인이지 설정이 아니다.
    지우면 "글꼴 초기화"가 로그아웃이 된다
  ★ 확인창(`askYesNo`)을 거치고, 문구에 **"업무 자료는 하나도 안 바뀝니다"** 를 적는다

**B. 계정·권한** (`wide`) — 지금 카드 **그대로 둔다**(4520~4535). 손대지 않는다.

**C. 보안** — 지금 카드 그대로(4536~4541).

**D. 앱** (`wide`, 새 카드)
- 설치 상태: `installState()` 를 **부른다**. `code` 별로 문구가 이미 있으므로
  `.msg` 를 그대로 쓴다. `ok` 면 설치 단추, `fix` 가 있으면 고정 주소 단추
- 접속 주소: `location.origin` 과 `FIXED_APP_ORIGIN` 을 나란히
- 알림 채널: `/api/notify_state` 가 있으면 그것을, 없으면
  **"앱 안 알림만 갑니다 — 밀어 주는 채널이 없습니다"**(`[259]` 그대로)

### ② 허브 → 상태판 (4541~4562 교체)

- `routeNav(` 15개 **전부 제거**. 타일은 `/api/system-audit` 의 `findings` 로 만든다
- 이상 없으면 타일 하나: **"확인할 것이 없습니다"**
- 못 읽으면 **"상태를 못 읽었습니다"** — `이상 없음`이라 하지 않는다(`[169]`)
- 각 타일은 그 항목의 화면으로 가는 단추 **하나**를 갖는다(그건 바로가기가 아니라
  '이 문제를 보러 감'이다)
- `adminHubState` 는 남긴다 — [103] 이 쓰는 자리다

### ③ `routeNav` 에 갈래 추가 (8590)

```js
if(v==='settings'){ loadAdminSettings(); loadSettingsPanel(); }
```
지금은 `settings` 갈래가 **없어서** 설정을 열어도 아무것도 새로 안 읽는다.
(카드 값이 처음 한 번만 채워지고 낡는다)

### ④ 담당자에게도 '이 기기' 구역을 연다

지금 `body.staff-mode` 에서 설정 메뉴가 **통째로** 숨는다(4508) — 그래서
류지영·오종현은 **글꼴을 못 바꾼다.** 축소판을 연다:
- 메뉴는 보이되 **A 구역만** 그린다(B·C·D 는 `staffSlug` 면 안 그림)
- 서버 권한은 한 글자도 안 넓힌다 — 화면 표시만이다(`[279]`)
- ★ 이건 형님 판단이 필요했던 것인데 "진행해"로 승인된 것으로 본다.
  **되돌리기 쉽다** — CSS 한 줄이다

## 3. 검증 `[284]` — `tests/synthetic_check.py`

쓰기 전에 번호가 비었는지 확인: `grep -n "def t284" tests/synthetic_check.py`

1. `#v-settings` 구간에 `routeNav(` 가 **0회**
2. `data-fontkey` 가 `#v-settings` 안에만 있고 `#v-run` 에는 **없다**(사본 금지)
3. `FONT-CARDS:BEGIN`·`END` 표식이 **둘 다 살아 있다**(font_switch 가 못 찾으면
   `--sync` 가 통째로 죽는다)
4. 설정 화면이 `applyTheme`·`installState` 를 **부르기만** 하고 같은 판정을 다시
   적지 않는다 (`prefers-color-scheme` 문자열이 설정 구간에 없을 것)
5. 초기화가 `cw_pin`·`cw_dev_auth` 를 **안 지운다**
6. 설정 구간에 업무 쓰기 호출(`/api/staff/entry`·`/api/ledger`)이 **없다**
7. `routeNav` 에 `settings` 갈래가 있다

## 4. 마친 뒤

```bash
python tests/synthetic_check.py > out.txt 2>&1; echo "exit=$?"
```
★ 파이프에 태우지 말 것 — `| tail` 을 붙이면 종료코드가 `tail` 것이 된다.
**`ALL GREEN` 글자를 눈으로 확인**한다.

```bash
python webapp/restart_server.py --force
```
★ 안 하면 화면이 안 바뀐다(`[156]`). 지금도 서버는 15:51 옛 코드다.
