# 휴대폰에서 Claude Code 이어서 하기

지금 상태 확인은 언제나 이 한 줄:

```bash
python ecount/remote_ready.py
```

길이 두 가지다. **하는 일이 다르니 둘 다 열어 두는 게 좋다.**

| | A. 웹 Claude Code | B. SSH 로 이 PC 접속 |
|---|---|---|
| 폰에 필요한 것 | 브라우저만 | SSH 앱(Termius·Blink 등) |
| PC 가 꺼져 있어도 | **된다** | 안 된다 |
| 볼 수 있는 것 | GitHub 에 **푸시된** 코드 | 이 PC 전부 |
| Z: 관리대장·이카운트·밴드 | **안 된다** | 된다 |
| 쓸 곳 | 코드·문서·검증·설계 | 실제 원장 작업 |

---

## A. 웹 Claude Code (claude.ai/code)

폰 브라우저로 `claude.ai/code` → 저장소 `Mulder4780/coupang-ecount-reconcile` 선택.
클라우드 샌드박스가 저장소를 복제해 연다.

**전제 하나뿐**: 하려는 작업이 **푸시돼 있어야** 한다. 미푸시 커밋이나 미커밋 변경은
폰에서 아예 보이지 않는다 — `remote_ready.py` 의 `저장소 동기` 항목이 이걸 본다.

여기서 열면 저장소 루트가 곧 `ecount/` 이므로 `CLAUDE.md`(정본)와 `AGENTS.md` 가
바로 로드된다. 시작 체크리스트는 그대로 따르되 **다음은 클라우드에서 실패하는 게 정상**이다:

- `session_handoff.py --check` — 관리대장(Z:)이 없어 버전 탐지가 안 된다
- `data_status.py` · `ledger_db.py --apply` · 밴드/ERP 관련 전부
- `tests/synthetic_check.py` 는 **대부분 돈다**(합성 데이터라서) — 코드 작업의 안전망은 유지된다

즉 폰에서 웹으로 하기 좋은 일: **코드 수정·리팩터·문서·합성검증·설계 검토**.
원장에 쓰는 일은 PC 로 돌아와서 한다(그게 원본 단일성 규칙이기도 하다).

폰에서 고친 것은 PC 에서 `git pull --rebase` 로 받는다.

### A-2. 폰에서 Codex 로 같은 일 하기 (2026-08-04)

Claude 크레딧이 없거나 Codex 로 교대할 때 — 방식은 A 와 완전히 같다.

1. 폰 브라우저(또는 ChatGPT 앱) → `chatgpt.com/codex`
2. GitHub 연결에서 `Mulder4780/coupang-ecount-reconcile` 선택 → 클라우드 샌드박스가 저장소를 복제
3. 저장소 루트에 `AGENTS.md`(Codex 자동 로드)가 있어 규칙·체크리스트가 그대로 적용된다
4. 할 수 있는 일·안 되는 일도 A 와 동일: 코드·문서·합성검증은 되고, Z:·이카운트·밴드는 안 된다
5. 결과는 PR 또는 커밋으로 남는다 → PC 복귀 후 `git pull --rebase`

**전제도 같다: 푸시돼 있어야 보인다.** 자리 뜨기 전 푸시가 폰 작업의 전부다.

### A-3. PC 를 안 켜도 자동으로 도는 것 (GitHub Actions, 2026-08-04)

`.github/workflows/auto_scout.yml` — 자동화 후보 일일 검색이 **GitHub 무료 서버에서
매일 10:00(KST) 자동 실행**되어 `reports/자동화_후보.md` 를 커밋한다. PC·폰 모두 꺼져
있어도 돈다. 비밀키가 필요 없는 읽기 작업만 이렇게 옮긴다 — Z:·이카운트·밴드·엑셀이
필요한 작업은 옮길 수 없고(자료가 PC/LAN 에만 있고, 웹앱은 LAN 전용·외부 개방 금지
규칙), 옮기지 않는다.

---

## B. SSH 로 이 PC 에 붙어 `claude` 실행

진짜 이 PC 이므로 Z: 관리대장까지 다 된다. 준비는 세 가지고, **셋 다 관리자 권한이
필요하거나 시스템 설정이라 사용자가 직접 실행해야 한다.**

### B-1. PC 가 자지 않게 (★ 지금 여기서 막혀 있다)

전원이 연결돼 있어도 **5시간 뒤 대기모드**로 들어간다. 그때부터 폰이 못 붙는다.

```bash
powercfg /change standby-timeout-ac 0
```

화면 끄기는 그대로 둬도 된다(무해하다). 최대 절전은 이미 사실상 꺼져 있다(30일).
바꾼 뒤 `python ecount/remote_ready.py` 로 `PC 절전 OK` 를 확인한다.

### B-2. OpenSSH 서버 켜기

Windows 는 Tailscale 내장 SSH 서버를 쓸 수 없어 OpenSSH 서버가 필요하다.
**관리자 PowerShell**에서:

```powershell
Add-WindowsCapability -Online -Name OpenSSH.Server~~~~0.0.1.0
Set-Service sshd -StartupType Automatic
Start-Service sshd
```

방화벽은 열지 않는다 — 다음 항목의 사설망으로만 들어온다.

### B-3. 사설망으로만 열기 (이미 준비됨)

Tailscale 이 이미 로그인돼 있고 이 PC 주소는 **100.119.175.113** 이다.
공유기 포트를 열 필요가 없고, 인터넷 전체에 SSH 가 노출되지도 않는다.

폰에도 Tailscale 앱을 설치해 **같은 계정으로 로그인**한다. 그러면 폰에서 이 주소가 보인다.

폰 SSH 앱 설정:

```
호스트  100.119.175.113
사용자  hueng
포트    22
```

키 인증을 권한다(비밀번호보다 안전하다). 폰 SSH 앱에서 키를 만들고 공개키를
이 PC 의 `C:\Users\hueng\.ssh\authorized_keys` 에 넣으면 된다.

### B-4. 붙어서 실행

```bash
cd C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT
"C:\Users\hueng\AppData\Roaming\Claude\claude-code\2.1.219\claude.exe"
```

버전 폴더는 갱신되면 바뀐다 — 정확한 경로는 `remote_ready.py` 가 매번 알려 준다.

세션 시작 절차는 PC 에서와 **완전히 같다**(CLAUDE.md 시작 체크리스트). 폰이라고 건너뛰지 않는다.

---

## 폰 화면이 좁을 때

- 작은 화면에서 긴 출력은 그 자체로 고통이다. 토큰 절약 규칙(`| tail -N`)이 폰에서 특히 유효하다.
- 앱(CSOS)으로 되는 일은 앱으로 하는 게 빠르다 — 조회·입력 예약·첨부는 Claude Code 없이 된다.
  고정 주소: https://mulder4780.github.io/coupang-ecount-reconcile/
- Claude Code 가 필요한 건 **코드를 고치거나 원장 로직을 다룰 때**다.

## 무엇이 PC 없이도 되는지 (참고)

| 하려는 일 | PC 꺼짐 | 방법 |
|---|---|---|
| 현황 조회 | 가능 | CSOS 앱(고정 주소, 암호화 사본) |
| 프로젝트코드 입력 예약 | 가능 | CSOS 앱 → 클라우드 큐 → PC 켜지면 반영 |
| 일반 입력·사진 첨부 | 가능 | CSOS 앱 → 기기 보관 → PC 켜지면 자동 전송 |
| 코드 수정 | 가능 | 웹 Claude Code (A) |
| 관리대장 실제 반영 | **불가** | PC 복귀 후 11:00·15:00 회차 |
