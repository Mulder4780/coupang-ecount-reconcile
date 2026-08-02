# -*- coding: utf-8 -*-
"""
excel_recalc.py — 엑셀을 **에이전트가 알아서 열고 계산하고 닫는다**
===============================================================================
사용자 지시(2026-07-30): "엑셀 파일을 열어야 될 사항이 발생되면 에이전트가 알아서 열고
닫고 하게 알고리즘 정리 (코덱스가 읽어서 넘겨받아 처리할 수 있는 구조)"

## 왜 이게 필요한가
관리대장의 접수ID·정산ID·금액·차액은 **수식**이다. 도구가 값(프로젝트NO·원천업무ID)을
넣어도 엑셀이 한 번 계산하기 전에는 캐시값이 없어 **앱에는 안 보인다.**
그동안 "엑셀을 한 번 열어 주세요" 라고 사람에게 부탁해 왔다. 이제 그 부탁을 없앤다.

## 어떻게 — pywin32 없이 PowerShell COM
이 프로젝트는 표준 라이브러리 + openpyxl 만 쓴다(새 의존성 금지). 그런데 엑셀 수식을
계산할 수 있는 건 엑셀뿐이다. 해법은 **PowerShell 의 COM**:
    New-Object -ComObject Excel.Application
파이썬 패키지를 하나도 늘리지 않고 진짜 엑셀에게 계산을 시킨다(확인: Excel 16.0).

## 안전 규칙 (이 순서가 곧 알고리즘이다)
  1. 대기 건수가 0이면 **아무것도 하지 않는다**(엑셀을 괜히 띄우지 않는다).
  2. 입력 보호시간(류지영 매니저 입력 08:00~09:30)에는 하지 않는다.
  3. `ai_claim` 으로 원장 점유를 잡는다 — Codex와 동시에 만지지 않는다.
  4. 사람이 그 파일을 열어 두었으면(`~$` 잠금 파일) **건드리지 않고 물러난다.**
     남이 편집 중인 파일을 자동화가 저장하면 그 사람 작업이 날아간다.
  5. 복구 경고 원인이 확인됐다는 체크포인트와 **현재 원본의 SHA-256 승인**이 모두 있어야 한다.
     승인 파일은 `reports/excel_recalc_clearance.json`이고, 파일명·해시가 현재 원본과 정확히
     일치해야 한다. 새 버전이 생기면 승인은 자동 만료된다.
  6. 엑셀은 **보이지 않게**, 대화상자 없이(DisplayAlerts=false) 띄운다.
  7. 원본을 덮어쓰지 않고 **vN+1 로 저장**한다 — 프로젝트의 버전 규칙 그대로, 되돌릴 수 있다.
  8. 성공·실패와 무관하게 **엑셀을 반드시 닫는다**(finally). 좀비 EXCEL.EXE 를 남기지 않는다.
  9. 끝나고 대기 건수를 다시 세어 **정말 0이 됐는지 확인**한다. 안 줄었으면 실패로 본다.

## Codex 가 이어받는 법
  · 상태는 전부 파일에 있다: `reports/재계산대기.json`(무엇이 몇 건 대기),
    `reports/excel_recalc.json`(마지막 실행 결과·실패 사유).
  · 진입점은 하나: `python excel_recalc.py --run`. daily_run 2.98 단계에 연결돼 있다.
    다만 복구 경고 차단 항목이 남아 있거나 현재 원본의 해시 승인이 없으면 절대 열지 않는다.
  · 실패해도 원장은 그대로다(vN+1 을 못 만들었을 뿐). 다시 실행하면 된다.
  · 엑셀이 없는 PC(서버·CI)에서는 `available()` 이 False 를 돌려주고 조용히 건너뛴다.

사용
  python excel_recalc.py            # 지금 필요한지 보기(열지 않는다)
  python excel_recalc.py --run      # 필요하면 열고 계산하고 닫는다
  python excel_recalc.py --self-test
"""
import sys, os, re, json, subprocess, hashlib
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

RESULT = os.path.join(ROOT, "reports", "excel_recalc.json")
CLEARANCE = os.path.join(ROOT, "reports", "excel_recalc_clearance.json")
CHECKPOINT = os.path.join(ROOT, "reports", "진행체크포인트.json")
TIMEOUT = 900          # 초. 1.6MB 워크북 재계산은 보통 1분 안쪽이지만 넉넉히 둔다.


def _s(v):
    return "" if v is None else str(v).strip()


# ── 순수 판정 (합성 검증 대상) ─────────────────────────────────
def need_recalc(pending):
    """재계산이 필요한가 — 대기 건수 문서를 받아 판단한다."""
    try:
        return int((pending or {}).get("대기합계") or 0) > 0
    except (TypeError, ValueError):
        return False


def lock_file(path):
    """엑셀이 열어 둔 파일에는 같은 폴더에 `~$이름.xlsx` 잠금 파일이 생긴다."""
    d, name = os.path.dirname(path), os.path.basename(path)
    return os.path.join(d, "~$" + name)


def someone_editing(path):
    """사람이 그 파일을 열어 두었나 — 열려 있으면 자동화는 물러난다."""
    return os.path.exists(lock_file(path))


def next_version_path(path):
    """…_vN.xlsx → …_v(N+1).xlsx. 원본을 덮어쓰지 않는다(프로젝트 버전 규칙)."""
    m = re.search(r"^(.*_v)(\d+)(\.xlsx)$", path)
    if not m:
        return None
    return f"{m.group(1)}{int(m.group(2)) + 1}{m.group(3)}"


def _sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def recovery_clearance(master):
    """현재 정본을 Excel로 열어도 되는가.

    `진행체크포인트.json`의 복구 경고 금지가 먼저 해소돼야 하고, 사람이 정상 개방을
    확인한 **그 파일 자체**의 파일명·SHA-256 승인이 있어야 한다. 파일명만 승인하면
    같은 이름으로 바뀐 손상본을 열 수 있으므로 해시까지 묶는다.
    """
    try:
        checkpoint = json.load(open(CHECKPOINT, encoding="utf-8"))
    except Exception:
        return False, "진행 체크포인트를 읽지 못해 안전 승인을 확인할 수 없습니다"
    pending = "\n".join(_s(v) for v in checkpoint.get("남은작업", []))
    if "복구 경고" in pending and "금지" in pending:
        return False, "Excel 복구 경고 원인 확인 전 정본 개방·재계산 금지"
    if not master or not os.path.exists(master):
        return False, "관리대장을 찾지 못했습니다"
    try:
        approval = json.load(open(CLEARANCE, encoding="utf-8"))
    except Exception:
        return False, "현재 관리대장의 정상 개방 승인 파일이 없습니다"
    if approval.get("승인") is not True:
        return False, "현재 관리대장의 정상 개방 승인이 비활성입니다"
    if _s(approval.get("파일")) != os.path.basename(master):
        return False, "정상 개방 승인이 현재 관리대장 버전과 다릅니다"
    approved_hash = _s(approval.get("sha256")).lower()
    if not approved_hash or approved_hash != _sha256(master).lower():
        return False, "정상 개방 승인이 현재 관리대장 내용과 다릅니다"
    return True, "정상 개방 승인 확인"


def decide(pending, master, editing, has_excel, recovery_safe):
    """(할까?, 사유) — 순서가 곧 규칙이다. 하나라도 어긋나면 하지 않는다."""
    if not need_recalc(pending):
        return False, "재계산 대기 없음 — 열 필요가 없습니다"
    if not master:
        return False, "관리대장을 찾지 못했습니다"
    if not recovery_safe:
        return False, "Excel 복구 경고 안전 승인이 없어 열지 않습니다"
    if editing:
        return False, "누군가 관리대장을 열어 두었습니다 — 저장하면 그 작업이 날아갑니다"
    if not has_excel:
        return False, "이 PC에 엑셀이 없습니다 — 사람이 한 번 열어야 합니다"
    if not next_version_path(master):
        return False, f"버전 이름 형식이 아닙니다: {os.path.basename(master)}"
    return True, "재계산 진행"


# ── PowerShell COM ───────────────────────────────────────────
PS_AVAILABLE = r"""
$ErrorActionPreference='Stop'
$x = $null
try { $x = New-Object -ComObject Excel.Application; $v = $x.Version; "OK $v" }
catch { "NO" }
finally {
  if ($x -ne $null) {
    try { $x.Quit() } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) } catch {}
  }
}
"""

# ★ finally 로 반드시 닫는다. 안 닫으면 보이지 않는 EXCEL.EXE 가 파일을 물고 남는다.
PS_RECALC = r"""
$ErrorActionPreference='Stop'
# ★ 경로를 명령줄 인자로 넘기지 않고 **환경변수**로 받는다.
#   -Command 는 $args 를 제대로 바인딩하지 않고(2026-07-30 실패), 게다가 이 프로젝트의
#   경로에는 공백과 ★ 같은 문자가 섞여 있어 따옴표 처리가 쉽게 깨진다.
$src = $env:CSOS_XL_SRC; $dst = $env:CSOS_XL_DST
$x = $null; $wb = $null
try {
  $x = New-Object -ComObject Excel.Application
  $x.Visible = $false            # 화면에 띄우지 않는다(사람 작업을 방해하지 않는다)
  $x.DisplayAlerts = $false      # '연결된 통합문서를 업데이트할까요' 같은 대화상자 차단
  $x.AskToUpdateLinks = $false
  $x.EnableEvents = $false       # 매크로·이벤트가 돌지 않게
  $x.ScreenUpdating = $false
  # UpdateLinks:0 = 외부 링크를 갱신하지 않는다(네트워크 대기·실패로 멈추는 것을 막는다)
  $wb = $x.Workbooks.Open($src, 0, $false)
  $x.CalculateFullRebuild()      # 수식 전체 재구축 — 캐시값이 이때 생긴다
  # SaveAs는 이 워크북에서 "Workbook 클래스 중 SaveAs 속성을 구할 수 없습니다"로
  # 실패했다(2026-08-02 v353). 형식은 이미 xlsx이므로 현재 계산 상태를 그대로
  # 복제하는 SaveCopyAs가 더 안전하고, 열린 원본의 이름·저장 위치도 바꾸지 않는다.
  $wb.SaveCopyAs($dst)
  "OK"
}
catch { "ERR " + $_.Exception.Message }
finally {
  if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
  if ($x  -ne $null) { try { $x.Quit() } catch {} }
  foreach ($o in @($wb,$x)) { if ($o -ne $null) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
"""


def _ps(script, env=None, timeout=120):
    """PowerShell 실행. 경로 같은 값은 env 로 넘긴다(따옴표 문제를 원천 차단)."""
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=timeout,
                           env={**os.environ, **(env or {})})
        lines = [l for l in (r.stdout or "").strip().splitlines() if l.strip()]
        return lines[-1] if lines else (("ERR " + (r.stderr or "").strip()[:120]) if r.stderr else "")
    except subprocess.TimeoutExpired:
        return "ERR 시간초과"
    except Exception as e:
        return f"ERR {type(e).__name__}"


def available():
    return _ps(PS_AVAILABLE, timeout=90).startswith("OK")


def pending_doc():
    try:
        return json.load(open(os.path.join(ROOT, "reports", "재계산대기.json"), encoding="utf-8"))
    except Exception:
        return {}


def refresh_pending():
    """대기 건수를 다시 센다(결과 확인용)."""
    try:
        subprocess.run([sys.executable, os.path.join(ROOT, "recalc_pending.py")], cwd=ROOT,
                       capture_output=True, timeout=600)
    except Exception:
        pass
    return pending_doc()


def save_result(doc):
    os.makedirs(os.path.dirname(RESULT), exist_ok=True)
    json.dump(doc, open(RESULT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def run():
    from operation_window import is_input_window, input_window_label
    if is_input_window():
        msg = f"입력 보호시간({input_window_label()}) — 엑셀을 열지 않습니다"
        print(msg); save_result({"시각": datetime.now().isoformat(timespec="seconds"),
                                 "상태": "보류", "사유": msg})
        return 0

    from workbook_patch import latest_master
    master, ver = latest_master()
    pend = refresh_pending()
    if not need_recalc(pend):
        why = "재계산 대기 없음 — 열 필요가 없습니다"
        print(why)
        save_result({"시각": datetime.now().isoformat(timespec="seconds"), "상태": "건너뜀",
                     "사유": why, "대기": 0})
        return 0
    safe, safe_why = recovery_clearance(master)
    if not safe:
        print(safe_why)
        save_result({"시각": datetime.now().isoformat(timespec="seconds"), "상태": "차단",
                     "사유": safe_why, "대기": pend.get("대기합계", 0)})
        return 0
    ok, why = decide(pend, master, someone_editing(master), available(), safe)
    if not ok:
        print(why)
        save_result({"시각": datetime.now().isoformat(timespec="seconds"), "상태": "건너뜀",
                     "사유": why, "대기": pend.get("대기합계", 0)})
        return 0

    import claim_guard
    claim_guard.require("ledger")
    before = int(pend.get("대기합계") or 0)
    # 점유를 잡는 사이 다른 작업이 새 버전을 만들었으면 예전에 본 원본으로 계속하면 안 된다.
    latest_now, ver_now = latest_master()
    if os.path.normcase(os.path.abspath(latest_now)) != os.path.normcase(os.path.abspath(master)):
        msg = f"관리대장이 v{ver}→v{ver_now}로 바뀌어 재계산을 중단합니다"
        print(msg)
        save_result({"시각": datetime.now().isoformat(timespec="seconds"), "상태": "보류",
                     "사유": msg, "대기": before})
        return 0

    dst = next_version_path(master)
    if os.path.exists(dst):
        msg = f"{os.path.basename(dst)}가 이미 있어 덮어쓰지 않습니다"
        print(msg)
        save_result({"시각": datetime.now().isoformat(timespec="seconds"), "상태": "보류",
                     "사유": msg, "대기": before})
        return 0
    tmp_dst = dst[:-5] + f".recalc-{os.getpid()}.xlsx"
    print(f"엑셀 재계산 시작 — v{ver} · 대기 {before}건")
    try:
        out = _ps(PS_RECALC, env={"CSOS_XL_SRC": master, "CSOS_XL_DST": tmp_dst}, timeout=TIMEOUT)
        if not out.startswith("OK") or not os.path.exists(tmp_dst):
            msg = f"엑셀 재계산 실패: {out[:160]}"
            print(msg)
            save_result({"시각": datetime.now().isoformat(timespec="seconds"), "상태": "실패",
                         "사유": msg, "대기": before})
            return 1

        # 정식 vN+1 이름을 주기 전에 임시본을 직접 검사한다. 개선되지 않은 파일은 정본 목록에
        # 한 순간도 나타나지 않는다.
        import recalc_pending
        after = sum(x["대기"] for x in recalc_pending.scan(tmp_dst))
        if after >= before:
            msg = f"임시 재계산본 검증 실패: 대기 {before}→{after}건 — 정본 승격 안 함"
            print(msg)
            save_result({"시각": datetime.now().isoformat(timespec="seconds"), "상태": "실패",
                         "사유": msg, "대기_전": before, "대기_후": after})
            return 1

        os.replace(tmp_dst, dst)
        refresh_pending()
        state = "완료" if after == 0 else "일부"
        print(f"v{ver} → v{ver + 1} 생성 · 대기 {before} → {after}건 ({state})")
        save_result({"시각": datetime.now().isoformat(timespec="seconds"), "상태": state,
                     "이전버전": ver, "새버전": ver + 1, "대기_전": before, "대기_후": after,
                     "파일": os.path.basename(dst)})
        return 0
    finally:
        if os.path.exists(tmp_dst):
            try:
                os.remove(tmp_dst)
            except OSError:
                pass


# ── 합성 검증 ─────────────────────────────────────────────────
def self_test():
    bad = 0
    if need_recalc({"대기합계": 3}) is not True or need_recalc({"대기합계": 0}) is not False:
        print("  [FAIL] 대기 판정"); bad += 1
    if need_recalc({}) or need_recalc(None) or need_recalc({"대기합계": "x"}):
        print("  [FAIL] 잘못된 문서를 필요로 판정"); bad += 1
    p = r"Z:\a\쿠팡_통합업무_일일보고_관리대장_v323.xlsx"
    if next_version_path(p) != r"Z:\a\쿠팡_통합업무_일일보고_관리대장_v324.xlsx":
        print("  [FAIL] 다음 버전 경로", next_version_path(p)); bad += 1
    if next_version_path(r"Z:\a\이름없음.xlsx") is not None:
        print("  [FAIL] 버전 없는 파일을 통과시킴"); bad += 1
    if os.path.basename(lock_file(p)) != "~$쿠팡_통합업무_일일보고_관리대장_v323.xlsx":
        print("  [FAIL] 잠금 파일 경로"); bad += 1
    # 순서 규칙: 사람이 열어 두었으면 절대 진행하지 않는다(그 사람 작업이 날아간다)
    cases = [
        ({"대기합계": 0}, p, False, True, False, False, "대기 없음"),
        ({"대기합계": 5}, p, False, True, False, False, "안전 승인"),
        ({"대기합계": 5}, p, True,  True, True, False, "열어 두었"),
        ({"대기합계": 5}, p, False, False, True, False, "엑셀이 없"),
        ({"대기합계": 5}, "", False, True, True, False, "찾지 못"),
        ({"대기합계": 5}, p, False, True, True, True,  "진행"),
    ]
    for pend, master, editing, has_excel, recovery_safe, want, frag in cases:
        got, why = decide(pend, master, editing, has_excel, recovery_safe)
        if got != want or frag not in why:
            print(f"  [FAIL] decide {pend} {editing} {has_excel} → {got}/{why}"); bad += 1
    # PowerShell 은 반드시 닫는 절차를 갖고 있어야 한다(좀비 EXCEL.EXE 방지)
    for token in ("finally", "$wb.Close($false)", "$x.Quit()", "DisplayAlerts = $false",
                  "$x.Visible = $false", "CalculateFullRebuild", "SaveCopyAs"):
        if token not in PS_RECALC:
            print(f"  [FAIL] PS 스크립트에 {token} 없음"); bad += 1
    print("excel_recalc self-test:", "OK" if not bad else f"{bad}건 실패")
    return bad == 0


def main():
    if "--self-test" in sys.argv:
        sys.exit(0 if self_test() else 1)
    if "--run" in sys.argv:
        sys.exit(run())
    from workbook_patch import latest_master
    master, ver = latest_master()
    pend = pending_doc()
    if not need_recalc(pend):
        ok, why = False, "재계산 대기 없음 — 열 필요가 없습니다"
    else:
        safe, safe_why = recovery_clearance(master)
        if not safe:
            ok, why = False, safe_why
        else:
            ok, why = decide(pend, master, someone_editing(master), available(), safe)
    print(f"관리대장 v{ver} · 재계산 대기 {pend.get('대기합계', 0)}건")
    print(("→ " if ok else "· ") + why)
    if ok:
        print("  실행: python excel_recalc.py --run")


if __name__ == "__main__":
    main()
