# -*- coding: utf-8 -*-
"""ERP 전표가 있는 확인필요 건을 **관리자 지시로** 한 번에 완료 처리한다.

사용자 지시(2026-08-28): *"ERP에 단서가 있는 건들은 전부 완료 처리 해버려"* ·
우려를 말씀드린 뒤 *"go"* · *"전부 진행해"*(재확인).

★★ **ERP 가 완료를 입증한 것이 아니다 — 그 사실을 `basis` 에 그대로 적는다**([169]).
  이 저장소에는 ERP 근거 완료가 **이미 있고 매일 돈다** —
  `settlement_completion.ERP_ISSUED_STATES = ("6.세금계산서발행", "7.수금완료")`.
  여기서 닫는 것은 그 기준 **밖**이다. 실측 2026-08-28 후보 125건의 ERP 단계는
  `3.오더처리` 99 · `4.세금계산서발행대기` 21 · 그 밖 5 — 곧 **ERP 기준으로는
  계산서 발행 전**이다. 그것을 `완료(ERP 확인)` 이라 적으면 **거짓**이 되고,
  나중에 "무엇을 근거로 닫았나"를 물을 사람이 반드시 있다.

★ **자동 회차로 만들지 않는다**([172]). `ERP_ISSUED_STATES` 를 넓히면 앞으로
  새로 생기는 건까지 즉시 닫혀 **아무도 안 챙긴다** — 확인필요 기능 자체가
  없어진다. 이것은 오늘 화면을 보고 내린 **일회성 정리**이지 판정 규칙이 아니다.

★ **되돌릴 수 있다** — `--undo`. 이 도구가 찍은 `status` 를 가진 것만 골라
  `ledger_db.resolution_retract` 로 지운다. 남의 근거는 한 글자도 안 건드린다.

★ **이미 닫힌 건은 안 덮는다** — `resolution_sync` 는 upsert 라, 덮으면
  `완료(ERP 수금확인)` 같은 **진짜 근거를 잃는다**([169]).

⚠ **미리보기가 기본이다**([457] 과 같은 규칙 — 되돌리기 어려운 쪽은 사람이
  명령할 때만 실제로 한다). 실제로 쓰려면 `--apply`.

쓰는 법:
    python erp_bulk_close.py            # 무엇이 닫힐지만 보여 준다
    python erp_bulk_close.py --apply    # 실제로 닫는다
    python erp_bulk_close.py --undo     # 이 도구가 닫은 것만 되돌린다(미리보기)
    python erp_bulk_close.py --undo --apply
"""
import io
import json
import os
import sys
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

# 무인 회차에서 `sys.stdout` 이 None 일 수 있다([235]).
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ★ 이 도구가 찍는 낱말 — **한 곳**이다([162]). `--undo` 가 이것으로 골라 되돌리고
#   검증도 이것을 본다. 손으로 두 번 적으면 되돌리기가 조용히 아무것도 못 지운다.
#   ⚠ `완료(` 로 시작해야 `app_server._issue_truth_rows` 가 목록에서 내린다.
STATUS = "완료(관리자 일괄지시 · ERP 전표 있음)"
REPORT = os.path.join(ROOT, "reports", "ERP단서_일괄완료.json")
CACHE = os.path.join(ROOT, "reports", ".앱캐시_issues.json")
ERP_INDEX = os.path.join(ROOT, "reports", "ERP판매_프로젝트색인.json")

# `settlement_completion` 이 **이미 매일 닫는** 단계. 여기서 또 닫으면 그 도구의
# 정직한 근거를 이 도구의 낱말로 덮어쓴다([169]) — 그러니 건너뛴다.
ALREADY_AUTO = ("6.세금계산서발행", "7.수금완료", "8.")


def _read_json(path):
    try:
        with io.open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception as exc:
        return {"__못읽음": "%s: %s" % (type(exc).__name__, exc)}


def candidates():
    """닫을 후보를 고른다 — `(목록, 못한이유)`.

    ★ **모르면 후보를 만들지 않는다**([169]). 캐시나 ERP 색인을 못 읽으면 빈 목록과
      이유를 돌려준다 — 근거 없이 완료를 찍는 것은 되돌릴 수 없는 쪽이다.
    """
    issues = _read_json(CACHE)
    if "__못읽음" in issues:
        return [], "확인필요 목록을 못 읽었다(%s) — 앱 화면을 한 번 열면 만들어진다" % issues["__못읽음"]
    idx = _read_json(ERP_INDEX)
    if "__못읽음" in idx:
        return [], "ERP 판매 색인을 못 읽었다(%s)" % idx["__못읽음"]
    index = idx.get("index") or {}
    if not index:
        return [], "ERP 판매 색인이 비어 있다 — ERP 자료가 아직 안 들어왔다"

    rows = (issues.get("값") or {}).get("rows") or []
    if not rows:
        return [], "확인필요 목록이 비어 있다 — 닫을 것이 없다"

    try:
        import ledger_db
        done = ledger_db.resolutions()
    except Exception as exc:
        return [], "이미 닫힌 건을 못 읽었다(%s) — 덮어쓰면 진짜 근거를 잃는다" % exc

    out = []
    for row in rows:
        project = str(row.get("프로젝트NO") or "").strip().upper()
        settle_id = str(row.get("ID") or "").strip()
        # `_issue_truth_rows` 가 목록에서 내리는 조건과 **같아야** 한다 — 다르면
        # 닫아 놓고도 화면에 그대로 남아 "고쳤는데 안 고쳐졌다"가 된다([165]).
        if str(row.get("구분") or "") != "정산" or not settle_id.startswith("JS-"):
            continue
        state = str((index.get(project) or {}).get("state") or "").strip()
        if not state:
            continue                      # ERP 에 단서가 없다 — 형님 지시 범위 밖
        if state.startswith(ALREADY_AUTO):
            continue                      # 매일 도는 도구 몫이다([172])
        if settle_id in done:
            continue                      # 이미 닫혔다 — 남의 근거를 안 덮는다
        out.append({
            "settle_id": settle_id,
            "project": project,
            "status": STATUS,
            "erp단계": state,
            "문제유형": str(row.get("문제유형") or ""),
            "캠프명": str(row.get("캠프명") or ""),
            "basis": (
                "관리자 일괄 지시(2026-08-28) — ERP 판매전표 있음(%s). "
                "⚠ ERP 기준으로는 계산서 발행 전이며 **ERP 가 완료를 입증한 것이 아니다**. "
                "되돌리기: python erp_bulk_close.py --undo --apply" % state
            ),
        })
    return out, ""


def _tally(items, key):
    out = {}
    for it in items:
        out[it.get(key) or "(빈칸)"] = out.get(it.get(key) or "(빈칸)", 0) + 1
    return dict(sorted(out.items(), key=lambda kv: -kv[1]))


def do_close(apply_it):
    items, why = candidates()
    if why:
        print("! 닫지 않았다 — " + why)
        return 2
    if not items:
        print("닫을 것이 없다 — ERP 전표가 있는 정산 건은 이미 다 닫혀 있다.")
        return 0

    print("ERP 전표가 있는 정산 건 %d 건" % len(items))
    print("  ERP 단계별: %s" % _tally(items, "erp단계"))
    print("  문제유형별: %s" % _tally(items, "문제유형"))
    print("  ⚠ ERP 는 이 건들을 '끝났다'고 말하지 않는다 — 관리자 지시로 닫는 것이다.")
    for it in items[:5]:
        print("    %-13s %-10s %-11s %s" % (it["settle_id"], it["project"],
                                            it["erp단계"], it["문제유형"]))
    if len(items) > 5:
        print("    … 그 밖 %d 건" % (len(items) - 5))

    if not apply_it:
        print()
        print("미리보기다. 실제로 닫으려면: python erp_bulk_close.py --apply")
        return 0

    import ledger_db
    n = ledger_db.resolution_sync([{k: it[k] for k in ("settle_id", "project", "status", "basis")}
                                   for it in items])
    stamp = datetime.now().isoformat(timespec="seconds")
    try:
        with io.open(REPORT, "w", encoding="utf-8") as f:
            json.dump({"at": stamp, "status": STATUS, "닫음": n, "항목": items},
                      f, ensure_ascii=False, indent=1)
        where = REPORT
    except Exception as exc:
        # ★ 못 적었으면 **못 적었다고 말한다**([169]) — 조용히 넘어가면 무엇을
        #   닫았는지 나중에 물을 데가 없다. 닫기 자체는 이미 됐다.
        where = "(기록 못 남김: %s)" % exc
    print()
    print("닫았다: %d 건 · 기록: %s" % (n, where))
    print("되돌리려면: python erp_bulk_close.py --undo --apply")
    return 0


def do_undo(apply_it):
    """★ **이 도구가 찍은 것만** 되돌린다 — 남의 근거는 한 글자도 안 건드린다."""
    try:
        import ledger_db
        done = ledger_db.resolutions()
    except Exception as exc:
        print("! 되돌리지 않았다 — 완료 표를 못 읽었다(%s)" % exc)
        return 2
    mine = sorted(sid for sid, v in done.items() if str(v.get("status") or "") == STATUS)
    if not mine:
        print("되돌릴 것이 없다 — 이 도구가 닫은 건이 없다.")
        return 0
    print("이 도구가 닫은 건 %d 건" % len(mine))
    print("  " + " · ".join(mine[:8]) + (" …" if len(mine) > 8 else ""))
    if not apply_it:
        print()
        print("미리보기다. 실제로 되돌리려면: python erp_bulk_close.py --undo --apply")
        return 0
    removed = ledger_db.resolution_retract(mine)
    print()
    print("되돌렸다: %d 건" % removed)
    return 0


def main():
    argv = sys.argv[1:]
    known = {"--apply", "--undo", "--help", "-h"}
    unknown = [a for a in argv if a.startswith("-") and a not in known]
    if unknown:
        # ★ 모르는 깃발을 조용히 무시하지 않는다([186]) — 오타 하나로 미리보기가
        #   실제 실행이 되거나 그 반대가 되면 안 된다.
        print("! 모르는 깃발: %s · 쓸 수 있는 것: %s"
              % (" ".join(unknown), " ".join(sorted(known))))
        return 2
    if "--help" in argv or "-h" in argv:
        print(__doc__)
        return 0
    apply_it = "--apply" in argv
    return do_undo(apply_it) if "--undo" in argv else do_close(apply_it)


if __name__ == "__main__":
    sys.exit(main())
