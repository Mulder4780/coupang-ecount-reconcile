# -*- coding: utf-8 -*-
"""[524] 무인 통과 하루 한 번 — 실행으로 잰다. 진짜 조율표는 한 글자도 안 건드린다."""
import io, json, os, sys, tempfile
sys.path.insert(0, r"C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount")
import coordinate, collect_gate as CG

tmp = tempfile.mkdtemp(prefix="t524_")
진짜MARK, 진짜MD = coordinate.MARK, coordinate.REPORT_MD
진짜지문 = None
try:
    진짜지문 = os.stat(진짜MARK).st_mtime
except OSError:
    pass
coordinate.MARK = os.path.join(tmp, "작업_조율.json")
coordinate.REPORT_MD = os.path.join(tmp, "작업_조율.md")

def 수집자국():
    try:
        return json.load(io.open(coordinate.MARK, encoding="utf-8")).get("수집", [])
    except Exception:
        return []

ok = []
try:
    # (1) 첫 통과 → 자국 1줄
    assert CG._note_unattended_once() is True, "첫 통과가 안 적혔다"
    r = 수집자국()
    assert len(r) == 1 and r[0]["갈래"] == "완주", "완주 한 줄이 아니다: %r" % r
    assert "무인 회차" in r[0]["왜"], "무인이라 안 적혔다"
    ok.append("1 첫 통과 완주 1줄")

    # (2) 같은 날 두 번째 → 안 적힌다
    for _ in range(5):
        assert CG._note_unattended_once() is False, "같은 날 또 적혔다"
    assert len(수집자국()) == 1, "하루 한 번이 아니다: %d줄" % len(수집자국())
    ok.append("2 같은 날 5번 더 불러도 1줄")

    # (3) 날짜가 바뀌면 다시 적는다
    io.open(CG._unattended_mark(), "w", encoding="utf-8").write("2020-01-01")
    assert CG._note_unattended_once() is True, "날이 바뀌었는데 안 적혔다"
    assert len(수집자국()) == 2, "이틀치가 2줄이 아니다"
    ok.append("3 날 바뀌면 다시 1줄")

    # (4) 마커가 조율표에서 파생된다([402])
    assert os.path.dirname(CG._unattended_mark()) == tmp, "마커가 조율표를 안 따라간다"
    ok.append("4 마커가 조율표 자리를 따라간다")

    # (5) 사람 창 갈래는 한 글자도 안 바뀐다([172]) — guard 가 무인만 이 길로 간다
    import inspect
    src = inspect.getsource(CG.guard)
    assert "_note_unattended_once()" in src, "guard 가 안 부른다([328])"
    i_un = src.index("_note_unattended_once()")
    i_check = src.index("v = check()")
    assert i_un < i_check, "무인 갈래보다 뒤에 있다 — 사람 창까지 적힌다"
    ok.append("5 무인 갈래에서만 부른다")

    # (6) 계기 자기시험([272]) — 하루 한 번 문을 없애면 (2)가 잡히나
    원래 = CG._unattended_mark
    CG._unattended_mark = lambda: os.path.join(tmp, "없는폴더_zz", "x.txt")
    n0 = len(수집자국())
    CG._note_unattended_once(); CG._note_unattended_once()
    잡힘 = len(수집자국()) > n0 + 1
    CG._unattended_mark = 원래
    assert 잡힘, "마커가 죽으면 매번 적혀야 한다(그래야 (2)가 진짜를 잰다)"
    ok.append("6 계기 자기시험 — 마커가 죽으면 매번 적힌다")
finally:
    coordinate.MARK, coordinate.REPORT_MD = 진짜MARK, 진짜MD   # [371]

# 진짜 조율표가 한 글자도 안 바뀌었나([247])
지금 = None
try:
    지금 = os.stat(진짜MARK).st_mtime
except OSError:
    pass
assert 지금 == 진짜지문, "진짜 조율표를 건드렸다([247])"
ok.append("7 진짜 조율표 그대로")
for x in ok:
    print("  OK", x)
print("[524] %d/7" % len(ok))
