# -*- coding: utf-8 -*-
"""'캠프가 말한 급함'이 색인 → 행 → 화면까지 가는지 **실행으로** 잰다([295]).

진짜 밴드 캐시·진짜 색인 파일에는 한 글자도 안 쓴다([247]) — 전부 목이고
`finally` 로 되돌린다([371]).
"""
import io, os, sys, json, tempfile, subprocess, shutil
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = r"C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "webapp"))

ok = fail = 0
def jae(name, got, want):
    global ok, fail
    if got == want:
        ok += 1; print("  OK   %-40s %r" % (name, got))
    else:
        fail += 1; print("  FAIL %-40s got=%r want=%r" % (name, got, want))

import app_server as A
import band_extract as BE

REC = [  # 접수 글(이른 날) · 완료 글(늦은 날 · 긴급도 없음)
    {"프로젝트NO": "UJ2601500", "게시일": "2026-09-05", "밴드": "84789192",
     "캠프긴급도": "긴급", "본문": "도어 센서 교체", "진행상태": "", "업무유형": "돌발AS"},
    {"프로젝트NO": "UJ2601500", "게시일": "2026-09-08", "밴드": "84789192",
     "캠프긴급도": "", "본문": "완료했습니다", "진행상태": "작업완료", "업무유형": "돌발AS"},
    {"프로젝트NO": "UJ2601501", "게시일": "2026-09-06", "밴드": "84789192",
     # ★ 등급 낱말이 안 잡히는 글 — 예전 구조(증상에 얹기)면 여기서 사라진다
     "캠프긴급도": "낮음", "본문": "확인 부탁드립니다", "진행상태": "", "업무유형": "돌발AS"},
    {"프로젝트NO": "UJ2601502", "게시일": "2026-09-06", "밴드": "84789192",
     "캠프긴급도": "", "본문": "그냥 글", "진행상태": "", "업무유형": "돌발AS"},
]

tmp = tempfile.mkdtemp(prefix="jae_urg_")
os.makedirs(os.path.join(tmp, "reports"), exist_ok=True)
os.makedirs(os.path.join(tmp, "band", "cache"), exist_ok=True)
_root, _load, _ev = A.ROOT, BE.load_records, dict(A._BAND_EV)
try:
    A.ROOT = tmp                       # 색인 파일이 임시 폴더로 간다([247])
    BE.load_records = lambda *a, **k: list(REC)
    A._BAND_EV["d"], A._BAND_EV["at"] = None, 0
    idx = A._band_completion_index() or {}
    cu = idx.get("캠프긴급도") or {}

    print("\n(1) 색인이 제 칸으로 담나 — 증상과 따로")
    jae("UJ2601500", (cu.get("UJ2601500") or {}).get("값"), "긴급")
    jae("가장 이른 글이 이긴다(때)", (cu.get("UJ2601500") or {}).get("때"), "2026-09-05")
    jae("★ 등급 낱말이 없는 글도 담김", (cu.get("UJ2601501") or {}).get("값"), "낮음")
    jae("긴급도 없는 글은 안 담김", "UJ2601502" in cu, False)
    jae("증상 색인과 별개", "캠프긴급도" in idx and "증상" in idx, True)
finally:
    A.ROOT, BE.load_records = _root, _load
    A._BAND_EV.clear(); A._BAND_EV.update(_ev)
    shutil.rmtree(tmp, ignore_errors=True)

print("\n(2) 행에 파생으로 붙나 — as_grade 와 무관하게")
_bci, _ag = A._band_completion_index, A._as_grade
try:
    A._band_completion_index = lambda: {"캠프긴급도": {
        "UJ2601500": {"값": "긴급", "때": "2026-09-05"},
        "UJ2601501": {"값": "낮음", "때": "2026-09-06"}}}
    A._as_grade = None                 # as_grade 가 죽어도 이 값은 실린다([169])
    rows = {"as": [
        {"프로젝트NO": "UJ2601500", "대응등급": ""},
        {"프로젝트NO": "UJ2601501", "대응등급": "B일반"},   # 이미 고른 행
        {"프로젝트NO": "UJ2609999", "대응등급": ""},        # 색인에 없는 건
    ], "pm": [{"점검ID": "x"}]}
    out = A._add_camp_urgency(rows)
    jae("등급 미정 행", out["as"][0].get("캠프긴급도"), "긴급")
    jae("★ 이미 등급 고른 행에도 붙음", out["as"][1].get("캠프긴급도"), "낮음")
    jae("색인에 없는 건", out["as"][2].get("캠프긴급도"), None)
    jae("pm 은 안 건드림", out["pm"][0], {"점검ID": "x"})
    jae("대응등급을 안 건드림", [r.get("대응등급") for r in out["as"]], ["", "B일반", ""])

    print("\n(3) 색인을 못 읽어도 목록을 안 죽인다([169])")
    def _boom():
        raise RuntimeError("색인 실패")
    A._band_completion_index = _boom
    r2 = {"as": [{"프로젝트NO": "UJ2601500"}]}
    jae("못 읽으면 그대로", A._add_camp_urgency(r2)["as"][0].get("캠프긴급도"), None)
finally:
    A._band_completion_index, A._as_grade = _bci, _ag

print("\n(4) 체인에 실제로 걸렸나([328])")
src = io.open(os.path.join(ROOT, "webapp", "app_server.py"), encoding="utf-8").read()
jae("체인 배선", "_add_camp_urgency(" in src.split("def _add_camp_urgency")[0]
    or "_add_camp_urgency(\n" in src, True)
jae("INPUT_SPEC 에 안 올림(저장 금지)",
    '"name": "캠프긴급도"' in src, False)
aw = io.open(os.path.join(ROOT, "archive_worker.py"), encoding="utf-8").read()
jae("보관본 표에도 안 올림", "캠프긴급도" in aw, False)

print("\n(5) 화면 — node 로 실제로 그려 본다([295])")
h = io.open(os.path.join(ROOT, "webapp", "index.html"), encoding="utf-8").read()
def cut(name):
    i = h.index("function " + name + "(")
    d = 0; j = h.index("{", i)
    for k in range(j, len(h)):
        if h[k] == "{": d += 1
        elif h[k] == "}":
            d -= 1
            if d == 0: return h[i:k + 1]
    raise SystemExit("못 잘랐다: " + name)
js = "\n".join([
    "function esc2(s){return String(s==null?'':s).replace(/[&<>\"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;',\"'\":'&#39;'}[c]));}",
    "function gradeVal(r){return String((r&&r.대응등급)||'').trim();}",
    "function regionTag(r){return '';}",
    cut("campUrgTag"), cut("gradeHint"), cut("gradeTag"), cut("gradeTagCore"),
    "const R=[{프로젝트NO:'a',캠프긴급도:'긴급'},{프로젝트NO:'b'},"
    "{프로젝트NO:'c',대응등급:'A긴급',캠프긴급도:'낮음'},"
    "{프로젝트NO:'d',캠프긴급도:'긴급 / 보통 / 낮음 (유니웍스에 반영예정)'}];",
    "console.log(JSON.stringify(R.map(r=>gradeTag(r))));",
])
f = os.path.join(tempfile.gettempdir(), "jae_urg.js")
io.open(f, "w", encoding="utf-8", newline="").write(js)
p = subprocess.run(["node", f], capture_output=True, text=True,
                   encoding="utf-8", errors="replace")
if p.returncode:
    print("  node 실패:", (p.stderr or "")[:400]); fail += 1
else:
    got = json.loads(p.stdout.strip())
    jae("값 있으면 배지", "캠프 긴급" in got[0], True)
    jae("값 없으면 안 그림", "캠프 " in got[1], False)
    jae("★ 등급 고른 행에도 그림", "캠프 낮음" in got[2] and "A긴급" in got[2], True)
    jae("목록 밖 값도 원문 그대로", "긴급 / 보통 / 낮음" in got[3], True)
    jae("대응등급이 아니라고 적음", "대응등급" in got[0] and "아닙니다" in got[0], True)

print("\n결과: OK %d · FAIL %d" % (ok, fail))
sys.exit(1 if fail else 0)
