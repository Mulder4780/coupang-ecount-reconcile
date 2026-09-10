# -*- coding: utf-8 -*-
"""진짜 밴드 접수 글로 잰다. 캐시는 읽기만 한다([247])."""
import os, sys, json, glob, copy
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
R = r"C:\Users\hueng\Documents\COUPANG_INTEGRATED_WORK_AGENT\ecount"
sys.path.insert(0, R)
import band_extract as BE

ok = fail = 0
def jae(name, got, want):
    global ok, fail
    if got == want:
        ok += 1; print("  OK   %-34s %r" % (name, got))
    else:
        fail += 1; print("  FAIL %-34s got=%r want=%r" % (name, got, want))

# ── 진짜 접수 글 하나를 뜬다(짐작한 양식을 쓰지 않는다 · [165])
real = None
for f in glob.glob(os.path.join(BE.CACHE_DIR, "*.json")):
    d = json.load(open(f, encoding="utf-8"))
    for no, p in (d.get("posts") or {}).items():
        c = p.get("content") or ""
        r0 = BE.parse_post(no, p, "84789192") if BE.RE_FORM_HEAD.search(c) else None
        if r0 and r0.get("프로젝트NO") and r0.get("캠프주소") and "신청내용" in c:
            real = (f, no, p); break
    if real: break
if not real:
    sys.exit("접수 글을 못 찾았다 — 재지 않는다([169])")
f, no, p = real
base = BE.parse_post(no, p, "84789192")
print("근거 글: %s / %s  프로젝트NO=%s" % (os.path.basename(f), no, base.get("프로젝트NO")))
print("본문 앞 3줄: %r" % ((p.get("content") or "").split(chr(10))[:3],))

def with_line(txt):
    q = copy.deepcopy(p); q["content"] = (p.get("content") or "") + chr(10) + txt
    return BE.parse_post(no, q, "84789192")

print()
print("(1) 협조요청문 그대로 적었을 때")
jae("캠프긴급도", with_line("● 긴급도 : 긴급").get("캠프긴급도"), "긴급")
jae("보통", with_line("● 긴급도 : 보통").get("캠프긴급도"), "보통")
jae("낮음", with_line("● 긴급도 : 낮음").get("캠프긴급도"), "낮음")

print()
print("(2) 없으면 빈 칸 · 빈 값도 빈 칸([169] 지어내지 않는다)")
jae("그 줄이 없는 글", base.get("캠프긴급도"), "")
jae("값이 빈 줄", with_line("● 긴급도 :").get("캠프긴급도"), "")

print()
print("(3) 양식을 안 고치고 그대로 둔 값도 원문 그대로([166] 낱말을 안 지어낸다)")
raw = "긴급 / 보통 / 낮음 (유니웍스에 반영예정)"
jae("목록 밖 값", with_line("● 긴급도 : " + raw).get("캠프긴급도"), raw)

print()
print("(4) ★ 기존 칸이 한 톨도 안 바뀌었나 — 이 반환값을 읽는 곳이 여럿이다")
after = with_line("● 긴급도 : 긴급")
diff = [k for k in base if base.get(k) != after.get(k)]
jae("캠프긴급도 말고 달라진 칸", [k for k in diff if k not in ("캠프긴급도", "본문", "본문잘림")], [])
jae("캠프긴급도 값", (base.get("캠프긴급도"), after.get("캠프긴급도")), ("", "긴급"))

print()
print("(5) 캠프주소를 잘못 잡지 않나([172] 좁히는 것도 넓히는 것도 고장이다)")
jae("캠프주소 그대로", after.get("캠프주소"), base.get("캠프주소"))

print()
print("결과: OK %d · FAIL %d" % (ok, fail))
sys.exit(1 if fail else 0)
