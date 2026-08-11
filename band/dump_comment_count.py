# -*- coding: utf-8 -*-
"""덤프 한 개가 **댓글을 정말 담았는지** 센다 (읽기 전용).

왜 따로 필요한가 ([162] · [169]):
  수집기는 '실패 0'으로 끝나도 댓글을 한 건도 안 담을 수 있다 — 2026-08-09 실측으로
  항목 선택자가 한 칸 아래를 가리켜 250건을 성공으로 긁고도 댓글이 0건이었다.
  진행률(`ok/total`)은 **글을 열었다**는 뜻이지 **댓글을 읽었다**는 뜻이 아니다.
  그래서 회차가 끝나면 캐시에 합치기 전에 덤프 자체를 한 번 세어 본다 —
  여기서 0이 나오면 남은 회차를 도는 것은 시간 낭비다(83분이 통째로 날아간다).

세 갈래로 답한다(뭉치면 [169] 의 '0건'이 된다):
  · 댓글 담김      — comments 에 내용이 있다
  · 확인된 0개     — comments_full 이고 목록이 비었다(진짜 댓글 없는 글)
  · 미확인        — 그 외. 이것이 대부분이면 수집기를 의심한다

쓰는 법: python band/dump_comment_count.py [덤프파일...]
        (인자가 없으면 Downloads 의 dump_*.json 을 최신순으로 본다)
"""
import glob
import io
import json
import os
import sys

# ★ 윈도우 콘솔은 cp949 라 '—' 한 글자에 UnicodeEncodeError 로 죽는다. 실측
#   2026-08-11: 표는 다 찍고 **마지막 판정 한 줄에서** 터졌다 — 세어 놓고 결론을
#   못 말하는 계기다([169]). 판정이 제일 중요한 줄이므로 여기서 먼저 막는다.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def count(path):
    with io.open(path, "r", encoding="utf-8") as f:
        d = json.load(f)
    posts = d.get("posts") or {}
    if isinstance(posts, list):
        posts = {str(i): p for i, p in enumerate(posts)}
    withc = full = unknown = ncomments = 0
    for p in posts.values():
        if not isinstance(p, dict):
            continue
        c = p.get("comments")
        if c:
            withc += 1
            ncomments += len(c)
        elif p.get("comments_full"):
            full += 1
        else:
            unknown += 1
    return {
        "글": len(posts), "댓글담김": withc, "댓글수": ncomments,
        "확인된0개": full, "미확인": unknown,
        "없는글": len(d.get("missing") or []),
        "실패": len(d.get("failed") or []),
        "시각없음": len(d.get("notime") or []),
    }


def main(argv):
    files = argv[1:]
    if not files:
        dl = os.path.join(os.path.expanduser("~"), "Downloads")
        files = sorted(glob.glob(os.path.join(dl, "dump_*.json")),
                       key=os.path.getmtime, reverse=True)
    if not files:
        print("볼 덤프가 없습니다 (인자를 주거나 Downloads 를 확인하세요)")
        return 1
    tot = {}
    for p in files:
        r = count(p)
        for k, v in r.items():
            tot[k] = tot.get(k, 0) + v
        print("%-46s %s" % (os.path.basename(p),
                            " · ".join("%s %d" % (k, r[k]) for k in
                                       ("글", "댓글담김", "댓글수", "확인된0개",
                                        "미확인", "없는글", "실패"))))
    if len(files) > 1:
        print("합계  " + " · ".join("%s %d" % (k, tot[k]) for k in
                                   ("글", "댓글담김", "댓글수", "확인된0개",
                                    "미확인", "없는글", "실패")))
    # ★ 판정을 사람에게 떠넘기지 않는다 — 0 이 '없다'인지 '못 읽었다'인지 말한다.
    if tot.get("댓글담김", 0) == 0 and tot.get("확인된0개", 0) == 0:
        print("\n⚠ 댓글이 한 건도 안 담겼고 '확인된 0개'도 없습니다 —"
              " 수집기가 못 읽은 쪽을 먼저 의심하십시오([162]). 남은 회차를 그냥"
              " 돌리면 시간만 씁니다.")
    elif tot.get("댓글담김", 0) == 0:
        print("\n댓글 담긴 글은 0건이지만 '확인된 0개'가 %d건입니다 —"
              " 입력창까지 그려진 뒤 목록이 비었다는 뜻이라 **진짜 댓글이 없는**"
              " 글로 봅니다([199])." % tot.get("확인된0개", 0))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
