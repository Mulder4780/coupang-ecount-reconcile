# -*- coding: utf-8 -*-
"""[329] 끝에 한 번만 저장해 죽을 때 진도를 통째로 잃는 자리를 전수로 센다.

이 저장소가 이미 네 번 밟았다: [388] 밴드 중간저장 · [381] 행수 캐시 ·
[406] 복구용 보관 색인 · [427] 게시글 보관 예산.
그래서 재는 것이 순서다([67]) — 짐작으로 고치면 멀쩡한 것을 건드린다([172]).

판정: 함수 안에 루프가 있고 · 루프 안에서 **비싼 일**을 하고 ·
      루프 안에 **저장이 하나도 없고** · 루프 **뒤**에 저장이 있다.
"""
import ast, io, os

ROOT = "C:/Users/hueng/Documents/COUPANG_INTEGRATED_WORK_AGENT/ecount"

# 루프 안에서 이것을 하면 그 루프는 오래 걸릴 수 있다
# get/run/call 은 dict.get 같은 흔한 이름이라 뺀다 - 넣으면 거짓이 대부분이
# 되어 아무도 안 본다([170]).  남긴 것은 "정말 Z: 나 바깥을 만지는" 이름뿐이다.
EXPENSIVE = {
    "run_tree", "Popen", "check_output",
    "load_workbook", "urlopen",
    "stat", "getmtime", "getsize", "scandir", "listdir", "walk",
    "glob", "iglob", "copy", "copy2", "copyfile", "move",
    "read_ledger", "load_records",
}
# 이것이 있으면 '저장'이다
SAVE = {"dump", "write", "writelines", "replace", "save", "commit",
        "put_record", "flush", "to_excel", "savefig"}
# ★ 계기 자기시험이 잡은 것([272]): 이미 고친 자리는 `_index_touch()` 처럼
#   **touch** 라는 이름으로 중간 저장을 한다.  안 넣으면 [406] 으로 이미
#   고친 archive_keep.collect 가 후보로 잡혀 **또 고치러 간다**([172]).
# ★ 예산으로 스스로 멈추는 것도 진도가 남는 쪽이다([427]) - 루프 안에
#   예산 검사가 있으면 후보가 아니다.
SAVE_SUFFIX = ("touch", "checkpoint", "_mark", "_note")
BUDGET_MARK = ("budget", "over_budget", "remaining", "deadline")


def calls(node):
    out = []
    for n in ast.walk(node):
        if isinstance(n, ast.Call):
            f = n.func
            if isinstance(f, ast.Attribute):
                out.append(f.attr)
            elif isinstance(f, ast.Name):
                out.append(f.id)
    return out


def loop_lines(node):
    return (getattr(node, "end_lineno", node.lineno) or node.lineno) - node.lineno + 1


rows = []
for dp, dn, fn in os.walk(ROOT):
    dn[:] = [d for d in dn if d not in
             (".git", "__pycache__", "node_modules", ".claude", "inbox", "db")]
    for f in fn:
        if not f.endswith(".py"):
            continue
        if os.sep + "tests" + os.sep in dp + os.sep:
            continue
        p = os.path.join(dp, f)
        rel = os.path.relpath(p, ROOT).replace(os.sep, "/")
        try:
            tree = ast.parse(io.open(p, encoding="utf-8").read())
        except Exception:
            continue
        for fun in ast.walk(tree):
            if not isinstance(fun, (ast.FunctionDef, ast.AsyncFunctionDef)):
                continue
            loops = [n for n in ast.walk(fun)
                     if isinstance(n, (ast.For, ast.While, ast.AsyncFor))]
            if not loops:
                continue
            big = [L for L in loops if loop_lines(L) >= 8]
            if not big:
                continue
            # 함수 전체 저장 호출
            allc = calls(fun)
            if not (set(allc) & SAVE):
                continue
            for L in big:
                inner = set(calls(L))
                if not (inner & EXPENSIVE):
                    continue          # 루프가 비싸지 않다
                if inner & SAVE:
                    continue          # 루프 안에 이미 저장이 있다 = 중간 저장 있음
                if any(c.endswith(SAVE_SUFFIX) for c in inner):
                    continue          # _index_touch() 류 = 중간 저장 있음([406])
                blob = " ".join(inner).lower()
                if any(b in blob for b in BUDGET_MARK):
                    continue          # 예산으로 스스로 멈춘다([427])
                rows.append((rel, fun.name, fun.lineno, L.lineno,
                             loop_lines(L), sorted(inner & EXPENSIVE)[:3]))
                break

rows.sort(key=lambda r: (r[0], r[3]))
print("후보 %d곳 / 파일 %d개" % (len(rows), len({r[0] for r in rows})))
print()
for r in rows:
    print("%-34s %-28s 루프 %d행(L%d)  %s"
          % (r[0], r[1][:28], r[4], r[3], ",".join(r[5])))
