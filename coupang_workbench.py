# -*- coding: utf-8 -*-
"""
coupang_workbench.py — Coupang Service Operations System — Workbench (PC 데스크톱 앱)
================================================================
관리대장·에이전트·이카운트·밴드·카톡 대조를 한 화면에서 다루는 통합 프로그램.
표준 라이브러리(Tkinter)만 사용 — 어떤 PC든 파이썬만 있으면 실행.

역할 분담:
  - 앱을 안 쓸 때: 09:50 스케줄러 에이전트가 알아서 대조·자동입력·리포트 (이 앱 없이도 동작)
  - 앱을 쓸 때: 즉시 실행·상태 확인·대장 열기·데이터 투입·전표 전송(사람 승인)을 클릭으로

실행:  python coupang_workbench.py            # GUI
       python coupang_workbench.py --status   # 콘솔 상태 요약(에이전트·테스트용)
"""
import sys, os, re, glob, json, threading, queue, subprocess
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
ENV = {**os.environ, "PYTHONIOENCODING": "utf-8"}
CONFIG_PATH = os.path.join(ROOT, "config", "ecount_config.json")


# ───────────────────────── 상태 수집 ─────────────────────────
def get_status():
    st = {"time": datetime.now().strftime("%H:%M:%S")}
    try:
        cfg = json.load(open(CONFIG_PATH, encoding="utf-8"))
        folder = os.path.dirname(cfg["reconcile"]["master_xlsx"])
        vers = []
        for p in glob.glob(os.path.join(folder, "쿠팡_통합업무_일일보고_관리대장_v*.xlsx")):
            m = re.search(r"_v(\d+)\.xlsx$", p)
            if m and "~$" not in p:
                vers.append((int(m.group(1)), p, os.path.getmtime(p)))
        vers.sort()
        if vers:
            v, p, mt = vers[-1]
            st["master"] = p
            st["master_label"] = f"v{v}  (수정 {datetime.fromtimestamp(mt):%m-%d %H:%M})"
            # 포크 감지: 구버전이 최신본보다 나중에 수정됨 = 잘못된 파일에 입력 중
            forks = [f"v{ov}" for ov, op, omt in vers[:-1] if omt > mt + 60]
            st["fork"] = forks
        else:
            st["master_label"] = "관리대장 없음!"
    except Exception as e:
        st["master_label"] = f"설정 오류: {e}"

    rep = sorted(glob.glob(os.path.join(ROOT, "reports", "종합리포트_*.md")))
    st["last_report"] = rep[-1] if rep else None
    st["report_summary"] = []
    if rep:
        try:
            for ln in open(rep[-1], encoding="utf-8").read().splitlines():
                m = re.match(r"\|\s*([^|]+?)\s*\|\s*(✅|⏭ 스킵|❌ 실패)\s*\|", ln)
                if m:
                    st["report_summary"].append(f"{m.group(2)} {m.group(1)}")
            st["report_time"] = re.search(r"_(\d{8}_\d{4})", rep[-1]).group(1)
        except Exception:
            pass

    try:
        q = json.load(open(os.path.join(ROOT, "updates", "pending_updates.json"), encoding="utf-8"))
        st["pending_updates"] = len(q)
    except Exception:
        st["pending_updates"] = 0
    st["inbox"] = len([f for f in glob.glob(os.path.join(ROOT, "inbox", "*.xlsx"))])
    st["kakao"] = len(glob.glob(os.path.join(ROOT, "kakao", "inbox", "*.txt")))
    st["band_auth"] = os.path.exists(os.path.join(ROOT, "band", ".band_token.json"))
    return st


def status_text(st):
    lines = [f"관리대장 최신: {st.get('master_label','?')}"]
    if st.get("fork"):
        lines.append(f"★ 경고: 구버전 {','.join(st['fork'])} 이 최신본보다 나중에 수정됨 — 입력 파일 확인 필요!")
    lines.append(f"마지막 에이전트 실행: {st.get('report_time','기록 없음')}")
    lines += ["  " + s for s in st.get("report_summary", [])]
    lines.append(f"자동입력 대기: {st['pending_updates']}건 / ERP inbox: {st['inbox']}개 / 카톡 inbox: {st['kakao']}개 / 밴드: {'인증됨' if st['band_auth'] else '심사 대기'}")
    return "\n".join(lines)


# ───────────────────────── GUI ─────────────────────────
def gui():
    import tkinter as tk
    from tkinter import messagebox, scrolledtext

    app = tk.Tk()
    app.title("Coupang Service Operations System — Workbench")
    app.geometry("900x640")
    logq = queue.Queue()
    running = {"flag": False}

    top = tk.LabelFrame(app, text=" 현재 상태 ", padx=8, pady=6)
    top.pack(fill="x", padx=10, pady=6)
    lbl = tk.Label(top, text="로딩...", justify="left", anchor="w", font=("Malgun Gothic", 10))
    lbl.pack(fill="x")

    logbox = scrolledtext.ScrolledText(app, height=18, font=("Consolas", 9))

    def log(msg):
        logq.put(msg)

    def pump():
        try:
            while True:
                logbox.insert("end", logq.get_nowait() + "\n")
                logbox.see("end")
        except queue.Empty:
            pass
        app.after(300, pump)

    def refresh():
        st = get_status()
        txt = status_text(st)
        lbl.config(text=txt, fg="red" if st.get("fork") else "black")
        app.after(30000, refresh)

    def run_bg(title, args, confirm=None):
        if running["flag"]:
            messagebox.showinfo("실행 중", "다른 작업이 끝난 뒤 실행하세요."); return
        if confirm and not messagebox.askyesno("확인", confirm):
            return
        running["flag"] = True
        log(f"\n===== {title} 시작 {datetime.now():%H:%M:%S} =====")

        def work():
            try:
                p = subprocess.Popen([PY] + args, cwd=ROOT, env=ENV,
                                     stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                     text=True, encoding="utf-8", errors="replace")
                for ln in p.stdout:
                    if "UserWarning" not in ln and "warn(msg)" not in ln:
                        log(ln.rstrip())
                p.wait()
                log(f"===== {title} 종료 (코드 {p.returncode}) =====")
            except Exception as e:
                log(f"오류: {e}")
            finally:
                running["flag"] = False
                app.after(100, refresh)
        threading.Thread(target=work, daemon=True).start()

    def open_path(p):
        if p and os.path.exists(p):
            os.startfile(p)
        else:
            messagebox.showinfo("없음", f"경로 없음: {p}")

    btns = tk.Frame(app); btns.pack(fill="x", padx=10)
    grid = [
        ("📂 관리대장 열기(최신)", lambda: open_path(get_status().get("master"))),
        ("▶ 지금 전체 대조 실행", lambda: run_bg("전체 대조", [os.path.join(ROOT, "daily_run.py")])),
        ("🧪 합성검증만", lambda: run_bg("합성검증", [os.path.join(ROOT, "tests", "synthetic_check.py")])),
        ("📄 종합리포트 열기", lambda: open_path(get_status().get("last_report"))),
        ("📥 ERP inbox 폴더", lambda: open_path(os.path.join(ROOT, "inbox"))),
        ("💬 카톡 inbox 폴더", lambda: open_path(os.path.join(ROOT, "kakao", "inbox"))),
        ("✏ 자동입력 미리보기", lambda: run_bg("자동입력 미리보기", [os.path.join(ROOT, "ledger_writer.py")])),
        ("✏ 자동입력 반영(vN+1)", lambda: run_bg("자동입력 반영", [os.path.join(ROOT, "ledger_writer.py"), "--apply"],
                                             confirm="확정 대기 항목을 관리대장 새 버전에 입력합니다.\n(빈 칸만 채움·이전 버전 보존) 진행할까요?")),
        ("🧾 전표 전송대기 확인", lambda: run_bg("전표 dry-run", [os.path.join(ROOT, "ecount_upload.py")])),
        ("🚀 전표 실전송(승인)", lambda: run_bg("전표 실전송", [os.path.join(ROOT, "ecount_upload.py"), "--post"],
                                            confirm="이카운트 ERP에 매출전표를 실제 등록합니다.\n(10.5초/건, 이중계상 가드 적용) 진행할까요?")),
        ("📊 리포트 폴더", lambda: open_path(os.path.join(ROOT, "reports"))),
        ("🔄 상태 새로고침", lambda: refresh()),
    ]
    for i, (name, fn) in enumerate(grid):
        tk.Button(btns, text=name, command=fn, width=24, height=2)\
          .grid(row=i // 4, column=i % 4, padx=4, pady=4, sticky="ew")

    tk.Label(app, text="실행 로그", anchor="w").pack(fill="x", padx=10)
    logbox.pack(fill="both", expand=True, padx=10, pady=(0, 8))
    log("워크벤치 시작 — 앱을 닫아도 09:50 자동 에이전트는 계속 동작합니다.")
    log("류지영 매니저 입력(08:00~09:30) 후 에이전트가 09:50 자동 대조하며, 급하면 [지금 전체 대조 실행]을 누르세요.")
    refresh(); pump()
    app.mainloop()


if __name__ == "__main__":
    if "--status" in sys.argv:
        print(status_text(get_status()))
    else:
        gui()
