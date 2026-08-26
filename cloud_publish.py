# -*- coding: utf-8 -*-
"""
cloud_publish.py — PC가 꺼져 있어도 폰·태블릿이 쓰는 사본을 고정 주소에 올린다
================================================================================
지금까지의 폰 접속은 전부 **사무실 PC가 살아 있어야** 했다(터널·로컬 서버). 이동이 잦아
PC를 못 켜 두면 폰에는 크롬 기본 오류만 떴다. 그래서 데이터를 고정 주소에 얹는다.

  항상              : 고정 주소 → **PC 독립 앱**(조회·프로젝트코드 자동채움·입력 예약)
  PC 켜져 있을 때   : 독립 앱 안의 선택 버튼으로 실시간 앱 연결(입력·대조·엑셀 반영)
  PC 꺼져 있을 때   : 잠긴 최신 사본으로 계속 사용
                    예약은 외부 영구 큐가 즉시 보관하고 PC가 켜지면 자동 합류한다.

고정 주소 저장소는 공개다. 그래서 데이터는 **잠가서** 올린다(csos_crypto).
폰에서 PIN을 넣으면 브라우저가 그 자리에서 푼다 — 서버도 계정도 필요 없다.

  python cloud_publish.py            # 사본 생성(로컬 확인)
  python cloud_publish.py --cloud    # 암호문을 외부 최신 사본으로 직접 갱신
  python cloud_publish.py --push     # 생성 + git commit·push (고정 주소에 실제 반영)
"""
import sys, os, re, json, zlib, subprocess, hashlib, urllib.error, urllib.request
from datetime import datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "webapp"))

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

DOCS = os.path.join(ROOT, "docs")
OUT = os.path.join(DOCS, "data.enc")
WEBCFG = os.path.join(ROOT, "config", "webapp.json")
QUEUECFG = os.path.join(ROOT, "config", "cloud_queue.local.json")
CLOUD_REPORT = os.path.join(ROOT, "reports", "cloud_continuity.json")
PUBLISH_FILES = (
    # 공유 캘린더(2026-08-06). 단톡방에 뿌린 링크가 늘 최신을 보게 매 회차 같이 올린다 —
    # 안 올리면 사람들은 며칠 전 일정을 보면서 그것이 오늘 것인 줄 안다.
    # cal-manifest.json 은 **캘린더 전용 설치 정보**다. 이것이 빠지면 링크를 받은 사람이
    # [설치]를 눌렀을 때 캘린더가 아니라 PIN 을 묻는 업무 앱이 홈 화면에 깔린다.
    "docs/cal.enc", "docs/cal.html", "docs/cal-manifest.json",
    "docs/data.enc", "docs/app.html", "docs/index.html", "docs/manifest.json", "docs/sw.js",
    "docs/icon.svg", "docs/icon-32.png", "docs/icon-180.png",
    "docs/icon-192.png", "docs/icon-512.png",
    "docs/icons/clipboard-copy.svg", "docs/icons/file-spreadsheet.svg",
    "docs/icons/image-down.svg", "docs/icons/printer.svg",
    "docs/icons/refresh-cw.svg", "docs/icons/LICENSE-lucide.txt",
    "docs/icons/bootstrap-tools.svg", "docs/icons/bootstrap-box-seam.svg",
    "docs/icons/bootstrap-calculator-fill.svg", "docs/icons/LICENSE-bootstrap-icons.txt",
)


def snapshot_key():
    """폰 사본 키와 공개 가능한 PIN 유도 메타데이터를 돌려준다.

    실시간 앱은 PIN 평문을 보관하지 않고 PBKDF2 해시만 보관한다. 게시기는 그
    해시를 2차 암호화 키로 쓰고, 폰은 사용자가 입력한 PIN에서 같은 해시를 유도한다.
    """
    try:
        cfg = json.load(open(WEBCFG, encoding="utf-8"))
    except Exception:
        sys.exit("config/webapp.json 인증 설정을 읽을 수 없습니다.")

    legacy = str(cfg.get("pin") or "").strip()
    if re.fullmatch(r"\d{4}", legacy):
        return legacy, None

    admin = ((cfg.get("auth") or {}).get("admin") or {})
    digest = str(admin.get("hash") or "").strip().lower()
    salt = str(admin.get("salt") or "").strip().lower()
    iterations = int(admin.get("iterations") or 0)
    if not (re.fullmatch(r"[0-9a-f]{64}", digest)
            and re.fullmatch(r"[0-9a-f]{32}", salt)
            and iterations >= 100_000):
        sys.exit("config/webapp.json 관리자 인증 해시가 없어 폰 사본을 암호화할 수 없습니다.")
    return digest, {
        "mode": "auth-pbkdf2-sha256",
        "salt_hex": salt,
        "iterations": iterations,
    }


def payload():
    """폰이 필요로 하는 것만 담는다. 인증키·엑셀 원본·밴드 원문은 담지 않는다."""
    import mobile_snapshot as M
    import project_resolve as P
    import app_server as A

    d = M.collect()                       # 확인필요·업무·정산·계산서 (앱과 같은 소스)
    ev = P.evidence()

    codes = {}
    for c in sorted(set(ev["band"]) | set(ev["ledger"]) | set(ev["book"])):
        if not re.fullmatch(r"UJ26\d{5}", str(c or ""), re.I):
            continue
        r = P.resolve(c, ev)
        if not A.app_project_result(c, r):
            continue
        rec = {k: r.get(k) for k in ("camp", "kind", "cost", "tech", "date", "status",
                                     "sheet", "row", "state") if r.get(k)}
        if r.get("ids"):
            rec["ids"] = r["ids"]
        if r.get("unknown"):
            rec["miss"] = r["unknown"]
        if not r.get("ok"):
            rec["why"] = r.get("reason", "")
        codes[c] = rec

    # 기사별 방문 — '이 캠프 누가 갔었지'를 현장에서 바로 확인한다
    try:
        import tech_report as TR
        visits, pending, unknown, _m = TR.collect()
        visits = A.app_year_rows(visits, "visit")
        pending = A.app_year_rows(pending, "visit_pending")
        unknown = A.app_year_rows(unknown, "visit")
        by = TR.summary(visits)
        d["tech"] = [{"기사": t, "총": v["총"], "돌발AS": v["돌발AS"],
                      "정기점검": v["정기점검"], "캠프수": len(v["캠프"]), "최근": v["최근"]}
                     for t, v in sorted(by.items(), key=lambda x: -x[1]["총"])]
        # 전부 담으면 사본이 커진다 — 최근 것부터 필요한 만큼만
        d["visits"] = [{k: v[k] for k in ("기사", "방문일", "업무", "캠프명", "프로젝트NO")}
                       for v in sorted(visits, key=lambda x: x["방문일"], reverse=True)[:1200]]
        d["tech_gap"] = {"미방문": len(pending), "기사미기입": len(unknown)}
    except Exception as e:
        print("  ! 기사 정리 건너뜀:", e)

    # 쿠팡 캘린더 — 단톡방에 주소를 뿌려도 열리게 (2026-08-06 지시).
    #   ★ 사설망(Tailscale) 주소는 각자 폰에 VPN 을 깔아야 열린다. 현장 기사들에게
    #     그것을 시킬 수는 없다. 고정 주소는 누구나 열리고, 자료는 여기서 **잠가서**
    #     올리므로 PIN 을 아는 사람만 본다(공개 HTML 에는 업무 내용이 남지 않는다).
    #   담는 것은 앱과 **같은 함수**의 결과다 — 따로 만들면 화면마다 숫자가 달라진다.
    try:
        _cal = A.get_calendar()
        d["calendar"] = {
            "갱신": _cal.get("갱신") or "",
            "분류목록": _cal.get("분류목록") or [],
            # 사본이 너무 커지지 않게 **올해 것만**. 지난해는 PC 앱에서 본다.
            "일정": [{k: e.get(k) for k in
                    ("날짜", "시간", "분류", "캠프명", "장소", "제목",
                     "프로젝트NO", "담당기사", "예측") if e.get(k) not in (None, "")}
                   for e in (_cal.get("일정") or [])
                   if str(e.get("날짜") or "").startswith(str(datetime.now().year))],
        }
    except Exception as e:
        print("  ! 캘린더 건너뜀:", e)

    # 대표 브리핑 — 폰에서 그대로 읽어 드릴 수 있게 담는다
    try:
        import daily_brief as DB
        _b = DB.brief(data=DB.load()[0])
        d["brief"] = {
            "기준일": _b["기준일"],
            "text": DB.text(_b),
            "돌발AS": _b["돌발AS"],
            "정기점검": _b["정기점검"],
            # 고정 주소의 PC 독립 앱에서도 숫자만 보이지 않고 원천 건을 열어야 한다.
            # 이 필드들은 data.enc 안에만 들어가며 공개 HTML에는 업무 내용이 남지 않는다.
            "당일처리목록": _b.get("당일처리목록", []),
            "대표전달목록": _b.get("대표전달목록", []),
            "대표지시목록": _b.get("대표지시목록", []),
            "대표대화판정": _b.get("대표대화판정", {}),
            "현장작업목록": _b.get("현장작업목록", []),
            "신규처리완료목록": _b.get("신규처리완료목록", []),
            "재방문예정목록": _b.get("재방문예정목록", []),
            "완료일미기입목록": _b.get("완료일미기입목록", []),
            "내용미기입": _b.get("내용미기입", []),
            "점검예정목록": _b.get("점검예정목록", []),
            "점검실행목록": _b.get("점검실행목록", []),
            "분기점검목록": _b.get("분기점검목록", []),
            "분기원본일정목록": _b.get("분기원본일정목록", []),
            # 대표 보고와 동일하게, 폰 사본에도 현장 일지 기준의 처리/미처리 사유를 넣는다.
            "일지대조": _b.get("일지대조", {}),
        }
    except Exception as e:
        print("  ! 브리핑 건너뜀:", e)

    d["unbilled"] = unbilled()        # 미청구 — 앱이 열릴 때 경과일을 계산해 띄운다
    d["codes"] = codes
    # ★ 앱이 먼저 답하게 하는 꾸러미 — **PC 가 꺼져 있어도** 폰이 답한다.
    #   규칙(파이썬)을 JS 로 옮겨 적지 않는다. 대신 답을 **미리 만들어** 싣는다:
    #   같은 판단을 두 곳에서 하면 언젠가 갈리고, 갈린 뒤에는 어느 쪽이 맞는지 아무도 모른다.
    #   실측 53KB · 만드는 데 0.2초. 꾸러미가 없어도 사본은 올라가야 하므로 조용히 넘어간다.
    try:
        import local_ai
        d["ask"] = local_ai.answer_pack()
    except Exception as e:
        print("  ! 답 꾸러미 건너뜀:", e)
    d["tail"] = ev["tail"]
    d["cap"] = ev["cap"]
    d["app_year"] = "2026"
    # 업무 흐름(종전·개선) 차트 — 정본은 DB(flow_step)이지만 폰 오프라인 사본은
    # /api/flow 에 못 닿아 차트가 비어 보이고 캡처가 실패했다(2026-08-12 실사고).
    # 답 꾸러미처럼 **미리 만들어** 싣는다 — 차트가 몇 장 안 되고 작다.
    try:
        import ledger_db
        d["flow"] = {"charts": ledger_db.flow_charts(),
                     "steps": {c["key"]: ledger_db.flow_steps(c["key"])
                               for c in ledger_db.FLOW_CHARTS},
                     "notes": {c["key"]: ledger_db.flow_notes(c["key"])
                               for c in ledger_db.FLOW_CHARTS}}
    except Exception as e:
        print("  ! 업무 흐름 꾸러미 건너뜀:", e)
    # 휴대폰 입력 전용 토큰은 공개 소스가 아니라 PIN 암호화 사본 안에만 넣는다.
    # PC 작업자 토큰은 이 사본에 절대 포함하지 않는다.
    try:
        qcfg = json.load(open(QUEUECFG, encoding="utf-8"))
        if qcfg.get("api_base_url") and qcfg.get("enqueue_token"):
            d["cloud_queue"] = {
                "url": str(qcfg["api_base_url"]).rstrip("/"),
                "token": str(qcfg["enqueue_token"]),
            }
    except Exception:
        pass
    d["gen"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    return d


def unbilled():
    """미청구 PO를 **잠긴 사본 안에** 담는다 — 금액·PO번호까지 그대로.

    ★ 예전에는 날짜만 담은 공개 파일(docs/aging.json)을 만들어 GitHub Actions가 매일 보고
      기준을 넘으면 메일을 보내게 했다. 사용자가 메일을 원하지 않아 그 경로는 걷어냈다.
      대신 여기 담아 두고 **앱을 열 때 화면 맨 위에 띄운다** — 경과일은 앱이 그 자리에서
      오늘 날짜로 계산하므로, 사본이 며칠 묵어도 '몇 일 지났는지'는 항상 정확하다.
    """
    import csv as _csv, glob as _g
    rep = sorted(_g.glob(os.path.join(ROOT, "reports", "PO대조_*.csv")))
    if not rep:
        return []
    out = []
    with open(rep[-1], encoding="utf-8-sig") as f:
        for r in _csv.DictReader(f):
            if r.get("유형") != "A":            # A = 미청구(계산서 미발행)
                continue
            d = (r.get("발행일") or "")[:10]
            if not re.match(r"2026-\d{2}-\d{2}", d):
                continue
            try:
                amt = int(float(r.get("쿠팡금액") or 0))
            except ValueError:
                amt = 0
            out.append({"PO": r.get("PO번호", ""), "발행일": d, "금액": amt,
                        "내용": (r.get("내용") or "")[:40],
                        "프로젝트NO": r.get("프로젝트NO", "")})
    return sorted(out, key=lambda x: x["발행일"])


def _cloud_config():
    with open(QUEUECFG, encoding="utf-8") as f:
        cfg = json.load(f)
    url = str(cfg.get("api_base_url") or "").rstrip("/")
    token = str(cfg.get("worker_token") or "")
    if not url.startswith("https://") or len(token) < 32:
        raise ValueError("클라우드 연속운영 주소 또는 작업 토큰 설정이 올바르지 않습니다")
    return url, token


def _cloud_request(url, token, path, body=None, timeout=35):
    data = None if body is None else json.dumps(
        body, ensure_ascii=False, separators=(",", ":")
    ).encode("utf-8")
    req = urllib.request.Request(
        url + path,
        data=data,
        method="GET" if body is None else "POST",
        headers={
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json",
            "User-Agent": "CSOS-Continuity-Publisher/2",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def _load_cloud_report():
    try:
        with open(CLOUD_REPORT, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_cloud_report(result):
    os.makedirs(os.path.dirname(CLOUD_REPORT), exist_ok=True)
    previous = _load_cloud_report()
    data = {**previous, **result, "checked_at": datetime.now().isoformat()}
    if result.get("ok"):
        data["last_success_at"] = datetime.now().isoformat()
        data.pop("error", None)
    tmp = CLOUD_REPORT + f".{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, CLOUD_REPORT)


# ── 폰 사본을 **정말 바뀐 때만** 민다 (2026-08-26 형님 지시 "최적으로") ──
#   실측 2026-08-26: `docs/data.enc` 커밋이 **하루 77~97개**인데 커밋 메시지의
#   업무 숫자가 전부 같았다 — **자료가 안 바뀌었는데 밀고 있었다.**
#   ★ 그 파일 자체로는 못 가른다. 암호화 salt·IV 가 매번 무작위라 **내용이 같아도
#     바이트가 다르다.** 그래서 `gen`(만든 때)을 뺀 **업무 내용 지문**으로 가른다.
#   ★ **못 읽으면 민다**([169]) — 모름을 '같다'로 치면 자료가 영영 안 올라가고,
#     그 사고는 폰 화면이 멀쩡해 보여서 아무도 모른다.
PAGES_MARK = os.path.join(os.path.dirname(CLOUD_REPORT), "폰사본_지문.json")


def pages_unchanged(content_sha256, out_path=None, mark_path=None):
    """지난번 민 것과 **업무 내용이 같고** 사본 파일도 그대로 있으면 참."""
    out_path = out_path or OUT
    mark_path = mark_path or PAGES_MARK
    if not content_sha256:
        return False
    try:
        if not os.path.exists(out_path):
            return False            # 사본이 없으면 만들어야 한다
        with open(mark_path, encoding="utf-8") as f:
            prev = json.load(f)
        return str(prev.get("content_sha256") or "") == str(content_sha256)
    except Exception:
        return False                # 못 읽으면 민다


def remember_pages(content_sha256, generated_at, mark_path=None):
    """민 뒤에 지문을 남긴다. **못 남겨도 조용히 넘어간다** — 다음에 한 번 더 밀 뿐이다."""
    mark_path = mark_path or PAGES_MARK
    try:
        os.makedirs(os.path.dirname(mark_path), exist_ok=True)
        tmp = mark_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"content_sha256": content_sha256, "generated_at": generated_at,
                       "at": datetime.now().isoformat(timespec="seconds")},
                      f, ensure_ascii=False)
        os.replace(tmp, mark_path)
    except Exception:
        pass


def upload_cloud_snapshot(sealed, generated_at, content_sha256):
    """암호문만 외부 D1에 보관한다. 같은 업무 내용이면 쓰기를 생략한다."""
    url, token = _cloud_config()
    previous = _load_cloud_report()
    if previous.get("ok") and previous.get("content_sha256") == content_sha256:
        try:
            ping = _cloud_request(url, token, "/api/ping", timeout=15)
            if ping.get("ok") and (ping.get("snapshot") or {}).get("available"):
                result = {
                    "ok": True,
                    "uploaded": False,
                    "unchanged": True,
                    "generated_at": previous.get("generated_at") or generated_at,
                    "content_sha256": content_sha256,
                    "api_base_url": url,
                }
                _save_cloud_report(result)
                return result
        except Exception:
            # 상태 확인이 실패하면 생략하지 않고 실제 업로드를 시도한다.
            pass
    try:
        out = _cloud_request(
            url,
            token,
            "/api/snapshot",
            {"payload": sealed, "generatedAt": generated_at},
        )
    except urllib.error.HTTPError as exc:
        # ★ 404 는 "자료가 없다" 가 아니라 **떠 있는 판에 그 길이 아직 없다** 는 뜻이다.
        #   실측 2026-08-26: /api/ping 은 200(version 1 · snapshot 칸 없음)인데
        #   /api/snapshot 만 404 다. 그런데 **소스에는 있다** —
        #   csos_cloud_queue_site/app/api/snapshot/route.ts (커밋 90fd7cf · 2026-08-14).
        #   곧 배포본이 그보다 낡았고, **재시도로는 영영 안 풀린다.**
        #   예전에는 "HTTP Error 404: Not Found" 만 남아 32회를 재시도하며 사람에게
        #   무엇을 하라는 말이 없었다 — 조치는 갈래마다 다르다([289]).
        if exc.code == 404:
            raise RuntimeError(
                "클라우드 워커에 /api/snapshot 이 없다 — 떠 있는 판이 낡았다. "
                "소스(csos_cloud_queue_site/app/api/snapshot/route.ts)에는 있으니 "
                "재시도가 아니라 워커 재배포가 필요하다(/api/ping 은 200 이다). "
                "GitHub Pages 사본은 그대로 올라간다 — 폰이 못 보는 것은 D1 최신본뿐이다."
            ) from exc
        raise
    if not out.get("ok"):
        raise RuntimeError(str(out.get("error") or "클라우드 사본 거부"))
    result = {
        "ok": True,
        "uploaded": True,
        "unchanged": not bool(out.get("changed", True)),
        "generated_at": generated_at,
        "content_sha256": content_sha256,
        "snapshot_sha256": out.get("sha256") or "",
        "api_base_url": url,
    }
    _save_cloud_report(result)
    return result


def git_publish(message, runner=subprocess.run):
    """게시 파일을 커밋·푸시한다. 어느 단계든 실패하면 성공으로 보지 않는다."""
    commands = (
        ["git", "add", "--", *PUBLISH_FILES],
        ["git", "commit", "-q", "-m", message],
        ["git", "push", "-q"],
    )
    for cmd in commands:
        r = runner(cmd, cwd=ROOT, capture_output=True, text=True,
                   encoding="utf-8", errors="replace")
        detail = ((r.stderr or "") + "\n" + (r.stdout or "")).strip()
        if r.returncode:
            # 변경이 없는 재실행은 오류가 아니다. 이미 커밋된 미푸시분은 계속 push한다.
            if cmd[1] == "commit" and "nothing to commit" in detail.lower():
                continue
            return False, cmd[1], detail[:300]
    return True, "", ""


def _main():
    from operation_window import input_window_label, is_input_window
    # 입력 보호시간(08:00~09:30)에는 원장이 사람 손에 있어 사본을 뜨지 않는다.
    # ★ `--force` 는 **사람이 직접 시킬 때만** 쓴다(2026-08-06 지시로 추가).
    #   스케줄러가 부르는 회차에는 붙이지 않는다 — 붙이면 규칙이 있으나 마나가 된다.
    if is_input_window() and "--force" not in sys.argv:
        print(f"입력 보호시간({input_window_label()}) — 사본 생성·게시 생략")
        print("  지금 꼭 올려야 하면: python cloud_publish.py --push --force")
        return
    if is_input_window():
        print(f"★ 입력 보호시간({input_window_label()})인데 --force 로 진행합니다.")
    # 앱 아이콘 원본의 내용 해시를 기준으로 PNG·PWA 매니페스트·서비스워커와
    # 현재 PC의 Chrome 설치앱/바로가기 ICO를 먼저 동기화한다. 로고가 바뀌어도
    # 사람이 날짜 버전을 손으로 올리지 않아도 다음 게시·앱 재기동 때 자동 갱신된다.
    icon_sync = os.path.join(ROOT, "webapp", "sync_app_icons.ps1")
    if os.name == "nt" and os.path.isfile(icon_sync):
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", icon_sync],
            cwd=ROOT, capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=60,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if result.returncode:
            print("  ! 앱 아이콘 자동동기화 실패:", (result.stderr or "")[-200:])
        elif (result.stdout or "").strip():
            print(" ", (result.stdout or "").strip().splitlines()[-1])
    import csos_crypto as C
    assert C.self_test(), "암호 자체검증 실패 — 올리면 폰에서 못 연다"

    print("사본 만드는 중…")
    d = payload()
    raw = json.dumps(d, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    stable = dict(d)
    stable.pop("gen", None)
    # * 지문에서 **그때그때 달라지는 값**을 뺀다 ([170] 이 밴드 재수집에서 배운 그 자리).
    #   실측 2026-08-26: `gen` 만 빼고 재니 지문이 **매번 달랐다** — `ask.만든때`
    #   ([184] 가 일부러 넣은 값: 폰이 "이 답이 언제 것인지" 보여 준다) 가 남아 있었다.
    #   그래서 업무가 하나도 안 바뀐 회차도 늘 '바뀜'이 되어 하루 77~97번 밀었다.
    #   ★ **페이로드에는 그대로 남긴다** — 빼는 것은 지문에서다([170] 과 같은 규칙).
    #   ★ 넓게 빼지 않는다([172]) — 실측으로 다른 칸은 이것 하나였다.
    #   실측으로 `ask` 는 **통째로** 빼야 했다. 그 꾸러미에는 `만든때` 말고도
    #   "(실패 · ?분째)" 같은 **상대 시각**이 들어 있어 **분마다 달라진다**.
    #   ★ 빼도 진짜 변경을 안 놓친다 — `ask` 는 미리 만든 **답**이고 그 근거는
    #     `codes`·`issues`·`as`·`pm`·`settle`·`erp` 다. 그것들이 바뀌면 지문이
    #     바뀌므로 사본은 그때 새로 나간다.
    #   ⚠ 다만 **답 만드는 규칙만 고쳤을 때**는 지문이 안 움직인다 —
    #     그때는 `python cloud_publish.py --push --force` 로 한 번 민다.
    stable.pop("ask", None)
    content_sha256 = hashlib.sha256(
        json.dumps(stable, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode("utf-8")
    ).hexdigest()
    # LTE로 받는 파일이라 크기가 곧 대기시간이다. 잠그기 **전에** 줄인다
    # (암호문은 무작위에 가까워서 나중에 줄이면 하나도 안 줄어든다).
    # 폰에서는 브라우저 내장 DecompressionStream('deflate')이 그대로 푼다.
    packed = zlib.compress(raw, 9)
    key, pin_auth = snapshot_key()
    sealed = C.seal(packed, key)
    if pin_auth:
        # salt·반복 횟수는 검증 메타데이터일 뿐 비밀이 아니다. 해시와 PIN은 게시하지 않는다.
        sealed["pin_auth"] = pin_auth
    sealed["zip"] = "deflate"

    sealed_bytes = json.dumps(sealed, separators=(",", ":")).encode("utf-8")
    # 10분 클라우드 회차는 git 추적 파일을 건드리지 않는다. GitHub Pages 백업을
    # 갱신하는 기본/--push 회차만 data.enc를 교체해 작업트리가 늘 깨끗하게 남는다.
    # * 업무 내용이 **지난번과 같으면 사본을 새로 쓰지 않는다** (2026-08-26 지시).
    #   안 쓰면 git 이 `nothing to commit` 으로 넘어가 **커밋이 0**이 된다 —
    #   그런데 `git_publish` 는 그대로 부르므로 **미푸시분은 계속 밀린다**.
    #   실측 2026-08-26: 하루 77~97개 커밋이 전부 같은 업무 숫자였다.
    same_as_before = pages_unchanged(content_sha256)
    write_pages_copy = (("--cloud" not in sys.argv or "--push" in sys.argv)
                        and ("--force" in sys.argv or not same_as_before))
    if write_pages_copy:
        os.makedirs(DOCS, exist_ok=True)
        tmp_out = OUT + ".tmp"
        with open(tmp_out, "wb") as f:
            f.write(sealed_bytes)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_out, OUT)
    kb, rkb = len(sealed_bytes) / 1024, len(raw) / 1024

    print(f"  프로젝트코드 {len(d['codes'])} · 확인필요 {len(d['issues'])} · "
          f"AS {len(d['as'])} · 점검 {len(d['pm'])} · 정산 {len(d['settle'])} · 계산서 {len(d['erp'])}")
    print(f"  {rkb:.0f}KB → 줄여서 {len(packed)/1024:.0f}KB → 잠근 뒤 {kb:.0f}KB  ({OUT})")
    print(f"  미청구 {len(d.get('unbilled', []))}건 — 잠긴 사본 안에(앱 상단에 뜹니다)")
    if same_as_before and not write_pages_copy:
        # * 조용히 넘어가지 않는다([169]) — '안 올렸다'와 '올릴 것이 없었다'는 다른 사실이다.
        print("  폰 사본은 그대로 둡니다 — 지난번과 업무 내용이 같습니다(커밋하지 않습니다).")

    cloud_error = ""
    if "--cloud" in sys.argv or "--push" in sys.argv:
        try:
            cloud = upload_cloud_snapshot(sealed, d["gen"], content_sha256)
            action = "업로드" if cloud.get("uploaded") else "변경 없음"
            print(f"  클라우드 연속운영 사본 {action} · 기준 {cloud.get('generated_at', d['gen'])}")
        except Exception as exc:
            cloud_error = f"{type(exc).__name__}: {str(exc)[:240]}"
            _save_cloud_report({
                "ok": False,
                "uploaded": False,
                "generated_at": d["gen"],
                "content_sha256": content_sha256,
                "error": cloud_error,
            })
            print("  ! 클라우드 연속운영 사본 실패:", cloud_error)

    if "--push" not in sys.argv:
        if "--cloud" in sys.argv and cloud_error:
            raise SystemExit(1)
        if "--cloud" in sys.argv:
            print("\n클라우드 사본을 확인했습니다 — PC가 꺼져도 고정 앱에서 읽습니다.")
            return
        print("\n로컬 생성만 했습니다 — 고정 주소 반영: python cloud_publish.py --push")
        return

    # 잠근 파일만 올린다. 원본(raw)은 디스크에도 남기지 않는다.
    ok, stage, detail = git_publish(
        f"폰 사본 갱신 {d['gen']} (프로젝트코드 {len(d['codes'])}·확인필요 {len(d['issues'])})")
    if not ok:
        print(f"  git {stage} 실패:", detail)
        sys.exit(1)
    if write_pages_copy:
        remember_pages(content_sha256, d["gen"])
    if cloud_error:
        # ★ **성공한 반쪽을 실패로 세지 않는다** (2026-08-26 지시).
        #   폰이 읽는 순서는 `D1 최신 → GitHub Pages → 기기 사본`([271])이라
        #   D1 이 없으면 **Pages 로 떨어지는 것이 설계**다. 그런데 예전에는 여기서
        #   `exit 1` 을 해서 자율복구가 **실패로 세고 10분마다 재시도**했고,
        #   재시도마다 관리대장 2MB 를 다시 읽고 git 을 또 밀었다(하루 77~97 커밋).
        #   게다가 그 가짜 실패가 매일 인계 맨 위를 차지해 **진짜 경보를 덮었다**([170]).
        #   ★ 조용히 넘어가지도 않는다([169]) — `reports/cloud_continuity.json` 에
        #     그대로 남고 인계가 그것을 읽어 **알림**으로 올린다([328]).
        print("  GitHub Pages 사본은 반영됐습니다 — 폰은 그것으로 읽습니다.")
        print("  (D1 최신 사본만 실패: 재시도로는 안 풀립니다 — 워커 재배포가 필요합니다.)")
    print("\n고정 주소에 반영했습니다 — PC를 꺼도 폰에서 열립니다.")
    print("  https://mulder4780.github.io/coupang-ecount-reconcile/")


def main():
    """게시 작업은 수동·앱서버·다른 AI 중 하나만 실행한다."""
    acquired_here = False
    if (("--push" in sys.argv or "--cloud" in sys.argv)
            and (os.environ.get("CSOS_AI") or "").strip()):
        import ai_claim
        from claim_guard import require
        me = (os.environ.get("CSOS_AI") or "").strip().lower()
        before = ai_claim.load().get("publish")
        require("publish", "폰 독립 사본 생성·Git 게시", who=me)
        acquired_here = not before
    try:
        return _main()
    finally:
        if acquired_here:
            from claim_guard import release
            release("publish", who=(os.environ.get("CSOS_AI") or "").strip().lower())


if __name__ == "__main__":
    main()
