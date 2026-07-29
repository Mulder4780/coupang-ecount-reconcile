# -*- coding: utf-8 -*-
"""
app_server.py — Coupang Service Operations System 앱 서버 (반응형 웹앱 백엔드)
============================================================
PC에서 실행하면 같은 와이파이의 휴대폰·다른 PC가 브라우저로 접속하는 ERP형 앱.
표준 라이브러리만 사용(설치 0). 데이터는 전부 사내 PC에 남는다(클라우드 전송 없음).

  실행:  python webapp/app_server.py            # 실서비스 (첫 실행 시 PIN 자동 생성)
         python webapp/app_server.py --demo     # 합성데이터 데모 (PIN 0000)
  접속:  PC      → http://localhost:8899
         휴대폰  → http://<PC IP>:8899   (같은 와이파이, 방화벽 허용 필요)

보안: 4자리 PIN(첫 요청 시 입력, 기기에 저장). 사내 LAN 전용 설계 — 외부 인터넷 개방 금지.
"""
import sys, os, re, json, glob, time, threading, random, subprocess, hashlib, io, shutil
from collections import deque
from datetime import datetime, date, timedelta
from email import policy as email_policy
from email.parser import BytesParser
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
sys.path.insert(0, ROOT)
PY = sys.executable
ENV = {**os.environ, "PYTHONIOENCODING": "utf-8"}

DEMO = "--demo" in sys.argv
PORT = int(sys.argv[sys.argv.index("--port") + 1]) if "--port" in sys.argv else 8899
WEBCFG = os.path.join(ROOT, "config", "webapp.json")

# 폰 홈 화면 아이콘이 가리켜야 할 **바뀌지 않는 주소**.
# 터널 주소(trycloudflare)는 띄울 때마다 새로 받으므로 아이콘에 박으면 안 된다.
FIXED_ENTRY = "https://mulder4780.github.io/coupang-ecount-reconcile/"


def load_pin():
    if DEMO:
        return "0000"
    try:
        return json.load(open(WEBCFG, encoding="utf-8"))["pin"]
    except Exception:
        pin = str(random.SystemRandom().randint(1000, 9999))
        os.makedirs(os.path.dirname(WEBCFG), exist_ok=True)
        json.dump({"pin": pin, "port": PORT}, open(WEBCFG, "w", encoding="utf-8"))
        return pin


PIN = load_pin()

# 앱에서 확정한 운영기준은 관리대장 수식과 섞지 않고 작은 런타임 DB로 보관한다.
# reports/는 git 제외 대상이며, 저장 성공 직후 대표보고·확인필요 화면이 같은 값을 읽는다.
POLICY_FILE = os.path.join(ROOT, "reports", "operating_policies.json")


def load_policy_state():
    try:
        data = json.load(open(POLICY_FILE, encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_policy_state(key, value):
    os.makedirs(os.path.dirname(POLICY_FILE), exist_ok=True)
    data = load_policy_state()
    data[str(key)] = value
    tmp = POLICY_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, POLICY_FILE)
    return data


def multipart_parts(content_type, raw):
    """표준 라이브러리만으로 multipart/form-data를 안전하게 푼다."""
    if "multipart/form-data" not in str(content_type or ""):
        raise ValueError("multipart/form-data 형식이 아닙니다")
    head = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("ascii", "ignore")
    msg = BytesParser(policy=email_policy.default).parsebytes(head + raw)
    fields, files = {}, {}
    for part in msg.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        data = part.get_payload(decode=True) or b""
        filename = part.get_filename()
        if filename:
            files[name] = {"filename": os.path.basename(str(filename)), "data": data,
                           "content_type": part.get_content_type()}
        else:
            charset = part.get_content_charset() or "utf-8"
            try:
                fields[name] = data.decode(charset, errors="replace").strip()
            except LookupError:
                fields[name] = data.decode("utf-8", errors="replace").strip()
    return fields, files


def _safe_upload_name(name):
    name = os.path.basename(str(name or "")).strip()
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', "_", name)
    return name[:160] or "kakao.txt"


def _kakao_text_kind(data):
    """두 카톡방 파일이 서로 바뀌어 올라와도 본문 제목으로 자동 분류한다."""
    text = ""
    for enc in ("utf-8-sig", "utf-8", "cp949"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if not text:
        raise ValueError("카카오톡 텍스트 인코딩을 읽을 수 없습니다")
    head = text[:1500]
    if "쿠팡정기점검" in head:
        return "정기점검", text
    if "쿠팡돌발점검" in head:
        return "돌발점검", text
    raise ValueError("파일 첫 부분에서 ‘쿠팡정기점검’ 또는 ‘쿠팡돌발점검’ 대화방을 확인하지 못했습니다")


def save_ryu_upload(fields, files):
    """류지영 업무센터 업로드를 원본 자료에 보존하고 카톡 자동대조 inbox로 넘긴다."""
    from source_dirs import KAKAO_DIR
    needed = [files.get("kakao_regular"), files.get("kakao_emergency")]
    if any(not x for x in needed):
        raise ValueError("정기점검방과 돌발점검방 텍스트 파일을 각각 첨부해 주세요")
    found = {}
    parsed = []
    for f in needed:
        if not f["filename"].lower().endswith(".txt"):
            raise ValueError("카카오톡 대화내역은 .txt 파일만 첨부할 수 있습니다")
        if len(f["data"]) > 20_000_000:
            raise ValueError("카카오톡 텍스트 파일은 각 20MB 이하만 가능합니다")
        kind, _text = _kakao_text_kind(f["data"])
        if kind in found:
            raise ValueError(f"{kind} 대화방 파일이 두 번 첨부되었습니다")
        found[kind] = f
    if set(found) != {"정기점검", "돌발점검"}:
        raise ValueError("정기점검방·돌발점검방 두 종류가 모두 필요합니다")
    evidence = files.get("evidence_file")
    if evidence and evidence.get("data"):
        if len(evidence["data"]) > 25_000_000:
            raise ValueError("추가 근거 파일은 25MB 이하만 가능합니다")
        evidence_ext = os.path.splitext(evidence["filename"])[1].lower()
        if evidence_ext not in (".png", ".jpg", ".jpeg", ".webp", ".pdf", ".xlsx", ".docx", ".txt"):
            raise ValueError("추가 근거는 이미지·PDF·Excel·Word·텍스트 파일만 가능합니다")

    now = datetime.now()
    day_dir = os.path.join(KAKAO_DIR, f"{now:%Y}", f"{now:%m}", f"{now:%Y-%m-%d}")
    inbox = os.path.join(ROOT, "kakao", "inbox")
    os.makedirs(day_dir, exist_ok=True)
    os.makedirs(inbox, exist_ok=True)
    stamp = now.strftime("%Y%m%d_%H%M%S")
    saved = []
    for kind in ("정기점검", "돌발점검"):
        f = found[kind]
        original = _safe_upload_name(f["filename"])
        name = f"(류지영)_{kind}_{stamp}_{original}"
        dest = os.path.join(day_dir, name)
        with open(dest, "wb") as out:
            out.write(f["data"])
        # 자동대조 도구는 로컬 inbox의 txt를 즉시 읽는다. 원본은 Z:에 보존하고,
        # 작은 텍스트 사본만 로컬에 두어 PC 용량과 대조 속도를 모두 지킨다.
        inbox_path = os.path.join(inbox, name)
        shutil.copy2(dest, inbox_path)
        try:
            from kakao.kakao_reconcile import parse_export
            count = len(parse_export(dest))
        except Exception:
            count = 0
        saved.append({"방": kind, "파일": name, "메시지": count})

    evidence_saved = ""
    if evidence and evidence.get("data"):
        evidence_saved = f"(류지영)_추가근거_{stamp}_{_safe_upload_name(evidence['filename'])}"
        with open(os.path.join(day_dir, evidence_saved), "wb") as out:
            out.write(evidence["data"])

    manifest = {
        "등록일시": now.isoformat(timespec="seconds"),
        "등록자": fields.get("submitter") or "류지영",
        "조사기준일": fields.get("survey_date") or f"{now:%Y-%m-%d}",
        "조사메모": fields.get("survey_note") or "",
        "업무구분": fields.get("work_kind") or "",
        "프로젝트NO": fields.get("project_no") or "",
        "캠프명": fields.get("camp_name") or "",
        "담당자": fields.get("assignee") or "",
        "처리상태": fields.get("work_status") or "",
        "완료일": fields.get("completed_date") or "",
        "조치내용": fields.get("action_note") or "",
        "추가근거": evidence_saved,
        "파일": saved,
    }
    with open(os.path.join(day_dir, f"(류지영)_업로드기록_{stamp}.json"), "w", encoding="utf-8") as out:
        json.dump(manifest, out, ensure_ascii=False, indent=2)
    os.makedirs(os.path.join(ROOT, "reports"), exist_ok=True)
    with open(os.path.join(ROOT, "reports", "ryu_submissions.jsonl"), "a", encoding="utf-8") as out:
        out.write(json.dumps(manifest, ensure_ascii=False) + "\n")
    return manifest


# 류지영 업무센터의 입력 항목은 원장의 수식·검증 열을 직접 건드리지 않는다.
# 아래 화이트리스트에 있는 "사람이 확인해서 보충하는 원천 열"만 빈 칸에 한해 기록한다.
RYU_ENTRY_CONFIG = {
    "as": {
        "label": "돌발AS", "sheet": "02_돌발AS접수", "key_col": "접수ID",
        "date_col": "접수일자", "kind": "as",
        "fields": [
            {"name": "담당기사", "label": "담당기사", "type": "text"},
            {"name": "방문예정일", "label": "방문예정일", "type": "date"},
            {"name": "방문예정시간", "label": "방문예정시간", "type": "text"},
            {"name": "작업완료일", "label": "작업완료일", "type": "date"},
            {"name": "재방문여부", "label": "재방문 여부", "type": "select",
             "options": ["예", "아니오"]},
            {"name": "재방문예정일", "label": "재방문 예정일", "type": "date"},
            {"name": "유상·무상·보험", "label": "비용 구분", "type": "select",
             "options": ["유상", "무상", "보험"]},
            {"name": "사진등록", "label": "사진 등록", "type": "select",
             "options": ["있음", "없음"]},
            {"name": "동영상등록", "label": "동영상 등록", "type": "select",
             "options": ["있음", "없음"]},
            {"name": "완료보고서등록", "label": "완료보고서 등록", "type": "select",
             "options": ["있음", "없음"]},
            {"name": "문제내용", "label": "문제 내용", "type": "textarea"},
            {"name": "조치내용", "label": "조치 내용", "type": "textarea"},
            {"name": "완료예정일", "label": "완료 예정일", "type": "date"},
            {"name": "비고", "label": "비고", "type": "textarea"},
        ],
    },
    "pm": {
        "label": "정기점검", "sheet": "04_정기점검", "key_col": "점검ID",
        "date_col": "점검예정일", "kind": "pm",
        "fields": [
            {"name": "점검예정일", "label": "점검 예정일", "type": "date"},
            {"name": "점검예정시간", "label": "점검 예정시간", "type": "text"},
            {"name": "담당기사", "label": "담당기사", "type": "text"},
            {"name": "실제점검일", "label": "실제 점검일", "type": "date"},
            {"name": "점검내용", "label": "점검 내용", "type": "textarea"},
            {"name": "이상발견여부", "label": "이상 발견 여부", "type": "select",
             "options": ["예", "아니오"]},
            {"name": "이상내용", "label": "이상 내용", "type": "textarea"},
            {"name": "돌발AS전환여부", "label": "돌발AS 전환 여부", "type": "select",
             "options": ["예", "아니오"]},
            {"name": "유상추가작업발생", "label": "유상 추가작업 발생", "type": "select",
             "options": ["예", "아니오"]},
            {"name": "추가작업내용", "label": "추가작업 내용", "type": "textarea"},
            {"name": "유상·무상·보험", "label": "비용 구분", "type": "select",
             "options": ["유상", "무상", "보험"]},
            {"name": "점검사진", "label": "점검 사진", "type": "select",
             "options": ["있음", "없음"]},
            {"name": "점검동영상", "label": "점검 동영상", "type": "select",
             "options": ["있음", "없음"]},
            {"name": "점검보고서", "label": "점검 보고서", "type": "select",
             "options": ["있음", "없음"]},
            {"name": "문제내용", "label": "문제 내용", "type": "textarea"},
            {"name": "조치내용", "label": "조치 내용", "type": "textarea"},
            {"name": "비고", "label": "비고", "type": "textarea"},
        ],
    },
    "field": {
        "label": "현장작업", "sheet": "03_현장작업실적", "key_col": "작업ID",
        "date_col": "작업일자", "kind": "field",
        "fields": [
            {"name": "작업일자", "label": "작업일자", "type": "date"},
            {"name": "작업시작시간", "label": "작업 시작시간", "type": "text"},
            {"name": "작업종료시간", "label": "작업 종료시간", "type": "text"},
            {"name": "담당기사", "label": "담당기사", "type": "text"},
            {"name": "실제작업항목", "label": "실제 작업항목", "type": "textarea"},
            {"name": "실제작업상세", "label": "실제 작업상세", "type": "textarea"},
            {"name": "접수외추가작업여부", "label": "접수 외 추가작업", "type": "select",
             "options": ["예", "아니오"]},
            {"name": "추가작업내용", "label": "추가작업 내용", "type": "textarea"},
            {"name": "사용부품", "label": "사용 부품", "type": "text"},
            {"name": "수량", "label": "수량", "type": "number"},
            {"name": "비용구분", "label": "비용 구분", "type": "select",
             "options": ["유상", "무상", "보험"]},
            {"name": "완료여부", "label": "완료 여부", "type": "select",
             "options": ["완료", "미완료"]},
            {"name": "재방문필요", "label": "재방문 필요", "type": "select",
             "options": ["예", "아니오"]},
            {"name": "재방문사유", "label": "재방문 사유", "type": "textarea"},
            {"name": "작업사진", "label": "작업 사진", "type": "select",
             "options": ["있음", "없음"]},
            {"name": "작업동영상", "label": "작업 동영상", "type": "select",
             "options": ["있음", "없음"]},
            {"name": "완료보고서", "label": "완료보고서", "type": "select",
             "options": ["있음", "없음"]},
            {"name": "기사보고내용", "label": "기사 보고내용", "type": "textarea"},
            {"name": "문제내용", "label": "문제 내용", "type": "textarea"},
            {"name": "조치내용", "label": "조치 내용", "type": "textarea"},
            {"name": "완료예정일", "label": "완료 예정일", "type": "date"},
            {"name": "비고", "label": "비고", "type": "textarea"},
        ],
    },
    "settle": {
        "label": "거래서류·청구", "sheet": "06_거래서류청구수금", "key_col": "정산ID",
        "date_col": "작업완료일", "kind": "settle",
        "fields": [
            {"name": "거래명세서번호", "label": "거래명세서 번호", "type": "text"},
            {"name": "거래명세서발행일", "label": "거래명세서 발행일", "type": "date"},
            {"name": "PO필요여부", "label": "PO 필요 여부", "type": "select",
             "options": ["예", "아니오"]},
            {"name": "PO번호", "label": "PO 번호", "type": "text"},
            {"name": "PO발행일", "label": "PO 발행일", "type": "date"},
            {"name": "세금계산서발행일", "label": "세금계산서 발행일", "type": "date"},
            {"name": "청구일", "label": "청구일", "type": "date"},
            {"name": "지급예정일", "label": "지급 예정일", "type": "date"},
            {"name": "입금일", "label": "입금일", "type": "date"},
            {"name": "입금액", "label": "입금액", "type": "number"},
            {"name": "담당자", "label": "담당자", "type": "text"},
            {"name": "문제내용", "label": "문제 내용", "type": "textarea"},
            {"name": "조치내용", "label": "조치 내용", "type": "textarea"},
            {"name": "완료예정일", "label": "완료 예정일", "type": "date"},
            {"name": "비고", "label": "비고", "type": "textarea"},
        ],
    },
}


def _ryu_display_value(value):
    if value in (None, ""):
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _ryu_field_records():
    """03_현장작업실적의 실제 행을 읽기 전용으로 반환한다."""
    if DEMO:
        return [
            {
                "작업ID": f"FW-2607-{i:03d}", "접수ID": f"AS-2607-{i:03d}",
                "프로젝트NO": f"UJ26{2000+i:05d}", "캠프명": f"데모{i}캠프",
                "작업일자": f"2026-07-{i+1:02d}", "담당기사": "김준형",
                "실제작업항목": "현장 조치", "완료여부": "완료" if i < 4 else "미완료",
                "검증결과": "정상" if i < 4 else "확인필요",
            }
            for i in range(1, 7)
        ]
    cached = _fresh("ryu_field")
    if cached is not None:
        return cached
    import openpyxl
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    wanted = [
        "작업ID", "접수ID", "프로젝트NO", "캠프명", "작업일자", "작업시작시간",
        "작업종료시간", "담당기사", "최초접수내용", "실제작업항목", "실제작업상세",
        "접수외추가작업여부", "추가작업내용", "사용부품", "수량", "비용구분",
        "완료여부", "재방문필요", "재방문사유", "작업사진", "작업동영상",
        "완료보고서", "기사보고내용", "관리자검증", "거래명세서반영", "ERP반영",
        "검증자", "검증일", "문제내용", "조치내용", "완료예정일", "비고", "검증결과",
    ]
    out = []
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    try:
        if "03_현장작업실적" not in wb.sheetnames:
            return _store_cache("ryu_field", out)
        ws = wb["03_현장작업실적"]
        header = next(ws.iter_rows(min_row=4, max_row=4, values_only=True))
        idx = {str(v).strip(): i for i, v in enumerate(header) if v not in (None, "")}
        for row in ws.iter_rows(min_row=5, values_only=True):
            key = row[idx["작업ID"]] if "작업ID" in idx and idx["작업ID"] < len(row) else None
            project = row[idx["프로젝트NO"]] if "프로젝트NO" in idx and idx["프로젝트NO"] < len(row) else None
            work_date = row[idx["작업일자"]] if "작업일자" in idx and idx["작업일자"] < len(row) else None
            camp = row[idx["캠프명"]] if "캠프명" in idx and idx["캠프명"] < len(row) else None
            if not key or (not work_date and not camp):
                continue
            iso = norm_date(work_date)
            id_blob = f"{key} {project}"
            if iso:
                if not iso.startswith(APP_YEAR + "-"):
                    continue
            elif not (re.search(r"(?:FW|AS|PM|JS)-?26", id_blob, re.I)
                      or re.search(r"\bUJ26\d{5}\b", id_blob, re.I)):
                continue
            rec = {}
            for name in wanted:
                pos = idx.get(name)
                rec[name] = _ryu_display_value(row[pos]) if pos is not None and pos < len(row) else ""
            out.append(rec)
    finally:
        wb.close()
    out.sort(key=lambda r: (norm_date(r.get("작업일자")) == "",
                            norm_date(r.get("작업일자")), str(r.get("작업ID") or "")))
    return _store_cache("ryu_field", out)


def _ryu_upload_history():
    path = os.path.join(ROOT, "reports", "ryu_submissions.jsonl")
    rows = []
    try:
        with open(path, encoding="utf-8") as f:
            for line in f:
                try:
                    item = json.loads(line)
                except Exception:
                    continue
                day = norm_date(item.get("등록일시") or item.get("조사기준일"))
                if day and not day.startswith(APP_YEAR + "-"):
                    continue
                files = item.get("파일") if isinstance(item.get("파일"), list) else []
                rows.append({
                    "key": str(item.get("등록일시") or ""),
                    "project_no": str(item.get("프로젝트NO") or ""),
                    "camp": str(item.get("캠프명") or ""),
                    "date": day,
                    "status": "원본 저장",
                    "assignee": str(item.get("담당자") or item.get("등록자") or "류지영"),
                    "summary": str(item.get("조치내용") or item.get("조사메모") or
                                   f"카카오톡 원본 {len(files)}개"),
                    "detail": item,
                    "editable": False,
                })
    except Exception:
        pass
    rows.sort(key=lambda r: (r["date"] == "", r["date"], r["key"]))
    return rows


def _ryu_issue_target(row):
    key = str(row.get("업무ID") or row.get("ID") or row.get("원천업무ID") or "").strip()
    upper = key.upper()
    if upper.startswith("AS-"):
        return "as", key
    if upper.startswith("PM-"):
        return "pm", key
    if upper.startswith("FW-"):
        return "field", key
    if upper.startswith("JS-"):
        return "settle", key
    return "", ""


def _ryu_row(rec, key_name, date_names, status_names, summary_names,
             assignee_names=(), editable=True):
    def first(names):
        for name in names:
            value = rec.get(name)
            if value not in (None, ""):
                return str(value)
        return ""
    detail = {str(k): _ryu_display_value(v) for k, v in rec.items() if v not in (None, "")}
    return {
        "key": first((key_name,)),
        "project_no": first(("프로젝트NO",)),
        "camp": first(("캠프명",)),
        "date": norm_date(first(date_names)),
        "status": first(status_names),
        "assignee": first(assignee_names),
        "summary": first(summary_names),
        "detail": detail,
        "editable": bool(editable),
    }


def get_ryu_records():
    """류지영 업무센터: 2026년 업무를 카테고리별 과거→최근 목록으로 제공한다."""
    works = get_works() or {"as": [], "pm": []}
    settlements = get_settlements() or []
    issues = get_issues() or {"rows": []}
    as_rows = [
        _ryu_row(r, "접수ID", ("접수일자", "작업완료일"),
                 ("진행상태", "검증결과"), ("신청내용", "문제내용"),
                 ("담당기사", "담당관리자"),
                 editable=str(r.get("출처") or "") != "ERP")
        for r in drop_side_work(works.get("as") or [])
    ]
    pm_rows = [
        _ryu_row(r, "점검ID", ("점검예정일", "실제점검일"),
                 ("점검상태", "검증결과"), ("점검내용", "이상내용"),
                 ("담당기사", "담당관리자"),
                 editable=str(r.get("출처") or "") not in ("ERP", "정기점검 스케줄 원본"))
        for r in drop_side_work(works.get("pm") or [])
    ]
    field_rows = [
        _ryu_row(r, "작업ID", ("작업일자",), ("완료여부", "검증결과"),
                 ("실제작업항목", "기사보고내용"), ("담당기사",))
        for r in drop_side_work(_ryu_field_records())
    ]
    settle_rows = [
        _ryu_row(r, "정산ID", ("완료일", "명세서발행일", "계산서발행일"),
                 ("상태",), ("업무구분", "적요"), ("담당자",),
                 editable=str(r.get("출처") or "") != "ERP")
        for r in drop_side_work(settlements)
    ]
    target_details = {
        "as": {r["key"]: r.get("detail") or {} for r in as_rows if r.get("key")},
        "pm": {r["key"]: r.get("detail") or {} for r in pm_rows if r.get("key")},
        "field": {r["key"]: r.get("detail") or {} for r in field_rows if r.get("key")},
        "settle": {r["key"]: r.get("detail") or {} for r in settle_rows if r.get("key")},
    }
    issue_rows = []
    for rec in drop_side_work((issues or {}).get("rows") or []):
        row = _ryu_row(rec, "업무ID", ("기준일", "일자", "접수일자", "점검예정일", "완료일"),
                       ("상태", "심각도"), ("문제내용", "내용·근거", "문제유형"),
                       ("담당자",))
        target_category, target_key = _ryu_issue_target(rec)
        if not row["key"]:
            row["key"] = target_key or str(rec.get("ID") or rec.get("원천업무ID") or "")
        if target_category and target_key:
            row["detail"] = {**row["detail"],
                             **target_details.get(target_category, {}).get(target_key, {})}
        row["target_category"] = target_category
        row["target_key"] = target_key
        row["editable"] = bool(target_category and target_key)
        issue_rows.append(row)
    rows = {
        "as": as_rows, "pm": pm_rows, "field": field_rows, "settle": settle_rows,
        "issue": issue_rows, "upload": _ryu_upload_history(),
    }
    for items in rows.values():
        items.sort(key=lambda r: (r.get("date", "") == "", r.get("date", ""), r.get("key", "")))
    def needs_attention(row, completed):
        verify = str((row.get("detail") or {}).get("검증결과") or "").strip()
        if verify and verify != "정상":
            return True
        return str(row.get("status") or "").strip() not in completed
    categories = [
        {"key": "as", "label": "돌발AS", "count": len(as_rows),
         "attention": sum(1 for r in as_rows
                          if needs_attention(r, ("작업완료", "완료", "정상")))},
        {"key": "pm", "label": "정기점검", "count": len(pm_rows),
         "attention": sum(1 for r in pm_rows if needs_attention(r, ("완료", "정상")))},
        {"key": "field", "label": "현장작업", "count": len(field_rows),
         "attention": sum(1 for r in field_rows if needs_attention(r, ("완료", "정상")))},
        {"key": "settle", "label": "거래서류·청구", "count": len(settle_rows),
         "attention": sum(1 for r in settle_rows if needs_attention(
             r, ("정상", "무상/보험", "ERP 계산서(묶음)")))},
        {"key": "issue", "label": "확인 필요", "count": len(issue_rows),
         "attention": len(issue_rows)},
        {"key": "upload", "label": "자료 등록", "count": len(rows["upload"]), "attention": 0},
    ]
    schema = {
        key: {"label": cfg["label"], "fields": cfg["fields"]}
        for key, cfg in RYU_ENTRY_CONFIG.items()
    }
    schema["issue"] = {
        "label": "확인 필요",
        "fields": [
            {"name": "조치내용", "label": "확인·조치 내용", "type": "textarea"},
            {"name": "완료예정일", "label": "완료 예정일", "type": "date"},
            {"name": "비고", "label": "비고", "type": "textarea"},
        ],
    }
    if DEMO:
        updated = datetime.now().isoformat(timespec="minutes")
    else:
        try:
            updated = datetime.fromtimestamp(_master_mtime()).isoformat(timespec="minutes")
        except Exception:
            updated = datetime.now().isoformat(timespec="minutes")
    return {"updated_at": updated, "year": APP_YEAR, "categories": categories,
            "rows": rows, "schema": schema}


def _ryu_find_master_record(category, record_key):
    cfg = RYU_ENTRY_CONFIG[category]
    import openpyxl
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    try:
        if cfg["sheet"] not in wb.sheetnames:
            raise ValueError("대상 시트를 찾지 못했습니다")
        ws = wb[cfg["sheet"]]
        header = next(ws.iter_rows(min_row=4, max_row=4, values_only=True))
        idx = {str(v).strip(): i for i, v in enumerate(header) if v not in (None, "")}
        if cfg["key_col"] not in idx:
            raise ValueError("업무 ID 열을 찾지 못했습니다")
        for row in ws.iter_rows(min_row=5, values_only=True):
            pos = idx[cfg["key_col"]]
            if pos >= len(row) or str(row[pos] or "").strip() != str(record_key).strip():
                continue
            rec = {name: _ryu_display_value(row[i]) for name, i in idx.items() if i < len(row)}
            day = norm_date(rec.get(cfg["date_col"]))
            blob = f"{rec.get(cfg['key_col'], '')} {rec.get('프로젝트NO', '')}"
            if not (day.startswith(APP_YEAR + "-")
                    or re.search(r"(?:AS|PM|FW|JS)-?26", blob, re.I)
                    or re.search(r"\bUJ26\d{5}\b", blob, re.I)):
                raise ValueError(f"{APP_YEAR}년 업무만 입력할 수 있습니다")
            return rec
    finally:
        wb.close()
    raise ValueError("선택한 업무를 최신 관리대장에서 찾지 못했습니다")


def _save_ryu_evidence(file_info, category, record_key):
    if not file_info or not file_info.get("data"):
        return ""
    data = file_info["data"]
    if len(data) > 25_000_000:
        raise ValueError("근거 파일은 25MB 이하여야 합니다")
    ext = os.path.splitext(file_info.get("filename") or "")[1].lower()
    allowed = (".png", ".jpg", ".jpeg", ".webp", ".pdf", ".xlsx", ".docx", ".txt")
    if ext not in allowed:
        raise ValueError("근거는 이미지·PDF·Excel·Word·텍스트 파일만 가능합니다")
    from source_dirs import KAKAO_DIR
    now = datetime.now()
    folder = os.path.join(KAKAO_DIR, f"{now:%Y}", f"{now:%m}", f"{now:%Y-%m-%d}")
    os.makedirs(folder, exist_ok=True)
    safe_key = re.sub(r"[^0-9A-Za-z가-힣_-]", "_", str(record_key))[:50]
    name = (f"(류지영)_업무근거_{category}_{safe_key}_{now:%Y%m%d_%H%M%S}_"
            f"{_safe_upload_name(file_info.get('filename'))}")
    with open(os.path.join(folder, name), "wb") as out:
        out.write(data)
    return name


def save_ryu_entry(fields, files, source_ip=""):
    """선택한 기존 업무의 빈 원천 칸만 큐에 넣고, 첨부 근거는 원본 폴더에 보존한다."""
    requested = str(fields.get("category") or "").strip()
    category = requested
    record_key = str(fields.get("record_key") or "").strip()
    if requested == "issue":
        category = str(fields.get("target_category") or "").strip()
        record_key = str(fields.get("target_key") or "").strip()
    if category not in RYU_ENTRY_CONFIG or not record_key:
        raise ValueError("카테고리와 업무를 먼저 선택해 주세요")
    if DEMO:
        current = {}
        for row in (get_ryu_records().get("rows") or {}).get(category, []):
            if str(row.get("key") or "") == record_key:
                current = row.get("detail") or {}
                break
        if not current:
            raise ValueError("선택한 데모 업무를 찾지 못했습니다")
    else:
        current = _ryu_find_master_record(category, record_key)
    cfg = RYU_ENTRY_CONFIG[category]
    allowed = {item["name"]: item for item in cfg["fields"]}
    if requested == "issue":
        allowed = {k: v for k, v in allowed.items() if k in ("조치내용", "완료예정일", "비고")}
    items = []
    for name, spec in allowed.items():
        raw = str(fields.get(name) or "").strip()
        if raw == "":
            continue
        vtype = spec["type"] if spec["type"] in ("date", "number") else "text"
        if vtype == "date" and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw):
            raise ValueError(f"{spec['label']}은 YYYY-MM-DD 형식이어야 합니다")
        value = raw
        if vtype == "number":
            cleaned = raw.replace(",", "")
            try:
                number = float(cleaned)
            except ValueError:
                raise ValueError(f"{spec['label']}은 숫자로 입력해 주세요")
            value = int(number) if number.is_integer() else number
        items.append({
            "sheet": cfg["sheet"], "key_col": cfg["key_col"], "key": record_key,
            "col": name, "value": value, "vtype": vtype, "only_if_empty": True,
        })
    evidence_name = _save_ryu_evidence(files.get("evidence_file"), category, record_key)
    note = str(fields.get("survey_note") or "").strip()
    evidence = (f"류지영 업무센터 입력({source_ip or '앱'})"
                f"{' · ' + note[:160] if note else ''}"
                f"{' · 근거 ' + evidence_name if evidence_name else ''}")
    for item in items:
        item["evidence"] = evidence
    manifest = {
        "등록일시": datetime.now().isoformat(timespec="seconds"),
        "등록자": "류지영", "입력유형": "업무 보충입력",
        "카테고리": requested, "반영카테고리": category,
        "업무ID": record_key, "프로젝트NO": str(current.get("프로젝트NO") or ""),
        "캠프명": str(current.get("캠프명") or ""), "조사메모": note,
        "추가근거": evidence_name, "입력항목": [item["col"] for item in items],
    }
    if DEMO:
        return {"queued": len(items), "pending": len(items), "manifest": manifest,
                "applying": False, "msg": "데모 입력"}
    os.makedirs(os.path.join(ROOT, "reports"), exist_ok=True)
    with open(os.path.join(ROOT, "reports", "ryu_submissions.jsonl"), "a", encoding="utf-8") as out:
        out.write(json.dumps(manifest, ensure_ascii=False) + "\n")
    if not items and not evidence_name:
        raise ValueError("보충할 항목 또는 근거 파일을 입력해 주세요")
    from ledger_writer import queue_add, load_queue
    added = queue_add(items) if items else 0
    applying = False
    msg = "근거 파일만 저장했습니다"
    if added:
        applying, msg = start_task("writer_apply")
        if not applying:
            defer_task_until_free("writer_apply")
    return {"queued": added, "pending": len(load_queue()), "manifest": manifest,
            "applying": applying, "msg": msg}


def rows_xlsx(payload):
    """담당자 회신용 독립 XLSX를 메모리에서 만든다(관리대장은 절대 열어 저장하지 않는다)."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    title = str(payload.get("title") or "확인목록")[:80]
    rows = payload.get("rows") if isinstance(payload.get("rows"), list) else []
    rows = [r for r in rows[:5000] if isinstance(r, dict)]
    columns = [
        ("프로젝트NO", "프로젝트NO"), ("업무ID", "업무ID"), ("캠프명", "캠프명"),
        ("구분", "구분"), ("확인사항", "확인사항"), ("현재상태", "현재상태"),
        ("기준일자", "기준일자"), ("담당자", "담당자"),
        ("담당자 입력", "담당자입력"), ("처리결과", "처리결과"),
        ("완료일", "완료일"), ("첨부파일·근거", "첨부파일근거"), ("회신메모", "회신메모"),
    ]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "담당자 회신"
    ws.append([x[0] for x in columns])
    head_fill = PatternFill("solid", fgColor="203A75")
    thin = Side(style="thin", color="D9E1EF")
    for c in ws[1]:
        c.font = Font(color="FFFFFF", bold=True)
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(bottom=thin)

    def safe(v):
        if v is None:
            return ""
        s = str(v)
        return "'" + s if s.startswith(("=", "+", "-", "@")) else s

    for r in rows:
        ws.append([safe(r.get(key, "")) for _label, key in columns])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:M{max(1, ws.max_row)}"
    widths = [18, 16, 28, 16, 42, 16, 14, 15, 22, 18, 14, 32, 38]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    guide = wb.create_sheet("작성안내")
    guide.append(["항목", "작성 방법"])
    guide.append(["담당자 입력", "확인한 사실이나 실제 조치내용을 적습니다."])
    guide.append(["처리결과", "완료 / 진행중 / 확인불가 중 하나를 적습니다."])
    guide.append(["완료일", "YYYY-MM-DD 형식으로 적습니다."])
    guide.append(["첨부파일·근거", "파일명, 밴드 글, 카카오톡 근거 또는 URL을 적습니다."])
    guide.append(["회신메모", "추가 확인이 필요한 내용을 자유롭게 적습니다."])
    guide.column_dimensions["A"].width = 20
    guide.column_dimensions["B"].width = 75
    for c in guide[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
    out = io.BytesIO()
    wb.save(out)
    wb.close()
    return out.getvalue(), title

# 브루트포스 차단: IP당 로그인 5회 실패 → 10분 잠금 (외부 터널 공개 대비)
_fails = {}
def _locked(ip):
    c, until = _fails.get(ip, (0, 0))
    return c >= 5 and time.time() < until
def _fail(ip):
    c, _ = _fails.get(ip, (0, 0))
    _fails[ip] = (c + 1, time.time() + 600)
def _ok_login(ip):
    _fails.pop(ip, None)

# ───────────────────────── 작업 러너 ─────────────────────────
TASKS = {
    "daily":         ("전체 대조 실행", [os.path.join(ROOT, "daily_run.py")]),
    "synthetic":     ("합성검증", [os.path.join(ROOT, "tests", "synthetic_check.py")]),
    "writer_prev":   ("자동입력 미리보기", [os.path.join(ROOT, "ledger_writer.py")]),
    "writer_apply":  ("자동입력 반영", [os.path.join(ROOT, "ledger_writer.py"), "--apply"]),
    "upload_dry":    ("전표 전송대기 확인", [os.path.join(ROOT, "ecount_upload.py")]),
    "upload_post":   ("전표 실전송", [os.path.join(ROOT, "ecount_upload.py"), "--post"]),
    "kakao":         ("카톡 대조", [os.path.join(ROOT, "kakao", "kakao_reconcile.py")]),
    "erp_ledger":    ("ERP원장 대조", [os.path.join(ROOT, "erp_ledger_check.py")]),
    "po":            ("쿠팡 PO 대조", [os.path.join(ROOT, "po_reconcile.py")]),
    "erp_docs":      ("ERP 매출서류 대조", [os.path.join(ROOT, "erp_docs_check.py")]),
    "band_ingest":   ("밴드 수집분 반영(24시트+백필)",
                      [os.path.join(ROOT, "band", "ingest.py"), "--sheet", "--backfill"]),
    "band_docs":     ("밴드 문서 이미지 대조", [os.path.join(ROOT, "band", "doc_ocr.py"), "--scan"]),
    "band_docs_apply": ("밴드 문서 → 대장 입력", [os.path.join(ROOT, "band", "doc_ocr.py"), "--scan", "--apply"]),
}
runner = {"busy": False, "task": "", "log": deque(maxlen=3000), "done_at": None,
          "agent_route": ""}
_rlock = threading.Lock()


_codes_cache = {"t": 0, "v": None}


def get_codes():
    """드롭다운 선택지를 **10_코드관리 시트에서** 읽어 온다.

    화면에 목록을 박아 두면 사람이 바뀔 때마다 코드를 고쳐야 하고, 결국 시트와 어긋난다.
    시트가 진실이므로 거기서 읽는다(류지영 매니저가 시트만 고치면 앱도 따라간다).
    관리자검증상태는 10시트에 없어 기본값을 함께 준다.
    """
    if _codes_cache["v"] and time.time() - _codes_cache["t"] < 300:
        return _codes_cache["v"]
    out = {"관리자검증상태": ["일치", "추가작업발생", "작업내용누락", "확인필요"]}
    try:
        import openpyxl
        from ecount_reconcile import load_config, resolve_master
        wb = openpyxl.load_workbook(resolve_master(load_config()["reconcile"]["master_xlsx"]),
                                    read_only=True, data_only=True)
        ws = wb["10_코드관리"]
        rows = list(ws.iter_rows(min_row=4, values_only=True))
        if rows:
            hdr = [str(h).strip() if h else "" for h in rows[0]]
            for i, name in enumerate(hdr):
                if not name:
                    continue
                vals = []
                for r in rows[1:]:
                    v = r[i] if i < len(r) else None
                    if v not in (None, "") and str(v).strip() not in vals:
                        vals.append(str(v).strip())
                if vals:
                    out[name] = vals
        wb.close()
    except Exception as e:
        out["_error"] = str(e)[:80]
    _codes_cache.update({"t": time.time(), "v": out})
    return out


def enqueue_codes(codes):
    """폰이 예약한 프로젝트 코드를 실제 원장 행으로 등록한다.
    쓰기는 전부 ledger_writer(빈 칸만·근거 필수·vN+1)를 거치므로 기존 값은 덮이지 않는다."""
    import project_resolve as P
    ev = P.evidence()
    items, done, skip = [], [], []
    for c in codes:
        r = P.resolve(c, ev)
        if not r.get("ok"):
            skip.append({"code": c, "why": r.get("reason", "형식 오류")})
        elif not app_project_result(c, r):
            skip.append({"code": c, "why": "2026년 업무로 확인되지 않아 제외"})
        elif r["state"] == "등록됨":
            skip.append({"code": c, "why": f"이미 {r['sheet']} {r.get('row')}행에 있습니다"})
        else:
            items += P.row_items(r, ev)
            done.append(c)
            # 같은 요청에 두 건이 오면 뒤엣것이 같은 행을 노린다 — 자리를 미리 물린다
            ev["tail"][r["sheet"]] = r["row"]
    if not items:
        return {"ok": True, "applied": 0, "skipped": skip}
    import ledger_writer as L
    L.queue_add(items)
    p = subprocess.run([PY, os.path.join(ROOT, "ledger_writer.py"), "--apply"],
                       cwd=ROOT, env=ENV, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    ok = p.returncode == 0
    runner["log"].append(f"[폰 예약] {len(done)}건 등록 시도 — {'성공' if ok else '실패'}")
    return {"ok": ok, "applied": len(done) if ok else 0, "codes": done, "skipped": skip,
            "msg": (p.stdout or "").strip().splitlines()[-1:] or [""]}


def start_task(key):
    with _rlock:
        if runner["busy"]:
            return False, "다른 작업 실행 중"
        if key not in TASKS:
            return False, "알 수 없는 작업"
        if DEMO:
            runner["log"].append(f"[데모] '{TASKS[key][0]}' — 합성 환경에서는 실행을 시뮬레이션합니다.")
            return True, "demo"
        # 작업 스크립트는 로컬에서 한 번만 실행한다. AI 연계는 검토·실패 후속조치용
        # 인수인계 큐로 분리해, Claude/Codex가 동시에 관리대장을 쓰지 못하게 한다.
        ticket = None
        try:
            from agent_dispatch import enqueue as enqueue_agent, route_label
            ticket = enqueue_agent(key, TASKS[key][0], TASKS[key][1])
            runner["agent_route"] = route_label(ticket)
        except Exception as exc:
            # AI CLI가 없거나 큐 작성에 실패해도, 사람이 누른 기존 업무 실행은 멈추지 않는다.
            runner["agent_route"] = "AI 연계 상태 확인 실패"
            runner["log"].append(f"[AI 연계] 요청 기록 실패: {str(exc)[:160]}")
        runner["busy"], runner["task"] = True, TASKS[key][0]
        runner["log"].clear()

    def work():
        title, args = TASKS[key]
        local_returncode = 1
        runner["log"].append(f"===== {title} 시작 {datetime.now():%H:%M:%S} =====")
        if runner.get("agent_route"):
            runner["log"].append(f"[AI 연계] {runner['agent_route']} · 로컬 업무 스크립트는 1회만 실행")
        try:
            p = subprocess.Popen([PY] + args, cwd=ROOT, env=ENV, stdout=subprocess.PIPE,
                                 stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace")
            for ln in p.stdout:
                if "UserWarning" not in ln and "warn(msg)" not in ln:
                    runner["log"].append(ln.rstrip())
            p.wait()
            local_returncode = p.returncode
            runner["log"].append(f"===== 종료 (코드 {p.returncode}) =====")
        except Exception as e:
            runner["log"].append(f"오류: {e}")
        finally:
            if ticket:
                try:
                    from agent_dispatch import dispatch_async
                    if dispatch_async(ticket, local_returncode):
                        runner["log"].append("[AI 연계] 로컬 작업 결과를 후속 검토 에이전트에 인계")
                except Exception as exc:
                    runner["log"].append(f"[AI 연계] 후속 검토 실행 실패: {str(exc)[:160]}")
            runner["busy"], runner["done_at"] = False, datetime.now().isoformat()
    threading.Thread(target=work, daemon=True).start()
    return True, "started"


_deferred_tasks = set()


def defer_task_until_free(key, max_wait_seconds=1800):
    """다른 작업 중이면 사람이 다시 누르지 않아도 끝나는 즉시 한 번 실행한다."""
    with _rlock:
        if key not in TASKS or key in _deferred_tasks:
            return False
        _deferred_tasks.add(key)

    def wait_and_start():
        try:
            deadline = time.time() + max_wait_seconds
            while time.time() < deadline:
                with _rlock:
                    busy = bool(runner["busy"])
                if not busy:
                    ok, _ = start_task(key)
                    if ok:
                        return
                time.sleep(5)
            runner["log"].append(f"[자동 대기] {TASKS[key][0]} 실행 대기 시간이 초과되었습니다.")
        finally:
            with _rlock:
                _deferred_tasks.discard(key)

    threading.Thread(target=wait_and_start, daemon=True).start()
    return True


# ───────────────────────── 데이터 ─────────────────────────
_cache = {"t": 0, "settle": None, "status": None}
_readlock = threading.RLock()  # Z:드라이브 엑셀 동시 읽기 직렬화(스레드 충돌 방지)
# ★ RLock이어야 한다: 정산 조회가 락을 쥔 채 업무 조회(대표번호 색인용)를 부르므로
#   일반 Lock이면 같은 스레드에서 자기 자신을 기다리다 멈춘다(실제로 응답 없음 발생)


# ── 날짜 정렬 공통 규칙 ────────────────────────────────────────
# 앱·리포트 어디서나 **과거가 맨 위, 최근이 맨 아래**(오름차순)로 통일한다.
# 새로 추가되는 행도 반드시 이 함수를 거치므로 따로 정렬해 줄 필요가 없다.
DATE_KEYS = {
    "settle": ("완료일", "계산서발행일", "명세서발행일", "입금일"),
    "as":     ("접수일자", "작업완료일", "방문예정일"),
    "pm":     ("점검예정일", "실제점검일"),
}
_DATE_RE = re.compile(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})")
APP_YEAR = "2026"
APP_YEAR_SHORT = APP_YEAR[-2:]
_APP_PROJECT_RE = re.compile(r"(?<![A-Za-z0-9])UJ(?P<yy>\d{2})\d{5}(?!\d)", re.I)
_APP_ID_RE = re.compile(r"(?<![A-Za-z0-9])(?:AS|PM|JS)-(?P<yy>\d{2})\d{2}(?:-|$)", re.I)
_APP_SLIP_RE = re.compile(r"(?<!\d)(?P<yy>\d{2})/\d{2}/\d{2}\s*-\s*\d+")
_OLD_APP_REF_RE = re.compile(
    r"(?<![A-Za-z0-9])UJ25\d{5}(?!\d)|"
    r"(?<![A-Za-z0-9])(?:AS|PM|JS)-25\d{2}-\d{3}(?!\d)|"
    r"(?<!\d)2025[-./]\d{1,2}(?:[-./]\d{1,2})?|2025년|(?<!\d)25(?:년도|년)",
    re.I,
)


def norm_date(v):
    """'2026.6.3' · '2026-06-03 00:00' → '2026-06-03' (문자열 비교로 시간순이 되게)"""
    m = _DATE_RE.search(str(v or ""))
    return "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3))) if m else ""


def row_date(rec, keys=()):
    """행의 대표 날짜. 지정 키를 우선 보고, 없으면 아무 날짜 값이나 찾아 쓴다."""
    for k in keys:
        d = norm_date(rec.get(k))
        if d:
            return d
    for v in rec.values():
        d = norm_date(v)
        if d:
            return d
    return ""


def app_year_record(rec, kind=None):
    """앱에는 2026년 업무만 노출한다.

    원본 엑셀은 그대로 두고 표시 경계에서만 판정한다. 프로젝트NO/업무ID가 있으면
    그것을 날짜보다 우선하며, ERP 묶음처럼 번호가 없는 행은 월·전표·날짜로 판정한다.
    연도를 확인할 수 없는 행도 섞어 보여 주지 않고 제외한다.
    """
    if not isinstance(rec, dict):
        return False

    def years(pattern, values):
        return {m.group("yy") for v in values for m in pattern.finditer(str(v or ""))}

    def date_year(keys):
        vals = [rec.get(k) for k in keys]
        found = {m.group(1) for v in vals
                 for m in re.finditer(r"(?<!\d)(20\d{2})[-./]", str(v or ""))}
        if found:
            return APP_YEAR if found == {APP_YEAR} else "other"
        short = years(_APP_SLIP_RE, vals)
        if short:
            return APP_YEAR if short == {APP_YEAR_SHORT} else "other"
        return ""

    def id_year(keys):
        found = years(_APP_ID_RE, [rec.get(k) for k in keys])
        if found:
            return APP_YEAR if found == {APP_YEAR_SHORT} else "other"
        return ""

    def project_year():
        found = years(_APP_PROJECT_RE,
                      [rec.get(k) for k in ("프로젝트NO", "포함프로젝트", "프로젝트명")])
        if found:
            return APP_YEAR if found == {APP_YEAR_SHORT} else "other"
        return ""

    # 데이터 종류마다 '그 건의 연도'를 정하는 열이 다르다. 수정일·확인일에 2026이
    # 찍혔다고 2025 업무를 되살리지 않도록 업무 발생일을 가장 먼저 본다.
    rules = {
        "as": (("접수일자",), ("접수ID", "업무ID")),
        "pm": (("점검예정일",), ("점검ID", "업무ID")),
        "settle": (("완료일",), ("원천업무ID", "정산ID", "업무ID")),
        "erp": (("월", "전표"), ()),
        "visit": (("방문일",), ()),
        "visit_pending": (("예정일",), ()),
        "unbilled": (("발행일",), ()),
        "issue": (("기준일", "접수일자", "점검예정일", "완료일", "발행일", "일자"),
                  ("업무ID", "접수ID", "점검ID", "정산ID", "ID")),
    }
    if kind in rules:
        date_keys, id_keys = rules[kind]
        y = date_year(date_keys)
        if y:
            return y == APP_YEAR
        y = id_year(id_keys) if id_keys else ""
        if y:
            return y == APP_YEAR
        y = project_year()
        return y == APP_YEAR

    # 표준 종류를 모르는 행은 실제 업무 날짜를 먼저 찾고, 없을 때만 ID·프로젝트로
    # 보완한다. 여러 연도가 함께 든 혼합 행은 통째로 제외한다.
    for keys in (("접수일자",), ("점검예정일",), ("완료일",), ("월", "전표"),
                 ("방문일",), ("발행일",), ("기준일", "일자")):
        y = date_year(keys)
        if y:
            return y == APP_YEAR
    y = id_year(("업무ID", "접수ID", "점검ID", "정산ID", "원천업무ID", "ID"))
    if y:
        return y == APP_YEAR
    y = project_year()
    return y == APP_YEAR


def app_year_rows(rows, kind=None):
    """2026년으로 판정되는 행만 새 목록으로 돌려준다."""
    out = []
    for r in rows:
        if not app_year_record(r, kind):
            continue
        clean = {}
        for k, v in r.items():
            if isinstance(v, str):
                v = _OLD_APP_REF_RE.sub("", v)
                v = re.sub(r"\s*[,·/]\s*(?=([,·/]|$))", "", v).strip(" ,·/")
            clean[k] = v
        out.append(clean)
    return out


def app_project_result(code, result):
    """프로젝트 자동조회 결과도 코드뿐 아니라 내부 날짜·ID까지 2026년인지 검사한다.

    UJ26 코드에 과거 AS-25 작업이 잘못 연결된 실데이터가 있으므로 프로젝트 코드만
    보고 통과시키면 오프라인 앱에서 2025년 내용이 다시 노출된다.
    """
    if not re.fullmatch(r"UJ26\d{5}", str(code or ""), re.I):
        return False
    if not isinstance(result, dict):
        return False
    blob = json.dumps(result, ensure_ascii=False, default=str)
    date_years = set(re.findall(r"(?<!\d)(20\d{2})[-./]", blob))
    id_years = {m.group("yy") for m in _APP_ID_RE.finditer(blob)}
    project_years = {m.group("yy") for m in _APP_PROJECT_RE.finditer(blob)}
    if date_years and date_years != {APP_YEAR}:
        return False
    if id_years and id_years != {APP_YEAR_SHORT}:
        return False
    if project_years and project_years != {APP_YEAR_SHORT}:
        return False
    return True


def sort_by_date(rows, kind, idkey=None):
    """과거 → 최근. 날짜가 없는 행은 맨 뒤(=가장 최근으로 취급), 동률은 ID순."""
    keys = DATE_KEYS.get(kind, ())
    return sorted(rows, key=lambda r: ((d := row_date(r, keys)) == "", d,
                                       str(r.get(idkey) or "") if idkey else ""))


def demo_settlements():
    camps = ["송파5MB(감일동)", "울산2캠프", "인천7MB(마곡동)", "부천3(BUC3)", "대전1캠프",
             "구리1캠프", "제주1Sub-hub", "창원1MB(팔용동)", "군포1Sub-Hub", "광주2Sub-hub"]
    rows = []
    rnd = random.Random(42)
    for i in range(1, 16):
        amt = rnd.choice([380000, 418000, 470800, 760000, 1230000, 1472500])
        st = rnd.choice(["정상", "세금계산서 미발행", "ERP 미확인", "미청구", "입금 대기"])
        d = (date(2026, 7, 1) + timedelta(days=i)).isoformat()
        rows.append({"정산ID": f"JS-2607-{i:03d}", "업무구분": rnd.choice(["돌발AS", "정기점검"]),
                     "캠프명": camps[i % len(camps)], "프로젝트NO": f"UJ26{1000+i}",
                     "원천업무ID": f"AS-2607-{i:03d}",
                     "공급가액": amt, "부가세": int(amt * 0.1), "합계": int(amt * 1.1),
                     "명세서": "있음" if st != "미청구" else "없음",
                     "명세서번호": f"2026/07/{i:02d}-1" if st != "미청구" else "",
                     "명세서발행일": d if st != "미청구" else "",
                     "계산서": "발행" if st == "정상" else "미발행",
                     "계산서발행일": d if st == "정상" else "", "승인번호": "",
                     "입금일": d if st == "정상" else "", "입금액": int(amt * 1.1) if st == "정상" else 0,
                     "미수금": 0 if st == "정상" else int(amt * 1.1), "비용구분": "유상",
                     "상태": st, "완료일": d})
    return sort_by_date(app_year_rows(rows, "settle"), "settle", "정산ID")


def real_settlements():
    from ecount_reconcile import read_ledger, load_config
    cfg = load_config()
    recs = read_ledger(cfg["reconcile"]["master_xlsx"])
    rows = []
    for sid, r in sorted(recs.items()):
        issued = r.get("원장_세금계산서실제발행일") or r.get("원장_세금계산서발행일")
        has_stmt = bool(str(r.get("원장_거래명세서번호") or "").strip())
        if r.get("비용구분") != "유상":
            st = "무상/보험"
        elif not r.get("원장_공급가액"):
            st = "금액 미입력"
        elif not has_stmt:
            st = "미청구(전표 없음)"
        elif not issued:
            st = "세금계산서 미발행"
        elif not r.get("원장_입금일"):
            st = "입금 대기"
        else:
            st = "정상"
        rows.append({"정산ID": sid, "업무구분": r.get("업무구분"), "캠프명": r.get("캠프명"),
                     "프로젝트NO": r.get("프로젝트NO"), "원천업무ID": r.get("원천업무ID"),
                     "공급가액": r.get("원장_공급가액") or 0, "합계": r.get("원장_합계") or 0,
                     "부가세": r.get("원장_부가세"),
                     "명세서": "있음" if has_stmt else "없음",
                     "명세서번호": r.get("원장_거래명세서번호") or "",
                     "명세서발행일": str(r.get("원장_거래명세서발행일") or "")[:10],
                     "계산서": "발행" if issued else "미발행",
                     "계산서발행일": str(issued or "")[:10],
                     "승인번호": r.get("원장_세금계산서승인번호") or "",
                     "청구일": str(r.get("원장_청구일") or "")[:10],
                     "지급예정일": str(r.get("원장_지급예정일") or "")[:10],
                     "입금일": str(r.get("원장_입금일") or "")[:10],
                     "입금액": r.get("원장_입금액") or 0,
                     "미수금": r.get("원장_미수금액") if r.get("원장_미수금액") is not None else "",
                     "비용구분": r.get("비용구분"),
                     "PO필요": r.get("원장_PO필요여부") or "",
                     "PO번호": r.get("원장_PO번호") or "",
                     "PO발행일": str(r.get("원장_PO발행일") or "")[:10],
                     "상태": st, "완료일": str(r.get("작업완료일") or "")[:10]})
    return sort_by_date(app_year_rows(rows, "settle"), "settle", "정산ID")


def real_works():
    """02 돌발AS·04 정기점검 + 27 원본일정 현황 (앱 '업무' 데이터)"""
    import openpyxl
    from verification_sync import derived_field_status_map
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    # 03시트는 접수ID·프로젝트NO가 02 완료행을 순서대로 끌어오는 배열수식이라
    # 캐시가 비어도 같은 순서를 재현해 돌발AS 카드에 현장 검증 상태를 붙인다.
    field_status = derived_field_status_map(wb)
    out = {"as": [], "pm": []}
    spec = {
        # 뒤쪽 3개는 '확인 완료' 표시 — 관리자가 검증한 건인지 카드에서 바로 보이게 한다.
        "02_돌발AS접수": ("as", ["접수ID", "프로젝트NO", "캠프명", "접수일자", "담당기사", "진행상태",
                                "작업완료일", "유상·무상·보험", "신청내용", "긴급도", "방문예정일",
                                "관리자검증상태", "최종확인일", "사진등록", "동영상등록",
                                "완료보고서등록", "ERP등록", "재방문여부",
                                "최초접수외추가작업", "추가작업확인상태",
                                # 밴드 원문 바로가기 — 목록에서 근거를 바로 열어 볼 수 있게 한다
                                "밴드 바로가기",
                                "검증결과", "검증문제코드"]),
        "04_정기점검": ("pm", ["점검ID", "프로젝트NO", "캠프명", "점검예정일", "실제점검일", "점검상태",
                              "담당기사", "이상발견여부", "돌발AS전환여부",
                              "최종확인일(유현민 체크)", "점검사진", "점검보고서",
                              "ERP판매전표", "거래명세서", "담당관리자",
                              "검증결과", "검증문제코드"]),
    }
    for sheet, (key, cols) in spec.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        hdr = next(ws.iter_rows(min_row=4, max_row=4, values_only=True))
        idx = {str(h).strip(): i for i, h in enumerate(hdr) if h is not None}
        for row in ws.iter_rows(min_row=5, values_only=True):
            # ID 열은 수식 — 새로 추가된 행은 엑셀을 열기 전까지 캐시값이 없다.
            # 프로젝트NO를 대체 키로 써서 백필 행도 앱에 표시한다.
            rid = row[idx[cols[0]]] if cols[0] in idx else None
            if not rid and "프로젝트NO" in idx and idx["프로젝트NO"] < len(row):
                rid = row[idx["프로젝트NO"]]
            if not rid:
                continue
            # 행 확장(expand_rows)으로 만들어 둔 **빈 예비행**은 ID 수식이 값을 내므로
            # rid만으로는 걸러지지 않는다. 날짜도 캠프명도 없으면 실제 업무가 아니다.
            _d = row[idx[cols[3]]] if cols[3] in idx and idx[cols[3]] < len(row) else None
            _c = row[idx["캠프명"]] if "캠프명" in idx and idx["캠프명"] < len(row) else None
            if not _d and not _c:
                continue
            rec = {}
            for c in cols:
                v = row[idx[c]] if c in idx and idx[c] < len(row) else None
                rec[c] = str(v)[:10] if hasattr(v, "year") else ("" if v is None else str(v))
            if key == "as":
                rec.update(field_status.get(str(rec.get("프로젝트NO") or "").upper(), {}))
            else:
                rec["검증자"] = rec.get("담당관리자") or ""
                rec["검증일"] = rec.get("최종확인일(유현민 체크)") or ""
            derive_status(rec, key)
            derive_effective_verification(rec, key)
            out[key].append(rec)
    # 류지영 원본 일정은 UJ번호가 없는 캠프·장비 일정이다. 04시트와 같은 캠프·같은 달이면
    # 이미 프로젝트 행으로 표시되므로 중복하지 않고, 아직 04에 없는 미래 월만 읽기 전용으로 보탠다.
    source_schedule = _sheet_records(wb, "27_정기점검원본일정")
    wb.close()
    try:
        out["as"] += erp_work_rows(out["as"], "as")
        out["pm"] += erp_work_rows(out["pm"], "pm")
        idx = build_prj_index(out)
        apply_rep_no(out["as"], idx, "접수ID")
        apply_rep_no(out["pm"], idx, "점검ID")
    except Exception:
        pass
    def pm_camp_key(v):
        return re.sub(r"[^0-9A-Za-z가-힣]", "", re.split(r"[（(]", str(v or ""))[0]).lower()

    represented = set()
    for r in out["pm"]:
        d = norm_date(r.get("점검예정일") or r.get("실제점검일"))
        if d and r.get("캠프명"):
            represented.add((pm_camp_key(r.get("캠프명")), d[:7]))
    for s in source_schedule:
        month = str(s.get("예정월") or "")[:7]
        key = (pm_camp_key(s.get("캠프명")), month)
        if not month.startswith(APP_YEAR + "-") or not key[0] or key in represented:
            continue
        projects = sorted(set(re.findall(r"\bUJ26\d{4,}\b", str(s.get("연결프로젝트NO") or ""),
                                         flags=re.I)))
        out["pm"].append({
            "점검ID": str(s.get("일정ID") or ""),
            "프로젝트NO": projects[0].upper() if len(projects) == 1 else "",
            "캠프명": str(s.get("캠프명") or ""),
            # 일자가 미확정이면 월까지만 보인다. 1일로 만들면 허위 지연 경고가 생긴다.
            "점검예정일": str(s.get("점검예정일") or month),
            "실제점검일": "",
            "점검상태": "예정" if s.get("점검예정일") else "예정월",
            "담당기사": str(s.get("담당기사") or ""),
            "장비수": s.get("장비수") or 0,
            "장비내역": str(s.get("장비내역") or ""),
            "반영상태": str(s.get("반영상태") or ""),
            "원본행": str(s.get("원본행") or ""),
            "원본파일": str(s.get("원본파일") or ""),
            "출처": "정기점검 스케줄 원본",
        })
        # 같은 캠프·같은 달이라도 담당기사/점검일이 다른 원본 그룹은 모두 보여 준다.
        # represented 는 04·ERP의 기존 프로젝트 중복을 막는 용도이고 원본끼리는 합치지 않는다.
    out["as"] = sort_by_date(app_year_rows(out["as"], "as"), "as", "접수ID")
    out["pm"] = sort_by_date(app_year_rows(out["pm"], "pm"), "pm", "점검ID")
    return out


def derive_status(rec, kind):
    """상태 열은 **수식**이라 새로 넣은 행은 엑셀을 한 번 열기 전까지 캐시값이 없다(None).
    그대로 두면 완료된 점검 90여 건이 전부 '미점검'으로 보인다 → 원본 열로 직접 판정한다.
    (판정 규칙은 시트 수식과 동일: 완료일 있으면 완료, 예정일이 지났으면 미점검)"""
    today = date.today().isoformat()
    if kind == "pm":
        if str(rec.get("점검상태") or "").strip():
            return
        if str(rec.get("실제점검일") or "").strip():
            rec["점검상태"] = "완료"
        elif str(rec.get("돌발AS전환여부") or "").strip():
            rec["점검상태"] = "AS전환"
        elif not str(rec.get("점검예정일") or "").strip():
            rec["점검상태"] = ""
        else:
            rec["점검상태"] = "미점검" if str(rec["점검예정일"])[:10] < today else "예정"
    else:
        if str(rec.get("진행상태") or "").strip():
            return
        rec["진행상태"] = "작업완료" if str(rec.get("작업완료일") or "").strip() else "접수"


def derive_effective_verification(rec, kind):
    """ZIP 패치 뒤 Excel을 열기 전에도 확정 상태가 앱에서 즉시 정상으로 보이게 한다.

    수식 캐시는 Excel 재계산 전까지 이전 결과를 유지한다. 여기서는 모든 필수 원인
    값이 명확히 충족된 경우에만 ``정상``으로 승격하고, 하나라도 불명확하면 기존
    검증결과를 그대로 둔다. 따라서 확인되지 않은 건을 정상으로 오판하지 않는다.
    """
    def text(name):
        return str(rec.get(name) or "").strip()

    if kind == "as" and text("진행상태") == "작업완료":
        required = [
            bool(text("담당기사")),
            bool(text("방문예정일")),
            bool(text("작업완료일")),
            text("사진등록") == "등록",
            text("동영상등록") != "누락",
            text("완료보고서등록") == "등록",
            text("유상·무상·보험") not in ("", "미확정"),
            text("ERP등록") in ("완료", "등록완료"),
            bool(text("재방문여부")),
            text("관리자검증상태") in ("일치", "추가작업발생"),
        ]
        if text("최초접수외추가작업") == "있음":
            required.append(text("추가작업확인상태") == "반영완료")
        # 03 현장행이 연결된 건은 문서·ERP·검증자·검증일까지 전부 확인돼야 한다.
        if rec.get("현장작업행"):
            required.extend([
                text("현장관리자검증") in ("일치", "추가작업발생"),
                text("거래명세서반영") == "반영완료",
                text("ERP반영") == "반영완료",
                text("검증자") == "유현민",
                bool(text("검증일")),
            ])
        if all(required):
            rec["검증결과"] = "정상"
            rec["검증문제코드"] = ""
    elif kind == "pm" and text("점검상태") == "완료":
        required = [
            bool(text("실제점검일")),
            text("점검사진") == "등록",
            text("ERP판매전표") in ("완료", "등록완료"),
            text("거래명세서") == "발행완료",
            text("검증자") == "유현민",
            bool(text("검증일")),
        ]
        # 점검보고서는 기존 앱 열 목록에 없던 열이라 새 파일에서는 반드시 등록돼야 한다.
        if "점검보고서" in rec:
            required.append(text("점검보고서") == "등록")
        if all(required):
            rec["검증결과"] = "정상"
            rec["검증문제코드"] = ""


def demo_works():
    rnd = random.Random(7)
    camps = ["송파5MB(감일동)", "울산2캠프", "인천7MB(마곡동)", "대전1캠프", "구리1캠프"]
    techs = ["김준형", "권오철", "김필우", "차동호"]
    a = [{"접수ID": f"AS-2607-{i:03d}", "캠프명": rnd.choice(camps),
          "접수일자": (date(2026, 7, 1) + timedelta(days=i)).isoformat(), "담당기사": rnd.choice(techs),
          "진행상태": rnd.choice(["접수", "방문예정", "작업중", "작업완료"]),
          "작업완료일": "", "유상·무상·보험": rnd.choice(["유상", "무상"]),
          "신청내용": "도어 센서 교체 외", "긴급도": rnd.choice(["보통", "긴급"])} for i in range(1, 11)]
    p = [{"점검ID": f"PM-2607-{i:03d}", "캠프명": rnd.choice(camps),
          "점검예정일": (date(2026, 7, 5) + timedelta(days=i * 2)).isoformat(),
          "실제점검일": "" if i % 3 == 0 else (date(2026, 7, 5) + timedelta(days=i * 2)).isoformat(),
          "점검상태": "예정" if i % 3 == 0 else "완료", "담당기사": rnd.choice(techs),
          "이상발견여부": rnd.choice(["없음", "있음"]), "돌발AS전환여부": "미전환"} for i in range(1, 8)]
    return {"as": a, "pm": p}


def _master_mtime():
    try:
        from ecount_reconcile import load_config, resolve_master
        return os.path.getmtime(resolve_master(load_config()["reconcile"]["master_xlsx"]))
    except Exception:
        return 0


_brief_cache = {"key": None, "value": None}
_brief_lock = threading.Lock()


def _brief_source_key(day):
    """Cheap invalidation key for every source used by daily_brief."""
    try:
        from source_dirs import WORK_LOG_DIR
        work_logs = glob.glob(os.path.join(WORK_LOG_DIR, "**", "*.xlsx"), recursive=True)
        work_log_mt = max((os.path.getmtime(p) for p in work_logs), default=0)
    except Exception:
        work_log_mt = 0
    try:
        event_mt = os.path.getmtime(os.path.join(ROOT, "reports", "manual_daily_events.json"))
    except Exception:
        event_mt = 0
    return day, _master_mtime(), work_log_mt, event_mt


def get_daily_brief(day=None):
    """Return the representative brief without re-reading the large workbook per request.

    The first brief read includes the master workbook and the field work log.  When that
    overlaps the other first-page API reads through Cloudflare, the browser can time out
    and the saved report loses Yoo Subi's daily activity.  Cache the canonical result for
    each exact source revision.  Do not take the global workbook read lock here: a slow
    work-log read must never block settlements/status/works and freeze the whole app.
    """
    day = day or (date.today() - timedelta(days=1)).isoformat()
    key = _brief_source_key(day)
    with _brief_lock:
        if _brief_cache["value"] is not None and _brief_cache["key"] == key:
            return _brief_cache["value"]
        import daily_brief as DB
        result = DB.brief(day, DB.load()[0])
        source_mtime = max((float(v or 0) for v in key[1:]), default=0)
        result["데이터업데이트일시"] = (
            datetime.fromtimestamp(source_mtime).isoformat(timespec="minutes")
            if source_mtime else datetime.now().isoformat(timespec="minutes")
        )
        _brief_cache.update({"key": key, "value": result})
        return result


def _fresh(key):
    """원장이 바뀌면 전체 무효화하고, 그 외에는 항목별 TTL만 적용한다.

    예전에는 120초마다 모든 대형 엑셀 캐시를 한꺼번에 지워 앱이 주기적으로
    20~50초 멈췄다. 원장 변경은 mtime으로 즉시 잡고, 외부 리포트 의존 항목만
    짧게 갱신한다.
    """
    mt = _master_mtime()
    if _cache.get("mt") != mt:
        _cache.clear()
        _cache["mt"] = mt
    ttl = {"status": 300, "exec": 300, "issues": 300, "erpdocs": 300,
           "works": 600, "settle": 600}.get(key, 600)
    if key in _cache and time.time() - _cache.get(key + "_ts", 0) > ttl:
        _cache.pop(key, None)
        _cache.pop(key + "_ts", None)
    return _cache.get(key)


def _store_cache(key, value):
    _cache[key] = value
    _cache[key + "_ts"] = time.time()
    return value


def get_works():
    if DEMO:
        return demo_works()
    with _readlock:
        w = _fresh("works")
        if w:
            return w
        w = real_works()
        return _store_cache("works", w)


def _fmtv(v):
    """01시트 값 표시용: 부동소수 오차 정리·천단위·날짜"""
    if v is None or v == "":
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        n = round(float(v), 2)
        n = int(n) if abs(n - round(n)) < 0.01 else n
        return f"{n:,}"
    return str(v).strip()


def read_exec_report(master):
    """01_대표보고 시트를 구조 그대로 읽는다(엑셀 수식이 곧 집계 로직 — 앱에서 재계산하지 않음)."""
    import openpyxl
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    if "01_대표보고" not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb["01_대표보고"]
    rows = [r for r in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 60), values_only=True)]
    wb.close()
    out = {"meta": {}, "summary": [], "sections": []}
    cur = None
    for i, row in enumerate(rows, 1):
        g = lambda j: row[j] if j < len(row) else None
        a = _fmtv(g(0))
        if i == 4:                                     # 보고일·집계기준일·보고자
            for li, vi in ((0, 1), (3, 4), (6, 7)):
                lab, val = _fmtv(g(li)), _fmtv(g(vi))
                if lab:
                    out["meta"][lab] = val
            continue
        if a.startswith("■"):
            out["summary"].append(a)
            continue
        # 섹션 헤더는 "2.  당일 업무 실적"처럼 숫자 뒤에 한글이 온다.
        # "1. [돌발AS] …" 같은 TOP5 항목은 대괄호로 시작하므로 헤더가 아니다.
        if re.match(r"^\d+\.\s*[가-힣]", a):
            cur = {"title": re.sub(r"\s+", " ", a), "items": [], "lines": []}
            out["sections"].append(cur)
            continue
        if not cur or a.startswith("※"):
            continue
        # 블록 헤더 행: "[ 돌발 AS · 현장 ]  [ 정기점검 ]  [ 거래서류 · 청구 ]"
        # → 열 위치별 그룹을 만들어 이후 행의 항목을 각 그룹에 담는다(AS/점검 구분 유지)
        heads = {li: _fmtv(g(li)).strip("[] ") for li in (0, 3, 6) if _fmtv(g(li)).startswith("[")}
        if heads:
            cur["colgroups"] = {}
            for li, name in heads.items():
                grp = {"name": name, "items": []}
                cur.setdefault("groups", []).append(grp)
                cur["colgroups"][li] = grp
            continue
        for li, vi in ((0, 1), (3, 4), (6, 7)):        # 3열 그룹: 라벨|값
            lab, val = _fmtv(g(li)), _fmtv(g(vi))
            if not lab or lab.startswith("["):
                continue
            if val == "" and len(lab) > 20:            # 값 없는 긴 문장 = 서술형(TOP5 등)
                cur["lines"].append(lab)
            else:
                grp = (cur.get("colgroups") or {}).get(li)
                (grp["items"] if grp else cur["items"]).append([lab, val])
    for s in out["sections"]:
        s.pop("colgroups", None)                       # 내부 매핑은 응답에서 제외
    old = _OLD_APP_REF_RE
    if any(old.search(str(v or "")) for v in out["meta"].values()):
        return {}
    out["summary"] = [x for x in out["summary"] if not old.search(str(x))]
    for s in out["sections"]:
        s["items"] = [x for x in s.get("items", []) if not old.search(" ".join(map(str, x)))]
        s["lines"] = [x for x in s.get("lines", []) if not old.search(str(x))]
        for g in s.get("groups", []):
            g["items"] = [x for x in g.get("items", []) if not old.search(" ".join(map(str, x)))]
    return out


def _sheet_records(wb, sheet):
    """머리글 4행 기준으로 시트를 JSON 안전한 dict 목록으로 읽는다."""
    if sheet not in wb.sheetnames:
        return []
    ws = wb[sheet]
    hdr = next(ws.iter_rows(min_row=4, max_row=4, values_only=True))
    heads = [(i, str(h).strip()) for i, h in enumerate(hdr) if h not in (None, "")]
    out = []
    for row in ws.iter_rows(min_row=5, values_only=True):
        rec = {}
        for i, name in heads:
            value = row[i] if i < len(row) else None
            if isinstance(value, (datetime, date)):
                value = value.strftime("%Y-%m-%d")
            rec[name] = "" if value is None else value
        if any(v not in ("", None) for v in rec.values()):
            out.append(rec)
    return out


def _metric_number(value):
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _report_date(value):
    """대표 예외보고 계산에 쓸 날짜. 모르는 값은 만들지 않고 None으로 둔다."""
    text = norm_date(value)
    if not text:
        return None
    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError:
        return None


def _report_business_days(start, end):
    """start~end(양 끝 포함)의 평일 수. 휴일 마스터가 없으므로 '영업일 추정'으로 표시한다."""
    if not start or not end or start > end:
        return 0
    return sum(1 for n in range((end - start).days + 1)
               if (start + timedelta(days=n)).weekday() < 5)


def _report_complete(value, when=""):
    return bool(_report_date(when)) or str(value or "").strip() in {
        "작업완료", "완료", "정상", "종결", "취소", "철회", "AS전환"
    }


def _report_missing(value, good):
    text = str(value or "").strip()
    return not text or text in {"미등록", "미작성", "미발행", "누락", "미확인"}


def _report_project(rec):
    """화면의 대표번호. 내부 AS/PM/JS 번호를 프로젝트NO처럼 가장하지 않는다."""
    value = str(rec.get("프로젝트NO") or "").strip()
    if re.match(r"^UJ\d{6,}$", value, re.I) or re.match(r"^ERP(?:[-_\s]|$)", value, re.I):
        return value
    for item in rec.values():
        hit = _UJ_RE.search(str(item or ""))
        if hit:
            return hit.group()
    return ""


def representative_summary(works, settlements, base_date=""):
    """유수비 대표 통화 요구를 기존 원장 값으로 계산한 읽기 전용 보고 모델.

    새 업무상태를 추측해 원장에 쓰지 않는다. 현장완료 확정은 완료일/완료상태가 있는 경우만,
    서류 미정리는 그 완료행의 사진·완료보고서·ERP 값이 실제로 빠진 경우만 센다.
    """
    today = _report_date(base_date) or date.today()
    as_rows = list((works or {}).get("as") or [])
    pm_rows = list((works or {}).get("pm") or [])

    def as_item(r, *, issue, state):
        received = _report_date(r.get("접수일자"))
        age = max(0, (today - received).days) if received else -1
        return {
            "프로젝트NO": _report_project(r),
            "ID": str(r.get("접수ID") or ""),
            "레코드ID": str(r.get("접수ID") or ""),
            "종류": "as",
            "프로젝트명": "돌발AS",
            "캠프명": str(r.get("캠프명") or ""),
            "일자": norm_date(r.get("접수일자")),
            "담당자": str(r.get("담당기사") or ""),
            "문제": issue,
            "상태": state,
            "경과일": age,
            "접수내용": str(r.get("신청내용") or ""),
        }

    backlog, paperwork = [], []
    for r in as_rows:
        done = _report_complete(r.get("진행상태"), r.get("작업완료일"))
        received = _report_date(r.get("접수일자"))
        age = max(0, (today - received).days) if received else -1
        if not done:
            grade = ("장기" if age > 30 else "심각" if age > 7 else
                     "경고" if age > 2 else "관심" if age > 1 else "정상")
            backlog.append(as_item(
                r, issue=f"전산상 미완료 · {age if age >= 0 else '날짜 미상'}일 경과",
                state=grade))
        else:
            missing = []
            if _report_missing(r.get("사진등록"), {"등록"}):
                missing.append("사진")
            if _report_missing(r.get("완료보고서등록"), {"등록", "완료"}):
                missing.append("완료보고서")
            if _report_missing(r.get("ERP등록"), {"등록", "완료"}):
                missing.append("ERP")
            if missing:
                paperwork.append(as_item(
                    r, issue="현장완료 · " + "·".join(missing) + " 미정리",
                    state="전산·서류 미정리"))

    backlog.sort(key=lambda r: (r.get("일자") or "9999-99-99", r.get("ID") or ""))
    paperwork.sort(key=lambda r: (r.get("일자") or "9999-99-99", r.get("ID") or ""))
    d1 = [r for r in backlog if r.get("경과일", -1) > 1]
    d2 = [r for r in backlog if r.get("경과일", -1) > 2]
    d7 = [r for r in backlog if r.get("경과일", -1) > 7]
    d30 = [r for r in backlog if r.get("경과일", -1) > 30]

    qmonth = ((today.month - 1) // 3) * 3 + 1
    qstart = date(today.year, qmonth, 1)
    qend = (date(today.year + (1 if qmonth == 10 else 0),
                 1 if qmonth == 10 else qmonth + 3, 1) - timedelta(days=1))
    qrows = [r for r in pm_rows if
             (lambda d: bool(d and qstart <= d <= qend))(_report_date(r.get("점검예정일")))]
    qdone = [r for r in qrows if
             _report_complete(r.get("점검상태"), r.get("실제점검일"))]
    total_days = (qend - qstart).days + 1
    elapsed_days = min(total_days, max(0, (today - qstart).days + 1))
    target = len(qrows)
    expected = round(target * elapsed_days / total_days) if total_days else 0
    actual = len(qdone)
    gap = actual - expected
    shortage_ratio = ((expected - actual) / expected * 100) if expected and actual < expected else 0
    signal = "적색" if shortage_ratio > 10 else "황색" if shortage_ratio >= 5 else "녹색"
    remaining = max(0, target - actual)
    remaining_business = _report_business_days(max(today + timedelta(days=1), qstart), qend)
    required_daily = remaining / remaining_business if remaining_business else (float(remaining) if remaining else 0)
    required_weekly = required_daily * 5
    techs = set()
    for r in qrows:
        techs.update(x.strip() for x in re.split(r"[,·/]|\s{2,}", str(r.get("담당기사") or ""))
                     if x.strip())
    available = len(techs)
    per_tech = required_daily / available if available else 0
    elapsed_business = _report_business_days(qstart, min(today, qend))
    total_business = _report_business_days(qstart, qend)
    forecast = min(target, round(actual / elapsed_business * total_business)) if elapsed_business else 0
    forecast_shortfall = max(0, target - forecast)

    statement_groups = {
        "돌발 AS": [], "정기점검": [], "신규·납품·설치": [], "기타": []
    }

    def statement_kind(r):
        text = str(r.get("업무구분") or "")
        if "돌발" in text or text.upper() == "AS":
            return "돌발 AS"
        if "정기" in text or "점검" in text:
            return "정기점검"
        if any(x in text for x in ("신규", "납품", "설치")):
            return "신규·납품·설치"
        return "기타"

    for r in settlements or []:
        # 유상임이 확인된 정산만 발행대상으로 센다. 비용구분 미확정은 확인 필요로 남긴다.
        if "유상" not in str(r.get("비용구분") or ""):
            continue
        statement_groups[statement_kind(r)].append(r)

    statement_rows = []
    unissued_rows = []
    for kind, rows in statement_groups.items():
        issued = [r for r in rows if str(r.get("명세서번호") or "").strip() or
                  str(r.get("명세서") or "").strip() in {"있음", "발행", "발행완료", "완료"}]
        unissued = [r for r in rows if r not in issued]
        dates = sorted(d for d in (norm_date(r.get("완료일")) for r in rows) if d)
        item = {
            "업무유형": kind, "발행대상": len(rows), "발행완료": len(issued),
            "미발행": len(unissued), "발행률": round(len(issued) / len(rows) * 100) if rows else None,
            "대상기간": f"{dates[0]} ~ {dates[-1]}" if dates else "대상 없음",
            "합계금액": sum(_metric_number(r.get("공급가액")) for r in rows),
        }
        statement_rows.append(item)
        for r in unissued:
            unissued_rows.append({
                "프로젝트NO": _report_project(r),
                "ID": str(r.get("정산ID") or ""),
                "레코드ID": str(r.get("정산ID") or ""),
                "종류": "settle",
                "프로젝트명": str(r.get("업무구분") or kind),
                "캠프명": str(r.get("캠프명") or ""),
                "일자": norm_date(r.get("완료일")),
                "담당자": str(r.get("담당자") or ""),
                "문제": "발행 대상이나 거래명세서 미발행",
                "상태": "미발행",
                "금액": _metric_number(r.get("공급가액")),
            })

    policy_names = [
        "돌발 AS·정기점검 거래명세서 묶음기간",
        "여러 거래명세서의 세금계산서 합산 기준",
        "세금계산서 건별·월합계 발행 기준",
        "PO 수신 전 세금계산서 발행 가능 여부",
        "다음 청구주기 이월 조건·승인자",
    ]
    saved_policy = load_policy_state()
    policies, confirmed_policies = [], []
    for name in policy_names:
        state = saved_policy.get(name) if isinstance(saved_policy.get(name), dict) else {}
        item = {
            "기준": name,
            "상태": str(state.get("상태") or "확인 필요"),
            "확정내용": str(state.get("확정내용") or ""),
            "저장일시": str(state.get("저장일시") or ""),
            "저장자": str(state.get("저장자") or ""),
        }
        (confirmed_policies if item["상태"] == "확정" and item["확정내용"] else policies).append(item)

    one_line = (
        f"전산상 미완료 돌발 AS {len(backlog)}건 중 D+2 초과 {len(d2)}건, "
        f"현장완료·서류미정리 {len(paperwork)}건입니다. "
        f"{today.month}월 기준 정기점검은 목표 누계 {expected}건 대비 {actual}건"
        f"({gap:+d}건), 거래명세서 미발행 대상은 {len(unissued_rows)}건입니다."
    )
    return {
        "meta": {
            "집계기준일": today.isoformat(), "적용마감시간": "관리대장 최신 저장 시점",
            "데이터최종갱신일": datetime.now().isoformat(timespec="seconds"),
            "원천업무건수": len(as_rows) + len(pm_rows) + len(settlements or []),
            "검증되지않은건수": len(backlog) + len(paperwork) + len(unissued_rows),
            "필터조건": f"{APP_YEAR}년·정상 상세 기본 접힘",
        },
        "한줄종합보고": one_line,
        "돌발AS": {
            "전산상미완료": len(backlog), "현장완료서류미정리": len(paperwork),
            "D+1초과": len(d1), "D+2초과": len(d2), "7일초과": len(d7),
            "30일초과": len(d30), "대표지속보고": len(d2),
            "미완료목록": backlog, "서류미정리목록": paperwork,
        },
        "정기점검": {
            "분기": f"{qstart.month}~{qend.month}월", "분기시작일": qstart.isoformat(),
            "분기종료일": qend.isoformat(), "전체대상": target,
            "경과율": round(elapsed_days / total_days * 100, 1) if total_days else 0,
            "목표누계": expected, "실제완료": actual, "계획대비": gap,
            "잔여대상": remaining, "잔여평일추정": remaining_business,
            "필요일일처리량": round(required_daily, 2),
            "필요주간처리량": round(required_weekly, 2),
            "투입기사수": available, "기사1인당필요일일": round(per_tech, 2),
            "예상완료": forecast, "예상미달": forecast_shortfall, "신호": signal,
            "목록": [{
                "프로젝트NO": _report_project(r), "ID": str(r.get("점검ID") or ""),
                "레코드ID": str(r.get("점검ID") or ""), "종류": "pm",
                "프로젝트명": "정기점검", "캠프명": str(r.get("캠프명") or ""),
                "일자": norm_date(r.get("점검예정일")), "담당자": str(r.get("담당기사") or ""),
                "문제": "분기 점검 대상", "상태": str(r.get("점검상태") or ""),
            } for r in qrows],
        },
        "거래명세서": {"업무유형별": statement_rows, "미발행목록": unissued_rows},
        "업무기준확인필요": policies,
        "업무기준확정": confirmed_policies,
    }


# ── 철거·신규납품 숨김 (사용자 지시 2026-07-29) ──────────────────────
#   "철거 및 신규건은 DB만 보관하고 앱에 표시하지마 / 추후에 앱에 추가할 수도 있으니
#    감안해서 정리해줘" → 원장에서 지우지 않고 **화면에서만** 뺀다.
#   대표보고(보고 탭)의 숫자는 서버가 계산해서 내려주므로 앱쪽 필터가 닿지 않는다.
#   그래서 여기서도 같은 규칙을 한 번 더 적용한다. 켜려면 아래 한 줄만 True 로 바꾼다
#   (index.html 의 SHOW_SIDE_WORK 도 같이 켠다 — 검증 [76]이 둘의 짝을 지킨다).
SHOW_SIDE_WORK = False
SIDE_WORK_RE = re.compile(r"철거|이전|납품|설치|계단|안전바|경보장치|메자닌")


def is_side_work(r):
    if SHOW_SIDE_WORK or not isinstance(r, dict):
        return False
    return any(SIDE_WORK_RE.search(str(r.get(k) or ""))
               for k in ("업무구분", "업무유형", "구분", "종류", "품목"))


def drop_side_work(rows):
    return [r for r in (rows or []) if not is_side_work(r)]


def get_representative_report():
    works = get_works()
    works = {k: (drop_side_work(v) if isinstance(v, list) else v) for k, v in (works or {}).items()}
    return representative_summary(works, drop_side_work(get_settlements()))


def read_exec_details(master, base_date=""):
    """대표보고 3·4절의 숫자를 만든 **동일 원천 행**을 건별 목록으로 돌려준다.

    앱에서 숫자를 다시 추정하면 엑셀 카드와 목록 건수가 갈릴 수 있다. 따라서
    01_대표보고/00_대시보드 수식이 참조하는 열과 조건을 그대로 재현한다.
    """
    import openpyxl
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    s06 = _sheet_records(wb, "06_거래서류청구수금")
    s07 = _sheet_records(wb, "07_불일치누락현황")
    s15 = _sheet_records(wb, "15_세금계산서관리")
    s17 = _sheet_records(wb, "17_문서대조현황")
    s02 = _sheet_records(wb, "02_돌발AS접수")
    s04 = _sheet_records(wb, "04_정기점검")
    wb.close()

    # 프로젝트NO·업무ID·정산ID 어느 것을 눌러도 캠프와 대표 날짜를 찾을 수 있게 한다.
    lookup = {}

    def remember(rec, ids, camp, when, kind="", owner=""):
        info = {
            "프로젝트NO": str(rec.get("프로젝트NO") or ""),
            "캠프명": str(rec.get(camp) or ""),
            "일자": norm_date(rec.get(when)),
            "프로젝트명": str(rec.get(kind) or ""),
            "담당자": str(rec.get(owner) or ""),
        }
        for key in ids:
            value = str(rec.get(key) or "").strip()
            if value:
                lookup.setdefault(value, info)

    for r in s06:
        remember(r, ("정산ID", "원천업무ID", "프로젝트NO"), "캠프명",
                 "작업완료일(자동)", "업무구분", "담당자")
    for r in s02:
        remember(r, ("접수ID", "프로젝트NO"), "캠프명", "접수일자", "", "담당기사")
    for r in s04:
        remember(r, ("점검ID", "프로젝트NO"), "캠프명", "점검예정일", "", "담당기사")

    rep_idx = build_prj_index({"as": s02, "pm": s04})

    def joined(rec):
        for key in ("프로젝트NO", "정산ID", "원천업무ID", "업무ID", "접수ID", "점검ID"):
            value = str(rec.get(key) or "").strip()
            if value in lookup:
                return lookup[value]
        return {}

    def detail(rec, *, when="", amount=0, issue="", status="", source=""):
        base = joined(rec)
        project = str(rec.get("프로젝트NO") or base.get("프로젝트NO") or "")
        if not project:
            candidate = {**base, **rec}
            candidate.setdefault("완료일", rec.get(when) if when else base.get("일자"))
            project, _ = rep_no(
                candidate, rep_idx,
                str(rec.get("거래명세서번호") or rec.get("명세서번호") or ""))
        rid = str(rec.get("정산ID") or rec.get("원천업무ID") or rec.get("업무ID")
                  or rec.get("접수ID") or rec.get("점검ID") or project or "")
        record_kind = ("settle" if rid.startswith("JS-") else
                       "as" if rid.startswith("AS-") else
                       "pm" if rid.startswith("PM-") else "")
        return {
            "프로젝트NO": project,
            "ID": rid,
            "레코드ID": rid,
            "종류": record_kind,
            "프로젝트명": str(rec.get("업무구분") or base.get("프로젝트명") or source or ""),
            "캠프명": str(rec.get("캠프명") or base.get("캠프명") or ""),
            "일자": norm_date(rec.get(when)) if when else str(base.get("일자") or ""),
            "금액": _metric_number(amount),
            "문제": str(issue or ""),
            "상태": str(status or ""),
            "담당자": str(rec.get("담당자") or rec.get("담당기사") or base.get("담당자") or ""),
            "출처": source,
        }

    def is_2026_settlement(r):
        d = norm_date(r.get("작업완료일(자동)") or r.get("작업완료일"))
        return bool(r.get("정산ID")) and d.startswith(APP_YEAR + "-")

    def sorted_rows(rows):
        return sort_by_date(rows, "metric", "ID")

    details = {}

    def add(label, rows, basis, kind):
        # 임의 기준일 보고를 만들 때 미래 행이 과거 캡처에 섞이지 않게 한다.
        # 일자가 없는 현재 잔여·문서 경고는 원천상 시점을 판별할 수 없어 그대로 남긴다.
        if base_date:
            rows = [r for r in rows
                    if not norm_date(r.get("일자")) or norm_date(r.get("일자")) <= base_date]
        rows = sorted_rows(rows)
        details[label] = {
            "rows": rows,
            "basis": basis,
            "kind": kind,
            "count": len(rows),
            "amount": sum(_metric_number(r.get("금액")) for r in rows),
        }

    # 3. 당일 금액 · 잔여 현황 — 06시트의 대표보고 수식과 같은 조건.
    add("청구액 (당일)",
        [detail(r, when="거래명세서발행일", amount=r.get("거래명세서합계"),
                status=r.get("청구상태"), source="거래명세서")
         for r in s06 if is_2026_settlement(r) and norm_date(r.get("거래명세서발행일")) == base_date],
        f"06_거래서류청구수금 · 거래명세서발행일={base_date} · 거래명세서합계", "amount")
    add("세금계산서 발행액 (당일)",
        [detail(r, when="세금계산서발행일", amount=r.get("세금계산서합계"),
                status=r.get("청구상태"), source="세금계산서")
         for r in s06 if is_2026_settlement(r) and norm_date(r.get("세금계산서발행일")) == base_date],
        f"06_거래서류청구수금 · 세금계산서발행일={base_date} · 세금계산서합계", "amount")
    add("입금액 (당일)",
        [detail(r, when="입금일", amount=r.get("입금액"),
                status=r.get("청구상태"), source="입금")
         for r in s06 if is_2026_settlement(r) and norm_date(r.get("입금일")) == base_date],
        f"06_거래서류청구수금 · 입금일={base_date} · 입금액", "amount")
    add("잔여 미청구액",
        [detail(r, amount=r.get("미청구액"), issue=r.get("문제내용"),
                status=r.get("청구상태"), source="미청구")
         for r in s06 if is_2026_settlement(r) and _metric_number(r.get("미청구액")) > 0],
        "06_거래서류청구수금 · 2026년 정산ID 보유 · 미청구액>0", "amount")
    add("잔여 미수금액",
        [detail(r, amount=r.get("미수금액"), issue=r.get("문제내용"),
                status=r.get("청구상태"), source="미수")
         for r in s06 if is_2026_settlement(r) and _metric_number(r.get("미수금액")) > 0],
        "06_거래서류청구수금 · 2026년 정산ID 보유 · 미수금액>0", "amount")
    add("작업금액 불일치 (현재)",
        [detail(r, amount=r.get("작업대비거래명세서차액"), issue=r.get("문제내용"),
                status=r.get("검증결과"), source="금액 불일치")
         for r in s06 if is_2026_settlement(r)
         and _metric_number(r.get("작업대비거래명세서차액")) != 0
         and _metric_number(r.get("거래명세서합계")) != 0
         and str(r.get("업무구분") or "") != "신규·납품·설치"],
        "06_거래서류청구수금 · 작업/명세서 차액≠0 · 명세서합계≠0 · 신규납품 제외", "risk")

    # 4. 리스크 — 00_대시보드 수식의 실제 원천행.
    issue_2026 = [r for r in s07 if str(r.get("업무기준연도(자동·숨김)") or "") == APP_YEAR]
    unique_work = {}
    for r in issue_2026:
        key = str(r.get("최상위 업무키") or "").strip()
        if not key:
            continue
        if key not in unique_work:
            unique_work[key] = detail(
                {**r, "업무ID": r.get("원천업무ID") or key},
                issue=r.get("문제상세"), status=r.get("조치상태"), source="확인필요")
        elif r.get("문제상세"):
            old = unique_work[key]["문제"]
            new = str(r.get("문제상세"))
            if new not in old:
                unique_work[key]["문제"] = " · ".join(x for x in (old, new) if x)
    add("문제 업무 건수(중복 제거)", list(unique_work.values()),
        "07_불일치누락현황 · 업무기준연도=2026 · 최상위 업무키 중복 제거", "risk")

    add("문서 경고 총계",
        [detail(r, issue=r.get("경고내용"), status=r.get("우선순위"),
                source="문서 경고")
         for r in s17 if str(r.get("정산ID") or "").startswith("JS-26")
         and str(r.get("경고내용") or "").strip()],
        "17_문서대조현황 · 정산ID=JS-26* · 경고내용 있음", "risk")

    tax_rows = []
    for r in s15:
        if not str(r.get("정산ID") or "").startswith("JS-26"):
            continue
        if str(r.get("발행기한임박여부") or "") == "예":
            tax_rows.append(detail(r, when="법정발행기한", amount=r.get("발행금액"),
                                   issue="세금계산서 발행기한 임박", status=r.get("발행상태(자동)"),
                                   source="세금계산서 기한"))
        if str(r.get("기한초과여부") or "") == "예":
            tax_rows.append(detail(r, when="법정발행기한", amount=r.get("발행금액"),
                                   issue="세금계산서 발행기한 초과", status=r.get("발행상태(자동)"),
                                   source="세금계산서 기한"))
    add("세금계산서 기한 임박·초과", tax_rows,
        "15_세금계산서관리 · JS-26* · 발행기한임박=예 또는 기한초과=예", "risk")

    add("PO 미발행 · 확인필요",
        [detail(r, issue=r.get("경고내용") or r.get("PO상태"),
                status=r.get("PO상태"), source="PO")
         for r in s17 if str(r.get("정산ID") or "").startswith("JS-26")
         and str(r.get("PO상태") or "") in ("PO 발행대기", "PO관리행 없음")],
        "17_문서대조현황 · JS-26* · PO상태=PO 발행대기/PO관리행 없음", "risk")
    add("거래명세서 미작성",
        [detail(r, issue=r.get("경고내용") or "거래명세서 미작성",
                status=r.get("거래명세서상태"), source="거래명세서")
         for r in s17 if str(r.get("정산ID") or "").startswith("JS-26")
         and str(r.get("거래명세서상태") or "") == "미작성"],
        "17_문서대조현황 · JS-26* · 거래명세서상태=미작성", "risk")
    add("아리바 청구 미등록",
        [detail(r, when="법정발행기한", amount=r.get("발행금액"),
                issue="아리바 청구 등록대기", status=r.get("아리바청구상태"),
                source="아리바")
         for r in s15 if str(r.get("정산ID") or "").startswith("JS-26")
         and str(r.get("아리바청구상태") or "") == "등록대기"],
        "15_세금계산서관리 · JS-26* · 아리바청구상태=등록대기", "risk")

    problem_rows = [
        detail({**r, "업무ID": r.get("원천업무ID")},
               amount=r.get("미청구액") or r.get("미수금액"),
               issue=r.get("문제상세"), status=r.get("조치상태"), source="문제 행")
        for r in issue_2026 if str(r.get("원천업무ID") or "").strip()
    ]
    add("문제 프로젝트 / 문제 행", problem_rows,
        "07_불일치누락현황 · 업무기준연도=2026 · 원천업무ID 보유 행", "risk")
    details["문제 프로젝트 / 문제 행"]["project_count"] = len({
        str(r.get("프로젝트NO") or "").strip() for r in issue_2026
        if str(r.get("프로젝트NO") or "").strip()
    })
    return details


def get_exec_report(day=None):
    requested = norm_date(day)
    if DEMO:
        labels = [
            "청구액 (당일)", "세금계산서 발행액 (당일)", "입금액 (당일)",
            "잔여 미청구액", "잔여 미수금액", "작업금액 불일치 (현재)",
            "문제 업무 건수(중복 제거)", "문서 경고 총계", "세금계산서 기한 임박·초과",
            "PO 미발행 · 확인필요", "거래명세서 미작성", "아리바 청구 미등록",
            "문제 프로젝트 / 문제 행",
        ]
        details = {label: {"rows": [], "basis": "합성 데모 · 2026년 원천 행", "kind": "risk",
                           "count": 0, "amount": 0} for label in labels}
        money = {"프로젝트NO": "UJ261001", "ID": "JS-2607-001", "레코드ID": "JS-2607-001",
                 "종류": "settle", "프로젝트명": "돌발AS", "캠프명": "울산2캠프",
                 "일자": "2026-07-24", "금액": 22000, "문제": "미청구 합성 예시",
                 "상태": "미청구", "담당자": "김준형", "출처": "미청구"}
        details["잔여 미청구액"].update(rows=[money], count=1, amount=22000, kind="amount")
        risk_rows = []
        for i in range(75):
            n = i % 15 + 1
            risk_rows.append({
                "프로젝트NO": f"UJ26{1000+n}", "ID": f"JS-2607-{n:03d}",
                "레코드ID": f"JS-2607-{n:03d}", "종류": "settle",
                "프로젝트명": "돌발AS" if n % 2 else "정기점검",
                "캠프명": f"합성 캠프 {n}", "일자": f"2026-07-{n:02d}", "금액": 0,
                "문제": f"문서 확인 필요 합성 항목 {i+1}", "상태": "P1",
                "담당자": "김준형", "출처": "문서 경고",
            })
        details["문서 경고 총계"].update(rows=risk_rows, count=len(risk_rows))
        return {
            "meta": {"보고일": requested or "2026-07-25",
                     "집계기준일": requested or "2026-07-24", "보고자": "유현민"},
            "summary": ["■ 데모 요약"],
            "sections": [
                {"title": "1. 당일 업무 실적",
                 "items": [["신규 접수", "3"], ["작업 완료", "1"]], "lines": []},
                {"title": "2. 당일 금액 · 잔여 현황",
                 "items": [["청구액 (당일)", "0"], ["세금계산서 발행액 (당일)", "0"],
                           ["입금액 (당일)", "0"], ["잔여 미청구액", "22,000"],
                           ["잔여 미수금액", "0"], ["작업금액 불일치 (현재)", "0"]]},
                {"title": "3. 리스크 (현재 기준)",
                 "items": [["문제 업무 건수(중복 제거)", "0"], ["문서 경고 총계", "75"],
                           ["세금계산서 기한 임박·초과", "0"], ["PO 미발행 · 확인필요", "0"],
                           ["거래명세서 미작성", "0"], ["아리바 청구 미등록", "0"],
                           ["문제 프로젝트 / 문제 행", "0개 / 0건"]]},
            ],
            "details": details,
        }
    with _readlock:
        r = _fresh("exec") if not requested else None
        if r:
            return r
        from ecount_reconcile import load_config, resolve_master
        master = resolve_master(load_config()["reconcile"]["master_xlsx"])
        r = read_exec_report(master)
        base = requested or norm_date((r.get("meta") or {}).get("집계기준일")
                                      or (r.get("meta") or {}).get("보고일"))
        r["details"] = read_exec_details(master, base)
        if requested:
            r.setdefault("meta", {})["보고일"] = requested
            r["meta"]["집계기준일"] = requested
            # 금액·리스크 타일은 선택일로 다시 계산한 상세 집계와 맞춘다.
            for sec in r.get("sections", []):
                for item in sec.get("items", []):
                    d = r["details"].get(str(item[0]))
                    if not d:
                        continue
                    if str(item[0]) == "문제 프로젝트 / 문제 행":
                        item[1] = f"{d.get('project_count', 0)}개 / {d.get('count', 0)}건"
                    elif d.get("kind") == "amount":
                        item[1] = f"{int(d.get('amount') or 0):,}"
                    else:
                        item[1] = f"{int(d.get('count') or 0):,}"
            return r
        return _store_cache("exec", r)


def get_issues():
    """07_불일치누락현황 — 엑셀의 '검증 안 된·확인해야 할' 항목 그대로"""
    if DEMO:
        return {"rows": [{"문제유형": "세금계산서 미발행", "업무ID": "JS-2607-002", "캠프명": "울산2캠프",
                          "문제내용": "명세서 발행 후 계산서 미발행", "담당자": "변재선(회계)"}], "cols": []}
    r = _fresh("issues")
    if r:
        return r
    import openpyxl
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
    rows = []
    # 1순위: 관리대장 통합 시트 23_확인필요현황 (에이전트가 매일 갱신 — 단일 엑셀 관리)
    if "23_확인필요현황" in wb.sheetnames:
        ws = wb["23_확인필요현황"]
        hdr = [str(h).strip() for h in next(ws.iter_rows(min_row=4, max_row=4, values_only=True)) if h]
        for row in ws.iter_rows(min_row=5, values_only=True):
            vals = {hdr[i]: ("" if i >= len(row) or row[i] is None else str(row[i]))
                    for i in range(len(hdr))}
            if any(v for v in vals.values()):
                rows.append(vals)
        wb.close()
        from responsibility import assign_issue_row
        rows = [assign_issue_row(row) for row in rows]
        rows = app_year_rows(apply_rep_no(rows), "issue")
        out = {"rows": sort_by_date(rows, "check"), "cols": hdr, "source": "23_확인필요현황"}
        return _store_cache("issues", out)
    if "07_불일치누락현황" in wb.sheetnames:
        ws = wb["07_불일치누락현황"]
        hdr = next(ws.iter_rows(min_row=4, max_row=4, values_only=True))
        heads = [(i, str(h).strip()) for i, h in enumerate(hdr) if h is not None]
        for row in ws.iter_rows(min_row=5, values_only=True):
            vals = {h: ("" if i >= len(row) or row[i] is None else
                        (str(row[i])[:10] if hasattr(row[i], "year") else str(row[i])))
                    for i, h in heads}
            if any(v for v in vals.values()):
                rows.append(vals)
            if len(rows) >= 300:
                break
    wb.close()
    # 대조 결과(밴드·카톡·ERP원장·쿠팡PO) 통합 — 07시트 위에 얹어 한 화면에서 전부 확인
    merged = []
    try:
        from findings_export import latest_csv
        for src, pat, filt in (("밴드 게시 미확인", "밴드대조_*.csv", lambda r: r.get("밴드게시") == "미확인"),
                               ("카톡 보고 미확인", "카톡대조_*.csv", lambda r: r.get("카톡보고") == "미확인"),
                               ("ERP원장 문제", "ERP원장대조_*.csv", lambda r: True),
                               ("쿠팡PO 문제", "PO대조_*.csv", lambda r: True)):
            for r in latest_csv(pat):
                if filt(r):
                    merged.append({
                        "문제유형": src + (f"({r['유형']})" if r.get("유형") else ""),
                        "업무ID": r.get("ID") or r.get("정산ID") or r.get("전표") or r.get("PO번호") or "",
                        "프로젝트NO": r.get("프로젝트NO") or "",
                        "기준일": (r.get("접수일자") or r.get("점검예정일") or r.get("완료일") or
                                   r.get("발행일") or r.get("일자") or ""),
                        "캠프명": r.get("캠프명", ""), "담당자": r.get("담당기사", ""),
                        "문제내용": (r.get("판정") or r.get("내용") or
                                     f"완료 {r.get('완료일','')}" ) [:100]})
    except Exception:
        pass
    from responsibility import assign_issue_row
    rows = [assign_issue_row(row) for row in merged + rows]
    rows = sort_by_date(app_year_rows(apply_rep_no(rows), "issue"), "check")
    cols = []
    for r in rows[:50]:
        for k in r:
            if k not in cols:
                cols.append(k)
    out = {"rows": rows, "cols": cols}
    return _store_cache("issues", out)


def build_id():
    """index.html이 바뀌면 값이 달라진다. 폰에 열려 있는 앱이 구버전인지 판별하는 기준."""
    try:
        st = os.stat(os.path.join(BASE, "index.html"))
        return hashlib.md5(f"{int(st.st_mtime)}-{st.st_size}".encode()).hexdigest()[:10]
    except Exception:
        return "0"


def brand_logo():
    """webapp/brand/ 에 넣어 둔 고객사 로고 파일명. 없으면 빈 문자열(기본 CSOS 마크 사용).
    파일은 gitignore 대상 — 상표 자산을 공개 저장소에 올리지 않기 위해서다."""
    d = os.path.join(BASE, "brand")
    if not os.path.isdir(d):
        return ""
    # 보고서 캡처에는 사용자가 지정한 유니버셜리프트 가로 CI 한 장을 우선한다.
    # 쿠팡 CI는 앱바에서 별도 자산으로 표시하므로 여기서 먼저 고르면 보고서의
    # 유니버셜리프트 CI가 쿠팡 파일명 정렬순서에 밀려 사라진다.
    preferred = "universal-lift-horizontal.png"
    if os.path.isfile(os.path.join(d, preferred)):
        return preferred
    for f in sorted(os.listdir(d)):
        if os.path.splitext(f)[1].lower() in (".png", ".svg", ".jpg", ".jpeg", ".webp"):
            return f
    return ""


def get_erpdocs():
    """25_ERP매출서류 — 이카운트 매출(세금)계산서 원본(2026년 전체).
    ERP는 여러 작업을 한 장으로 묶어 발행하므로 1행 = 작업 1건이 아니다."""
    if DEMO:
        return {"rows": [], "months": {}, "total": 0}
    r = _fresh("erpdocs")
    if r:
        return r
    import openpyxl
    from ecount_reconcile import load_config, resolve_master
    master = resolve_master(load_config()["reconcile"]["master_xlsx"])
    out = {"rows": [], "months": {}, "total": 0, "kinds": {}}
    try:
        wb = openpyxl.load_workbook(master, read_only=True, data_only=True)
        if "25_ERP매출서류" in wb.sheetnames:
            for row in wb["25_ERP매출서류"].iter_rows(min_row=5, values_only=True):
                if not row or not row[0]:
                    continue
                slip, mo, kind, sup = row[0], row[1], row[2], int(row[3] or 0)
                rec = {"전표": str(slip), "월": str(mo), "유형": str(kind or ""),
                       "공급가액": sup, "거래처": str(row[6] or ""),
                       "프로젝트명": str(row[7] or "")}
                if not app_year_record(rec, "erp"):
                    continue
                out["rows"].append(rec)
                m = out["months"].setdefault(str(mo), {"합계": 0, "건수": 0})
                m["합계"] += sup
                m["건수"] += 1
                m[str(kind)] = m.get(str(kind), 0) + sup
                out["kinds"][str(kind)] = out["kinds"].get(str(kind), 0) + sup
                out["total"] += sup
        # 26_계산서구성 — 계산서 1장에 어떤 프로젝트가 묶였는지(추정). 전표번호로 붙인다.
        if "26_계산서구성" in wb.sheetnames:
            comp = {}
            for row in wb["26_계산서구성"].iter_rows(min_row=5, values_only=True):
                if not row or not row[0]:
                    continue
                comp[str(row[0])] = {"포함건수": int(row[5] or 0),
                                     "포함프로젝트": str(row[6] or ""),
                                     "후보합계": int(row[7] or 0),
                                     "판정": str(row[8] or "")}
            for r in out["rows"]:
                r.update(comp.get(r["전표"], {"포함건수": 0, "포함프로젝트": "",
                                             "후보합계": 0, "판정": "미상"}))
        wb.close()
    except Exception as e:
        out["error"] = str(e)
    out["rows"] = sort_by_date(app_year_rows(out["rows"], "erp"), "erpdocs")
    return _store_cache("erpdocs", out)


def get_recalc_pending():
    """원장엔 올라왔지만 엑셀이 아직 계산하지 않아 화면에 안 나오는 건수.

    이걸 안 알려주면 사용자는 '넣었다는데 왜 없지?' 로 읽는다 — 숫자가 틀린 게 아니라
    아직 안 나온 것이다. recalc_pending.py 가 만들어 둔 캐시만 읽는다(원장 재읽기는 느리다)."""
    try:
        return json.load(open(os.path.join(ROOT, "reports", "재계산대기.json"), encoding="utf-8"))
    except Exception:
        return {"대기합계": 0, "항목": [], "안내": ""}


def get_calendar():
    """구글 캘린더(COUPANG 설치+납품+AS) 대조 캐시.

    gcal_sync.py 가 매일 만들어 둔 파일만 읽는다 — 앱은 절대 네트워크를 타지 않는다.
    폰에서 열 때 구글을 기다리면 화면이 멈추고, 터널이 죽으면 통째로 안 뜬다."""
    p = os.path.join(ROOT, "reports", "gcal_events.json")
    try:
        d = json.load(open(p, encoding="utf-8"))
    except Exception:
        return {"갱신": "", "일정": [], "원천": ["아직 수집되지 않음"], "안내":
                "구글 캘린더 설정 → 캘린더 통합 → '비공개 주소의 iCal 형식'을 config/gcal.json 에 넣어 주세요."}
    return d


def get_checks():
    """최근 카톡·밴드·ERP원장·쿠팡PO 대조 CSV를 ID별로 조인 — 4원천 검증 배지"""
    import csv as _csv
    out = {}
    def latest(pat):
        fs = sorted(glob.glob(os.path.join(ROOT, "reports", pat)))
        return fs[-1] if fs else None
    f = latest("카톡대조_*.csv")
    if f:
        for r in _csv.DictReader(open(f, encoding="utf-8-sig")):
            out.setdefault(r.get("ID", ""), {})["kakao"] = r.get("카톡보고", "")
    f = latest("밴드대조_*.csv")
    if f:
        for r in _csv.DictReader(open(f, encoding="utf-8-sig")):
            out.setdefault(r.get("ID", ""), {})["band"] = r.get("밴드게시", "")
    f = latest("ERP원장대조_*.csv")
    if f:
        for r in _csv.DictReader(open(f, encoding="utf-8-sig")):
            sid = r.get("정산ID", "") or r.get("전표", "")
            for one in str(sid).split(","):
                if one.strip():
                    out.setdefault(one.strip(), {})["erp"] = r.get("유형", "") + " " + r.get("판정", "")
    f = latest("PO대조_*.csv")
    if f:
        for r in _csv.DictReader(open(f, encoding="utf-8-sig")):
            sid = r.get("정산ID", "") or r.get("ID", "")
            for one in str(sid).split(","):
                if one.strip():
                    out.setdefault(one.strip(), {})["po"] = (r.get("판정", "") or r.get("유형", "")).strip()
    if DEMO and not out:
        out = {"JS-2607-002": {"kakao": "확인", "band": "미확인", "erp": "D 금액불일치",
                               "po": "PO 미발행"}}
    return {
        k: {a: _OLD_APP_REF_RE.sub("", str(b or "")) for a, b in v.items()}
        for k, v in out.items() if app_year_record({"ID": k})
    }


_UJ_RE = re.compile(r"(?<![A-Za-z0-9])UJ\d{6,}(?![0-9])")
# 계산서 제목에서 캠프명만 뽑는다: '쿠팡신규_송파1MB(감일동)-이동식…' → '송파1MB(감일동)'
_CAMP_RE = re.compile(r"[가-힣A-Za-z]+\d*(?:BMB|MB|캠프|Sub-?FC|Sub-?hub|FC)(?:\([^)]*\))?",
                      re.I)   # sub-hub / Sub-Hub 표기가 섞여 있어 대소문자 무시


def camp_of(title):
    m = _CAMP_RE.search(str(title or ""))
    if m:
        return m.group()
    t = re.sub(r"^(쿠팡\S*|돌발AS|정기점검)[_\s-]*", "", str(title or "")).strip()
    return (t or str(title or ""))[:28]


def erp_settlement_rows(ledger_rows):
    """관리대장에 정산 행이 **아예 없는 달**만 ERP 계산서로 채워 넣는다.

    왜 이렇게 하나
      · 06시트는 '작업 1건 = 1행' 구조다(업무구분이 원천업무ID 기반 수식).
        ERP 계산서는 여러 작업을 묶은 것이라 그 시트에 그대로 넣으면 수식이 어긋난다.
      · 그렇다고 1~6월을 비워두면 앱에서 그 달 매출이 0으로 보인다(사실과 다름).
      → 대장에 자료가 있는 달은 대장 우선, 없는 달만 ERP로 보완하고 출처를 표시한다.
        (같은 달을 양쪽에서 세지 않으므로 이중 계상이 없다)
    """
    have = {str(r.get("완료일") or "")[:7].replace("-", "/") for r in ledger_rows
            if r.get("공급가액")}
    docs = get_erpdocs()
    out = []
    for d in docs.get("rows", []):
        mo = d.get("월") or ""
        if not mo or mo in have:
            continue
        slip = d.get("전표") or ""
        iso = slip[:10].replace("/", "-")
        title = d.get("프로젝트명") or ""
        prj = (_UJ_RE.search(title) or [""])
        prj = prj.group() if hasattr(prj, "group") else ""

        sup = int(d.get("공급가액") or 0)
        out.append({
            "정산ID": "ERP-" + slip.replace("/", "").replace("-", "-"),
            "업무구분": d.get("유형") or "", "캠프명": camp_of(title),
            "프로젝트NO": prj, "원천업무ID": "",
            "공급가액": sup, "합계": sup + round(sup * 0.1),
            "명세서": "있음", "명세서번호": slip, "명세서발행일": iso,
            "계산서": "발행", "계산서발행일": iso, "승인번호": "",
            "청구일": "", "지급예정일": "", "입금일": "", "입금액": 0, "미수금": "",
            "비용구분": "유상", "PO필요": "", "PO번호": "", "PO발행일": "",
            "상태": "ERP 계산서(묶음)", "완료일": iso,
            "출처": "ERP", "적요": title})
    return out


# ── 대표 프로젝트NO ────────────────────────────────────────────
# 모든 건이 번호로 식별되게 한다. 우선순위:
#   1) 행에 이미 있는 프로젝트NO
#   2) 행 안 어딘가(내용·근거 등)에 적힌 UJ 번호
#   3) 같은 캠프·같은 달의 실제 작업에서 찾은 대표 번호
#   4) 그래도 없으면 ERP 전표번호 기반 식별자(UJ처럼 보이지 않게 'ERP-' 접두)
#      — 없는 UJ 번호를 지어내면 실제 번호와 헷갈리므로 절대 만들지 않는다.
def _camp_key(v):
    return re.sub(r"[\s()·]", "", str(v or "")).lower()[:14]


def build_prj_index(works):
    idx = {}
    for kind, dk in (("as", "접수일자"), ("pm", "점검예정일")):
        for r in works.get(kind, []):
            if r.get("출처") == "ERP":
                continue
            prj = str(r.get("프로젝트NO") or "").strip()
            if not prj:
                continue
            mo = norm_date(r.get(dk) or r.get("작업완료일") or r.get("실제점검일"))[:7]
            idx.setdefault((_camp_key(r.get("캠프명")), mo), []).append(prj)
    return idx


def rep_no(rec, idx=None, slip=""):
    """대표 프로젝트NO를 정한다(순수 함수 — 합성검증 대상)"""
    cur = str(rec.get("프로젝트NO") or "").strip()
    if cur:
        return cur, ""
    for v in rec.values():                       # 내용·근거 등 본문에 적힌 UJ
        m = _UJ_RE.search(str(v or ""))
        if m:
            return m.group(), "본문"
    if idx:
        mo = ""
        for k in ("완료일", "접수일자", "점검예정일", "일자", "작업완료일"):
            mo = norm_date(rec.get(k))[:7]
            if mo:
                break
        hits = idx.get((_camp_key(rec.get("캠프명")), mo))
        if hits:
            return sorted(hits)[0], "동일캠프·동월"
    if slip:
        m = re.match(r"(\d{2})(\d{2})/(\d{2})/(\d{2})\s*-\s*(\d+)", str(slip))
        if m:
            return f"ERP-{m.group(2)}{m.group(3)}{m.group(4)}-{m.group(5)}", "전표"
        return "ERP-" + re.sub(r"[^0-9A-Za-z-]", "", str(slip))[-10:], "전표"
    # 최후: 그 행 자신의 ID를 대표번호로 쓴다 — 번호 없는 행이 하나도 남지 않게
    for k in ("정산ID", "접수ID", "점검ID", "업무ID", "ID"):
        v = str(rec.get(k) or "").strip()
        if v:
            return v, "자체ID"
    return "", ""


def apply_rep_no(rows, idx=None, slipkey=None):
    for r in rows:
        no, how = rep_no(r, idx, str(r.get(slipkey) or "") if slipkey else "")
        if no and not str(r.get("프로젝트NO") or "").strip():
            r["프로젝트NO"] = no
            r["대표번호출처"] = how
    return rows


def erp_work_rows(existing, kind):
    """02/04 시트에 자료가 **아예 없는 달**만 ERP 계산서로 보완한다(정산과 같은 규칙).
    ERP 계산서 1장 = 작업 여러 건 묶음이므로 '건수'가 아니라 '그 달에 이런 업무가 있었다'는
    사실을 보여주는 용도다. 자료가 있는 달은 건드리지 않아 이중 계상이 없다."""
    dk = {"as": ("접수일자", "작업완료일"), "pm": ("점검예정일", "실제점검일")}[kind]
    want = {"as": "돌발AS", "pm": "정기점검"}[kind]
    have = {norm_date(r.get(dk[0]) or r.get(dk[1]))[:7] for r in existing}
    have = {h.replace("-", "/") for h in have if h}
    out = []
    for d in get_erpdocs().get("rows", []):
        if (d.get("유형") or "") != want:
            continue
        mo = d.get("월") or ""
        if not mo or mo in have:
            continue
        slip = d.get("전표") or ""
        iso = slip[:10].replace("/", "-")
        title = d.get("프로젝트명") or ""
        prj = _UJ_RE.search(title)
        base = {"프로젝트NO": prj.group() if prj else "", "캠프명": camp_of(title),
                "담당기사": "", "유상·무상·보험": "유상", "비고": "ERP 계산서 기준(작업 묶음)",
                "출처": "ERP", "적요": title}
        if kind == "as":
            base.update({"접수ID": "ERP-" + slip.replace("/", ""), "접수일자": iso,
                         "작업완료일": iso, "진행상태": "작업완료", "신청내용": title,
                         "긴급도": "", "방문예정일": ""})
        else:
            base.update({"점검ID": "ERP-" + slip.replace("/", ""), "점검예정일": iso,
                         "실제점검일": iso, "점검상태": "완료",
                         "이상발견여부": "", "돌발AS전환여부": ""})
        out.append(base)
    return out


def get_settlements():
    if DEMO:
        return demo_settlements()
    with _readlock:
        r = _fresh("settle")
        if r:
            return r
        rows = real_settlements()
        try:
            rows = rows + erp_settlement_rows(rows)
            idx = build_prj_index(get_works())
            apply_rep_no(rows, idx, "명세서번호")
            rows = app_year_rows(rows, "settle")
            rows = sort_by_date(rows, "settle", "정산ID")
        except Exception:
            pass
        return _store_cache("settle", rows)


def get_status():
    """★ 이 함수는 Z: 네트워크 드라이브를 여러 번 읽는다(원장·ERP 내보내기·대조 CSV).
    로컬 단독으로는 4초면 끝나지만, Codex·일일실행이 같은 드라이브를 쓰는 동안에는
    280초~600초까지 늘어난다. 대시보드는 이걸 주기적으로 폴링하므로 캐시가 없으면
    **앱이 열리지 않는다**(2026-07-29 실측). 다른 데이터 API와 같은 캐시를 쓴다:
    원장이 바뀌면 즉시, 아니면 120초 TTL."""
    with _readlock:
        c = _fresh("status")
        if c:
            return c
        c = _compute_status()
        if "error" not in c:
            _store_cache("status", c)
        return c


def _compute_status():
    if DEMO:
        return {"master": "쿠팡_통합업무_일일보고_관리대장_v23.xlsx (데모)", "fork": [],
                "agent_last": "2026-07-24 09:50", "steps": [
                    {"n": "합성검증", "s": "ok"}, {"n": "판매·세금계산서 대조", "s": "ok"},
                    {"n": "ERP원장 4유형 대조", "s": "ok"}, {"n": "밴드 수집·대조", "s": "skip"},
                    {"n": "카톡 대조", "s": "ok"}, {"n": "관리대장 자동입력", "s": "ok"},
                    {"n": "전표 전송대기", "s": "ok"}],
                "pending_updates": 2, "inbox": 1, "kakao": 2, "band": False, "demo": True}
    try:
        from coupang_workbench import get_status as ws
        st = ws()
        # 동기화 백본: 에이전트가 쓴 agent_status.json 우선 (없으면 md 리포트 파싱)
        steps, rt = [], ""
        try:
            aj = json.load(open(os.path.join(ROOT, "reports", "agent_status.json"), encoding="utf-8"))
            steps = aj.get("steps", [])
            rt = aj.get("time", "")[:16].replace("T", " ")
        except Exception:
            for s in st.get("report_summary", []):
                mark = "ok" if "✅" in s else ("skip" if "스킵" in s else "fail")
                steps.append({"n": re.sub(r"[✅❌⏭]|스킵|실패", "", s).strip(), "s": mark})
            rt = st.get("report_time", "")
            if rt:
                rt = f"{rt[:4]}-{rt[4:6]}-{rt[6:8]} {rt[9:11]}:{rt[11:13]}"
        tunnel = ""
        try:
            tunnel = open(os.path.join(ROOT, "reports", "tunnel_url.txt"), encoding="utf-8").read().strip()
        except Exception:
            pass
        # 3원천 검증 실집계 (최신 대조 CSV) — 보고서·대시보드가 하드코딩 없이 실데이터를 쓰도록
        srcs = {}
        try:
            from findings_export import latest_csv
            for key, pat, col, okv in (("band", "밴드대조_*.csv", "밴드게시", "확인"),
                                       ("kakao", "카톡대조_*.csv", "카톡보고", "확인")):
                rows = app_year_rows(latest_csv(pat), "issue")
                if rows:
                    ok = sum(1 for r in rows if r.get(col) == okv)
                    miss = [r for r in rows if r.get(col) != okv]
                    srcs[key] = {"total": len(rows), "ok": ok, "miss": len(miss),
                                 "miss_prj": [r.get("프로젝트NO") or r.get("ID") for r in miss[:8]]}
            erp = app_year_rows(latest_csv("ERP원장대조_*.csv"), "issue")
            if erp:
                srcs["erp"] = {"total": len(erp), "ok": 0, "miss": len(erp),
                               "miss_prj": [r.get("정산ID") or r.get("전표") for r in erp[:8]]}
            po = app_year_rows(latest_csv("PO대조_*.csv"), "issue")
            if po:
                srcs["po"] = {"total": len(po), "ok": 0, "miss": len(po),
                              "miss_prj": [r.get("PO번호") or r.get("정산ID") for r in po[:8]]}
        except Exception:
            pass

        # ★ 대조 리포트만 보면 '자료가 있는데 왜 없다고 하냐'는 말이 나온다 — 리포트가
        #   없는 이유가 (1) 파일을 안 넣었다 (2) 넣었는데 **파일이 비어 있다**
        #   (3) 대조를 안 돌렸다 로 갈리기 때문이다. 2026-07-27에 실제로 ERP 파일 3개가
        #   회사명 한 줄만 있는 빈 파일이었고, 아무도 그걸 몰랐다. 그래서 여기서 같이 본다.
        try:
            from inbox_scan import pick
            import openpyxl as _ox
            for key, kinds in (("erp", ("ledger", "slips")), ("po", ("po",)),
                               ("tax", ("tax",)), ("stmt", ("stmt",))):
                files = []
                for kd in kinds:
                    files += pick(kd) or []
                files = [f for f in files
                         if not re.search(r"(?<!\d)2025(?!\d)|(?<!\d)25[/._-]\d{2}", os.path.basename(f))]
                info = {"files": len(files), "rows": 0, "empty": []}
                for f in files[:6]:
                    try:
                        w = _ox.load_workbook(f, read_only=True, data_only=True)
                        n = 0
                        for sn in w.sheetnames:
                            n += sum(1 for r in w[sn].iter_rows(values_only=True)
                                     if sum(1 for x in r if x not in (None, "")) >= 3)
                        w.close()
                        info["rows"] += n
                        if n < 2:                      # 회사명 한 줄만 있는 '빈 내보내기'
                            info["empty"].append(os.path.basename(f))
                    except Exception:
                        pass
                if files:
                    srcs.setdefault(key, {})["inbox"] = info
        except Exception:
            pass

        try:
            from agent_dispatch import status as agent_dispatch_status
            agent_route = agent_dispatch_status()
        except Exception:
            agent_route = {}
        return {"master": os.path.basename(st.get("master", "") or "") + "  " + st.get("master_label", ""),
                "fork": st.get("fork", []), "agent_last": rt or "기록 없음", "steps": steps,
                "pending_updates": st["pending_updates"], "inbox": st["inbox"],
                "kakao": st["kakao"], "band": st["band_auth"], "demo": False, "tunnel": tunnel,
                "sources": srcs, "build": build_id(), "recalc": get_recalc_pending(),
                "agent_dispatch": agent_route}
    except Exception as e:
        return {"error": str(e)}


def latest_reports():
    out = []
    old_year = _OLD_APP_REF_RE
    # '자료현황' 을 맨 앞에 둔다 — "그거 지금 몇 건이지?" 를 매번 다시 세지 않으려고 만든 장이다
    # (사용자 지시 2026-07-29). data_status.py 가 만들고 daily_run 이 매일 갱신한다.
    for pat, name in [("자료현황.md", "자료현황"),
                      ("종합리포트_*.md", "종합"), ("카톡대조_*.md", "카톡"), ("밴드대조_*.md", "밴드"),
                      ("ERP원장대조_*.md", "ERP원장"), ("이카운트대조_*.md", "판매·계산서")]:
        fs = sorted(glob.glob(os.path.join(ROOT, "reports", pat)))
        if fs:
            text = open(fs[-1], encoding="utf-8").read()[:20000]
            text = "\n".join(line for line in text.splitlines() if not old_year.search(line))
            out.append({"kind": name, "file": os.path.basename(fs[-1]),
                        "text": text})
    if DEMO and not out:
        out = [{"kind": "종합", "file": "demo.md",
                "text": "# 데모 리포트\n\n| 단계 | 결과 |\n|---|---|\n| 합성검증 | ✅ |\n| 카톡 대조 | ✅ |\n\n## 문제 예시\n- JS-2607-002 금액불일치(원장 620,000 / EC 500,000)"}]
    return out


# ───────────────────────── HTTP ─────────────────────────
class H(BaseHTTPRequestHandler):
    def handle_one_request(self):
        # 어떤 예외도 소켓을 조용히 죽이지 않게(ERR_EMPTY_RESPONSE 방지) 전역 가드
        try:
            super().handle_one_request()
        except (ConnectionError, BrokenPipeError):
            pass
        except Exception as e:
            try:
                self._send(500, {"error": str(e)[:300]})
            except Exception:
                pass

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        data = body if isinstance(body, bytes) else json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Private-Network", "true")
        self.end_headers()
        self.wfile.write(data)

    def _auth(self):
        # 잠금 카운트는 /api/login 에서만 증가 — 구 PIN이 저장된 브라우저의
        # 자동 폴링이 잠금을 유발하던 문제(자기 잠금) 방지
        if _locked(self.client_address[0]):
            return False
        return self.headers.get("X-Pin", "") == PIN

    def do_OPTIONS(self):
        """브라우저 수집기(band.us 페이지)에서 보내는 사전 요청 허용 — 로컬에서만 쓴다"""
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Pin")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        # Chrome의 Private Network Access: 공개 사이트(https)에서 로컬 주소로 보낼 때 필요
        self.send_header("Access-Control-Allow-Private-Network", "true")
        self.send_header("Access-Control-Max-Age", "600")
        self.end_headers()

    def do_GET(self):
        p = self.path.split("?")[0]
        if p in ("/", "/index.html", "/ryu"):
            html = open(os.path.join(BASE, "index.html"), encoding="utf-8").read()
            # ★★ 터널 주소로 들어온 경우에는 **설치 가능하게 만들지 않는다**.
            #   터널 호스트는 띄울 때마다 바뀌는데, 여기서 [설치]를 하면 그 임시 호스트가
            #   앱 아이콘에 **영구히 박힌다**. 주소가 바뀌는 순간 그 아이콘은 영영
            #   'ERR_FAILED'만 띄운다 — PC·폰 둘 다 그렇게 죽었다(2026-07-28 실사고).
            #   설치는 오직 고정 주소에서만 되게 하고, 여기서는 매니페스트·서비스워커를 뺀다.
            host = (self.headers.get("Host") or "").lower()
            if "trycloudflare.com" in host:
                html = html.replace('<link rel="manifest" href="/manifest.json">', "")
                html = html.replace("navigator.serviceWorker.register('/sw.js')",
                                    "Promise.reject()")
            return self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
        if p.startswith("/brand/"):                    # 고객사 CI(쿠팡 로고) — 로컬 파일만 서빙
            fn = os.path.basename(p)
            fp = os.path.join(BASE, "brand", fn)
            ext = os.path.splitext(fn)[1].lower()
            ct = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
                  ".svg": "image/svg+xml", ".webp": "image/webp"}.get(ext)
            if not ct or not os.path.exists(fp):
                return self._send(404, {"error": "no brand asset"})
            return self._send(200, open(fp, "rb").read(), ct)
        if p.startswith("/icons/"):
            fn = os.path.basename(p)
            fp = os.path.join(BASE, "icons", fn)
            if not fn.endswith(".svg") or not os.path.isfile(fp):
                return self._send(404, {"error": "no icon asset"})
            return self._send(200, open(fp, "rb").read(), "image/svg+xml")
        if p == "/api/brief":
            # 대표 보고용 '내용' 브리핑. 화면·PC 리포트·폰 사본이 **같은 문장**을 쓰도록
            # daily_brief 하나만 출처로 삼는다(따로 만들면 숫자가 갈린다).
            try:
                import daily_brief as DB
                day = None
                m = re.search(r"[?&]date=(\d{4}-\d{2}-\d{2})", self.path)
                if m:
                    day = m.group(1)
                    if not day.startswith(APP_YEAR + "-"):
                        return self._send(200, {"ok": False,
                                               "error": f"앱 브리핑은 {APP_YEAR}년만 표시합니다"})
                b = get_daily_brief(day)
                return self._send(200, {"ok": True, "text": DB.text(b), **b})
            except Exception as e:
                return self._send(200, {"ok": False, "error": str(e)[:200]})
        if p == "/api/codes":
            return self._send(200, get_codes())
        if p == "/api/brand":
            return self._send(200, {"logo": brand_logo()})
        if re.fullmatch(r"/icon(?:-\d+)?\.(svg|png)", p):      # 아이콘(벡터/래스터 공용)
            try:
                return self._send(200, open(os.path.join(BASE, p.lstrip("/")), "rb").read(),
                                  "image/svg+xml" if p.endswith(".svg") else "image/png")
            except Exception:
                return self._send(404, {"error": "no icon"})
        if p == "/sw.js":
            # 크롬이 [설치 및 바로가기 만들기]로 진짜 앱 설치를 해 주려면
            # fetch 핸들러를 가진 서비스 워커가 필요하다. 캐시는 하지 않는다 —
            # 이 앱은 매일 바뀌는 실데이터를 보여주므로 옛 화면이 남으면 안 된다.
            js = ("self.addEventListener('install',e=>self.skipWaiting());\n"
                  "self.addEventListener('activate',e=>e.waitUntil(self.clients.claim()));\n"
                  "self.addEventListener('fetch',e=>{e.respondWith(fetch(e.request));});\n")
            # ★ _send 는 str을 받으면 JSON으로 감싼다 — 반드시 bytes로 넘겨야 스크립트가 된다
            return self._send(200, js.encode("utf-8"), "application/javascript; charset=utf-8")
        if p == "/manifest.json":                      # 홈 화면에 추가 시 앱처럼 보이게
            # ★ start_url 은 **이 페이지와 같은 출처**여야 한다. 다른 도메인을 넣으면
            #   크롬이 매니페스트를 통째로 무시해 [설치 및 바로가기 만들기]가 먹통이 된다
            #   (2026-07-27에 고정 주소를 넣었다가 실제로 설치가 안 됐다).
            #   그래서 여기는 "/" 로 두고, **오래 쓸 아이콘은 고정 주소에서 설치**한다.
            #   터널 주소로 들어온 사람에게는 index.html이 배너로 그 사실을 알린다.
            return self._send(200, {
                "name": "Coupang Service Operations System", "short_name": "CSOS",
                "start_url": "/", "scope": "/", "display": "standalone",
                "background_color": "#060D2B", "theme_color": "#060D2B",
                "icons": [
                    {"src": "/icon-192.png?v=csos-20260729", "sizes": "192x192",
                     "type": "image/png", "purpose": "any"},
                    {"src": "/icon-512.png?v=csos-20260729", "sizes": "512x512",
                     "type": "image/png", "purpose": "any maskable"}]},
                "application/manifest+json")
        if p == "/api/ping":
            return self._send(200, {"app": "coupang-work", "demo": DEMO, "build": build_id()})
        if not self._auth():
            return self._send(401, {"error": "PIN"})
        if p == "/api/status":
            return self._send(200, get_status())
        # 철거·신규납품은 응답 단계에서 뺀다 — 앱이 아예 받지 않게 한다.
        # (앱에도 같은 필터가 있지만, 화면 코드가 바뀌어도 새어 나가지 않게 서버가 먼저 막는다)
        if p == "/api/settlements":
            return self._send(200, {"rows": drop_side_work(get_settlements())})
        if p == "/api/works":
            w = get_works()
            return self._send(200, {k: (drop_side_work(v) if isinstance(v, list) else v)
                                    for k, v in (w or {}).items()})
        if p == "/api/issues":
            iss = get_issues()
            if isinstance(iss, dict) and isinstance(iss.get("rows"), list):
                iss = {**iss, "rows": drop_side_work(iss["rows"])}
            return self._send(200, iss)
        if p == "/api/ryu/records":
            return self._send(200, get_ryu_records())
        if p == "/api/exec_report":
            m = re.search(r"[?&]date=(\d{4}-\d{2}-\d{2})", self.path)
            day = m.group(1) if m else None
            if day and not day.startswith(APP_YEAR + "-"):
                return self._send(400, {"error": f"{APP_YEAR}년 날짜만 선택할 수 있습니다"})
            return self._send(200, get_exec_report(day))
        if p in {
            "/api/v1/reports/daily/exceptions",
            "/api/v1/reports/daily/as-backlog",
            "/api/v1/reports/daily/inspection-progress",
            "/api/v1/reports/daily/statement-progress",
            "/api/v1/as-requests/backlog-summary",
            "/api/v1/as-requests/backlog-detail",
            "/api/v1/inspections/quarter-progress",
            "/api/v1/statements/eligibility-summary",
            "/api/v1/statements/unissued",
            "/api/v1/tax-invoices/composition-check",
        }:
            report = get_representative_report()
            if p.endswith(("as-backlog", "backlog-summary", "backlog-detail")):
                return self._send(200, {"meta": report["meta"], **report["돌발AS"]})
            if p.endswith(("inspection-progress", "quarter-progress")):
                return self._send(200, {"meta": report["meta"], **report["정기점검"]})
            if p.endswith(("statement-progress", "eligibility-summary", "unissued",
                           "composition-check")):
                return self._send(200, {"meta": report["meta"], **report["거래명세서"],
                                        "업무기준확인필요": report["업무기준확인필요"]})
            return self._send(200, report)
        if p == "/api/calendar":
            return self._send(200, get_calendar())
        if p == "/api/erpdocs":
            return self._send(200, get_erpdocs())
        if p == "/api/checks":
            return self._send(200, get_checks())
        if p == "/api/reports":
            return self._send(200, {"reports": latest_reports()})
        if p == "/api/tasklog":
            return self._send(200, {"busy": runner["busy"], "task": runner["task"],
                                    "log": list(runner["log"])[-300:]})
        return self._send(404, {"error": "not found"})

    def do_POST(self):
        p = self.path.split("?")[0]
        ip = self.client_address[0]
        if p == "/api/login":
            if _locked(ip):
                return self._send(429, {"ok": False, "error": "시도 초과 — 10분 후 다시"})
            ln = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(ln) or b"{}")
            ok = body.get("pin", "") == PIN
            (_ok_login if ok else _fail)(ip)
            return self._send(200 if ok else 401, {"ok": ok})
        if p == "/api/band_dump":
            return self._band_dump()
        if not self._auth():
            return self._send(401, {"error": "PIN"})
        if p == "/api/export_xlsx":
            ln = int(self.headers.get("Content-Length", 0))
            if ln <= 0 or ln > 2_000_000:
                return self._send(400, {"ok": False, "error": "내보낼 목록 크기가 올바르지 않습니다"})
            try:
                payload = json.loads(self.rfile.read(ln) or b"{}")
                data, _title = rows_xlsx(payload)
                return self._send(
                    200, data,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                return self._send(500, {"ok": False, "error": f"엑셀 생성 실패: {str(e)[:160]}"})
        if p == "/api/policy":
            ln = int(self.headers.get("Content-Length", 0))
            if ln <= 0 or ln > 100_000:
                return self._send(400, {"ok": False, "error": "저장 내용 크기가 올바르지 않습니다"})
            b = json.loads(self.rfile.read(ln) or b"{}")
            key = str(b.get("기준") or "").strip()
            value = str(b.get("확정내용") or "").strip()
            if not key or not value:
                return self._send(400, {"ok": False, "error": "기준과 확정 내용을 입력하세요"})
            state = {
                "상태": "확정", "확정내용": value,
                "저장자": str(b.get("저장자") or "앱 사용자")[:40],
                "저장일시": datetime.now().isoformat(timespec="seconds"),
            }
            save_policy_state(key, state)
            return self._send(200, {"ok": True, **state})
        if p == "/api/ryu/upload":
            ln = int(self.headers.get("Content-Length", 0))
            if ln <= 0 or ln > 45_000_000:
                return self._send(400, {"ok": False, "error": "첨부 용량은 합계 45MB 이하여야 합니다"})
            try:
                fields, files = multipart_parts(self.headers.get("Content-Type", ""),
                                                self.rfile.read(ln))
                result = save_ryu_upload(fields, files)
                started, msg = start_task("kakao")
                queued = False if started else defer_task_until_free("kakao")
                return self._send(200, {"ok": True, "saved": result,
                                        "auto_check_started": started,
                                        "auto_check_queued": queued, "msg": msg})
            except Exception as e:
                return self._send(400, {"ok": False, "error": str(e)[:260]})
        if p == "/api/ryu/entry":
            ln = int(self.headers.get("Content-Length", 0))
            if ln <= 0 or ln > 30_000_000:
                return self._send(400, {"ok": False, "error": "입력·첨부 용량은 합계 30MB 이하여야 합니다"})
            try:
                fields, files = multipart_parts(self.headers.get("Content-Type", ""),
                                                self.rfile.read(ln))
                result = save_ryu_entry(fields, files, ip)
                return self._send(200, {"ok": True, **result})
            except Exception as e:
                return self._send(400, {"ok": False, "error": str(e)[:300]})
        if p == "/api/enqueue":
            # 폰이 **PC 꺼진 동안 예약해 둔** 프로젝트 코드를 받아 원장에 등록한다.
            # 오프라인 앱이 PC가 살아난 걸 확인하는 즉시 스스로 보낸다(사람 개입 없음).
            ln = int(self.headers.get("Content-Length", 0))
            codes = (json.loads(self.rfile.read(ln) or b"{}").get("codes") or [])[:50]
            if DEMO:
                return self._send(200, {"ok": True, "applied": 0, "msg": "데모"})
            return self._send(200, enqueue_codes(codes))
        m = re.match(r"^/api/run/(\w+)$", p)
        if m:
            ok, msg = start_task(m.group(1))
            return self._send(200 if ok else 409, {"ok": ok, "msg": msg})
        if p == "/api/open":
            # 워크벤치 대체: 관리대장·폴더 열기. 원격(터널)에서는 의미가 없고 위험하므로
            # 서버가 도는 PC에서 접속했을 때만 허용한다.
            if ip not in ("127.0.0.1", "::1", "localhost"):
                return self._send(403, {"ok": False, "error": "이 기능은 사무실 PC에서만 됩니다"})
            ln = int(self.headers.get("Content-Length", 0))
            what = json.loads(self.rfile.read(ln) or b"{}").get("what", "")
            try:
                from ecount_reconcile import load_config, resolve_master
                master = resolve_master(load_config()["reconcile"]["master_xlsx"])
            except Exception:
                master = ""
            targets = {"master": master, "master_dir": os.path.dirname(master),
                       "inbox": os.path.join(ROOT, "inbox"),
                       "kakao": os.path.join(ROOT, "kakao", "inbox"),
                       "band_docs": os.path.join(ROOT, "band", "docs_inbox"),
                       "band_cache": os.path.join(ROOT, "band", "cache"),
                       "reports": os.path.join(ROOT, "reports")}
            path = targets.get(what)
            if not path or not os.path.exists(path):
                return self._send(404, {"ok": False, "error": f"경로 없음: {what}"})
            try:
                os.startfile(path)
                return self._send(200, {"ok": True, "opened": os.path.basename(path) or path})
            except Exception as e:
                return self._send(500, {"ok": False, "error": str(e)[:200]})
        if p == "/api/set_dates":
            # 보고일·집계기준일 → 00_대시보드 B3·B4 (일일 갱신 입력칸 — 덮어쓰기 허용 화이트리스트)
            ln = int(self.headers.get("Content-Length", 0))
            b = json.loads(self.rfile.read(ln) or b"{}")
            items = []
            for cell, key, label in (("B3", "보고일", "보고일"), ("B4", "집계기준일", "집계기준일")):
                v = str(b.get(key, "")).strip()
                if v:
                    if not re.match(r"^\d{4}-\d{2}-\d{2}$", v):
                        return self._send(400, {"ok": False, "error": f"{label} 형식 오류(YYYY-MM-DD)"})
                    items.append({"sheet": "00_대시보드", "cell": cell, "key": cell, "key_col": "-",
                                  "col": label, "value": v, "vtype": "date",
                                  "evidence": f"앱 기준일 설정({ip})", "only_if_empty": False})
            if not items:
                return self._send(400, {"ok": False, "error": "날짜 없음"})
            if DEMO:
                return self._send(200, {"ok": True, "demo": True})
            from ledger_writer import queue_add
            queue_add(items)
            ok, msg = start_task("writer_apply")     # 즉시 반영(vN+1)
            return self._send(200, {"ok": True, "applying": ok, "msg": msg})
        if p == "/api/input":
            # 앱 → 엑셀 입력: ledger_writer 큐에 적재(빈 칸만 정책은 반영 단계에서 강제)
            ln = int(self.headers.get("Content-Length", 0))
            b = json.loads(self.rfile.read(ln) or b"{}")
            ALLOW = {"02_돌발AS접수", "04_정기점검", "06_거래서류청구수금",
                     "15_세금계산서관리", "16_입금수금관리"}
            if b.get("sheet") not in ALLOW or not b.get("key") or not b.get("col") or b.get("value") in (None, ""):
                return self._send(400, {"ok": False, "error": "sheet/key/col/value 필요"})
            if b.get("vtype") not in ("text", "date", "number"):
                b["vtype"] = "text"
            if DEMO:
                return self._send(200, {"ok": True, "queued": 1, "demo": True})
            from ledger_writer import queue_add, load_queue
            n = queue_add([{"sheet": b["sheet"], "key_col": b.get("key_col", "정산ID"), "key": b["key"],
                            "col": b["col"], "value": b["value"], "vtype": b["vtype"],
                            "evidence": f"앱 입력({ip}) {datetime.now():%m-%d %H:%M}", "only_if_empty": True}])
            return self._send(200, {"ok": True, "queued": n, "pending": len(load_queue())})
        return self._send(404, {"error": "not found"})

    def _band_dump(self):
        """브라우저 수집기가 밴드 게시글 원본을 직접 전송(no-cors POST) — PIN 쿼리로 보호"""
        from urllib.parse import parse_qs, urlparse
        q = parse_qs(urlparse(self.path).query)
        if (q.get("pin") or [""])[0] != PIN:
            return self._send(401, {"ok": False})
        ln = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(ln)
        try:
            d = json.loads(raw.decode("utf-8"))
            band = re.sub(r"\D", "", str(d.get("band", ""))) or "unknown"
            os.makedirs(os.path.join(ROOT, "band", "cache"), exist_ok=True)
            path = os.path.join(ROOT, "band", "cache", f"dump_{band}.json")
            open(path, "w", encoding="utf-8").write(json.dumps(d, ensure_ascii=False))
            return self._send(200, {"ok": True, "saved": len(d.get("posts", {}))})
        except Exception as e:
            return self._send(400, {"ok": False, "error": str(e)[:200]})

    def log_message(self, *a):
        pass


def lan_ip():
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


PUBLISH_EVERY = 3 * 3600      # 폰이 보는 사본을 몇 초마다 새로 올릴지


def publish_loop():
    """PC가 켜져 있는 동안 **주기적으로** 폰용 사본을 올린다.

    ★ PC가 꺼져도 폰이 쓰이려면 사본이 최신이어야 한다. 예전에는 daily_run 이 돌 때만
      올려서, 아침에 한 번 돌리고 저녁에 PC를 끄면 폰은 **아침 숫자**를 보게 됐다.
      그러면 '꺼져도 된다'는 말이 사실이 아니게 된다. 그래서 3시간마다 올린다.
      (사본은 잠겨 있고 60~80KB라 부담이 없다)
    """
    if DEMO:
        return
    time.sleep(120)                        # 기동 직후 혼잡할 때는 피한다
    while True:
        from operation_window import is_input_window
        if is_input_window():
            time.sleep(60)
            continue
        try:
            # 자동 게시도 사람·다른 AI의 수동 게시를 밟지 않게 publish 점유를 강제한다.
            publish_env = {**ENV, "CSOS_AI": "server"}
            r = subprocess.run([PY, os.path.join(ROOT, "cloud_publish.py"), "--push"],
                               cwd=ROOT, env=publish_env, capture_output=True, text=True,
                               encoding="utf-8", errors="replace", timeout=900)
            tail = [l for l in (r.stdout or "").splitlines() if l.strip()][-1:] or [""]
            runner["log"].append(f"[사본 자동 게시] {tail[0][:120]}")
        except Exception as e:
            runner["log"].append(f"[사본 자동 게시] 실패 {type(e).__name__}")
        time.sleep(PUBLISH_EVERY)


class _Server(ThreadingHTTPServer):
    """★ 사고 #16의 진짜 원인 — 윈도우에서 SO_REUSEADDR 이 켜져 있으면 **이미 남이 쓰는
    포트에도 바인드가 성공한다.** 그래서 새 서버가 '시작됨'을 찍고도 요청은 계속 옛
    프로세스가 받아 갔다. 코드를 고쳐도 화면이 안 바뀌는 증상의 정체가 이것이다.
    재사용을 끄면 두 번째 서버는 조용히 뜨는 대신 **에러로 죽는다** — 그게 옳다."""
    allow_reuse_address = False


def main():
    try:
        srv = _Server(("0.0.0.0", PORT), H)
    except OSError:
        print(f"★ {PORT} 포트를 이미 다른 앱 서버가 쓰고 있습니다. 새로 뜨지 않았습니다.")
        print("  옛 서버가 계속 응답하므로 **코드를 고쳐도 화면이 안 바뀝니다.**")
        print("  먼저 정리하세요 (PowerShell):")
        print("    Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" |"
              " ? { $_.CommandLine -like '*app_server.py*' } | % { Stop-Process -Id $_.ProcessId -Force }")
        return 1
    threading.Thread(target=publish_loop, daemon=True).start()
    mode = "데모(합성데이터)" if DEMO else "실서비스"
    print(f"Coupang Service Operations System 앱 서버 [{mode}] 시작")
    print(f"  PC:      http://localhost:{PORT}")
    print(f"  휴대폰:  http://{lan_ip()}:{PORT}   (같은 와이파이)")
    print(f"  PIN:     {PIN}")
    srv.serve_forever()


if __name__ == "__main__":
    main()
