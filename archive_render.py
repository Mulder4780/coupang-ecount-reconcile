# -*- coding: utf-8 -*-
"""
archive_render.py — HTML → PDF 렌더러 (크롬 헤드리스)

왜: 이 프로젝트에는 PDF 라이브러리가 없고(설치도 안 한다), 엑셀 COM 은 엑셀 파일에만 쓴다.
    밴드 글·거래명세서 건별 보관본은 **사람이 보던 그대로**를 고정해야 하므로
    이미 설치된 크롬을 `--headless --print-to-pdf` 로 쓴다(추가 설치·네트워크 없음).

  from archive_render import html_to_pdf
  html_to_pdf(html_text, "out.pdf")

동시에 여러 개를 만들면 크롬이 프로필을 잠근다 — 호출마다 임시 프로필을 준다.
"""
import os
import subprocess
import sys
import tempfile

CHROME = next((p for p in (
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
) if os.path.exists(p)), None)

CSS = """
<meta charset="utf-8">
<style>
 @page { size: A4; margin: 12mm; }
 body { font-family: 'Malgun Gothic', sans-serif; font-size: 11pt; color:#111; }
 h1 { font-size: 15pt; margin: 0 0 4mm; border-bottom: 2px solid #333; padding-bottom: 2mm; }
 .meta { font-size: 9.5pt; color:#444; margin-bottom: 4mm; }
 .meta b { color:#000; }
 pre { white-space: pre-wrap; word-break: break-all; font-family: inherit;
       background:#fafafa; border:1px solid #ddd; padding:3mm; }
 table { border-collapse: collapse; width:100%; font-size:10pt; }
 th,td { border:1px solid #bbb; padding:1.5mm 2mm; }
 th { background:#f0f0f0; }
 td.num { text-align:right; }
 .photos img { max-width: 88mm; margin:2mm; border:1px solid #ccc; }
 .foot { margin-top:5mm; font-size:8.5pt; color:#666; border-top:1px solid #ddd; padding-top:2mm; }
</style>
"""


def html_to_pdf(html, out_pdf, timeout=90):
    """HTML 문자열을 PDF 로 굳힌다. 성공하면 out_pdf 경로, 실패하면 None."""
    if not CHROME:
        return None
    os.makedirs(os.path.dirname(out_pdf) or ".", exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, "page.html")
        with open(src, "w", encoding="utf-8") as f:
            f.write(CSS + html)
        cmd = [CHROME, "--headless=new", "--disable-gpu", "--no-sandbox",
               "--no-pdf-header-footer", f"--user-data-dir={os.path.join(td, 'prof')}",
               f"--print-to-pdf={out_pdf}", "file:///" + src.replace("\\", "/")]
        try:
            subprocess.run(cmd, capture_output=True, timeout=timeout, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except Exception:
            return None
    return out_pdf if os.path.exists(out_pdf) else None


def esc(s):
    return (str(s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


if __name__ == "__main__":
    ok = html_to_pdf("<h1>렌더 시험</h1><p>크롬 헤드리스 PDF 확인</p>",
                     os.path.join(tempfile.gettempdir(), "csos_render_test.pdf"))
    print("PDF 생성", "성공" if ok else "실패", ok or "")
    sys.exit(0 if ok else 1)
