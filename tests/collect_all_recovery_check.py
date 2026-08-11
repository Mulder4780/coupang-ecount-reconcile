# -*- coding: utf-8 -*-
"""[216] 미수집 보관 자율복구 회귀검증 — 실데이터·네트워크 접촉 0."""
import json
import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.environ["CSOS_SYNTHETIC"] = "1"

import collect_all as C
from band import archive_posts as A


def main():
    with tempfile.TemporaryDirectory(prefix="csos-collect-recovery-216-") as td:
        cache = os.path.join(td, "cache")
        band_root = os.path.join(td, "band-root")
        os.makedirs(cache)

        post = {
            "created_at": 1767225600000,
            "content": "● 프로젝트NO : UJ2600001\n돌발 AS 완료",
            "images": ["https://example.invalid/one", "https://example.invalid/two"],
        }
        with open(os.path.join(cache, "90610953.json"), "w", encoding="utf-8") as fh:
            json.dump({"band_name": "합성 밴드", "posts": {"10": post}}, fh,
                      ensure_ascii=False)

        archive_root = os.path.join(band_root, "게시글보관")
        set_root = os.path.join(archive_root, A.safe("합성 밴드", 30))
        paths = A.archive_paths(set_root, "10", post)
        os.makedirs(paths["folder"], exist_ok=True)
        for path in (paths["pdf"], paths["txt"], paths["photos"][0]):
            with open(path, "wb") as fh:
                fh.write(b"synthetic")

        # 예전 개정본·다른 사진은 현재 캐시의 보관 완료 수에 섞이면 안 된다.
        with open(os.path.join(archive_root, "stray-old.pdf"), "wb") as fh:
            fh.write(b"old")
        with open(os.path.join(archive_root, "stray-old.jpg"), "wb") as fh:
            fh.write(b"old")
        survey = C.survey(cache_dir=cache, band_root=band_root)
        assert survey["밴드글_캐시"] == survey["밴드글_보관"] == 1, survey
        assert survey["텍스트_보관"] == 1, survey
        assert survey["사진_URL"] == 2 and survey["사진_보관"] == 1, survey
        assert survey["사진_남음"] == 1, survey

        # PDF·텍스트가 이미 있어도 빠진 두 번째 사진을 다시 받는다.
        old_root, old_fetch, old_render = A.out_root, A.fetch_photo, A.render_pdf
        try:
            A.out_root = lambda: archive_root

            def fake_fetch(_url, path):
                if os.path.exists(path):
                    return "skip"
                with open(path, "wb") as fh:
                    fh.write(b"recovered")
                return "ok"

            A.fetch_photo = fake_fetch
            A.render_pdf = lambda _no, _post, _name, _photos, path: path
            stat = {"made": 0, "skip": 0, "photo": 0, "pdf_fail": 0}
            A.archive_band("90610953", {"10": post, "_band_name": "합성 밴드"},
                           limit=1, force=False, stat=stat)
        finally:
            A.out_root, A.fetch_photo, A.render_pdf = old_root, old_fetch, old_render
        assert os.path.exists(paths["photos"][1]) and stat["photo"] == 1, stat

    collect_src = open(os.path.join(ROOT, "collect_all.py"), encoding="utf-8").read()
    autopilot_src = open(os.path.join(ROOT, "autopilot.py"), encoding="utf-8").read()
    daily_src = open(os.path.join(ROOT, "daily_run.py"), encoding="utf-8").read()
    assert C.DEFAULT_BUDGET_SECONDS <= 7 * 60
    assert C.CONTINUE_RETURN_CODE == 75
    assert "run_tree(" in collect_src and "subprocess.run(" not in collect_src
    assert "INCREMENTAL_RETURN_CODE = 75" in autopilot_src
    assert '"status": "waiting"' in autopilot_src and '"continuations"' in autopilot_src
    assert "INCREMENTAL_RETURN_CODE = 75" in daily_src
    assert 'got.get("returncode") == INCREMENTAL_RETURN_CODE' in daily_src
    assert 'local_rc = 124 if item.get("kind") == "timeout"' in autopilot_src
    assert 'item["last_returncode"]' in autopilot_src
    print("[216/217] 현재 세트 집계 · 누락 사진 재시도 · 7분 자진복귀 · 증분 대기 ✅")


if __name__ == "__main__":
    main()
