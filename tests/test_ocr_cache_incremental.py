# -*- coding: utf-8 -*-
"""OCR 캐시 집중 회귀검사 — 실 Z:·실 OCR·관리대장을 건드리지 않는다."""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
from unittest import mock


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BAND = os.path.join(ROOT, "band")
for _p in (ROOT, BAND):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import doc_ocr as D
import ocr_crosscheck as X


def _make_images(folder: str, names: tuple[str, ...]) -> list[str]:
    out = []
    for i, name in enumerate(names):
        path = os.path.join(folder, "nested", name)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as fh:
            fh.write(("image-%d" % i).encode("ascii"))
        os.utime(path, (1_700_000_000 + i, 1_700_000_000 + i))
        out.append(path)
    return out


def test_stat_is_reused_and_only_new_two_are_read() -> None:
    td = tempfile.mkdtemp(prefix="ocr205_")
    real_cache, real_cross = D.OCR_CACHE, X.XCACHE
    real_batch, real_run = D.OCR_BATCH_SIZE, D._paddle_batch
    try:
        source = os.path.join(td, "source")
        paths = _make_images(source, ("done.jpg", "new1.jpg", "new2.jpg"))
        imgs, stats = D.image_manifest(source)
        assert imgs == paths and set(stats) == set(paths)

        # 목록에서 받은 stat을 넘기면 cache key가 원본 경로를 다시 찌르지 않는다.
        with mock.patch.object(os, "stat", side_effect=AssertionError("path stat 재호출")):
            D._cache_path(paths[0], stats[paths[0]])
            X._cache_path("windows", paths[0], stats[paths[0]])

        D.OCR_CACHE = os.path.join(td, "cache")
        X.XCACHE = os.path.join(D.OCR_CACHE, "cross")
        D._write_cache(paths[0], "already-done", stats[paths[0]])

        calls: list[list[str]] = []

        def fake_paddle(batch, _timeout):
            calls.append(list(batch))
            return {p: "fresh:" + os.path.basename(p) for p in batch}

        D._paddle_batch = fake_paddle
        D.OCR_BATCH_SIZE = 8
        with mock.patch.object(D, "_ocr_run", side_effect=AssertionError("Paddle 결과 폴백")):
            first = D.ocr_images(imgs, stats=stats)
            assert calls == [[paths[1], paths[2]]], calls
            assert first[paths[0]] == "already-done"
            assert first[paths[1]].startswith("fresh:") and first[paths[2]].startswith("fresh:")

            # 같은 세 장을 다시 보면 완료 OCR은 0장이다.
            calls.clear()
            second = D.ocr_images(imgs, stats=stats)
            assert calls == [], "완료 사진을 다시 OCR 했다: %r" % calls
            assert second == first
    finally:
        D.OCR_CACHE, X.XCACHE = real_cache, real_cross
        D.OCR_BATCH_SIZE, D._paddle_batch = real_batch, real_run
        shutil.rmtree(td, ignore_errors=True)


def test_interruption_keeps_the_last_completed_image() -> None:
    td = tempfile.mkdtemp(prefix="ocr205_cut_")
    real_cache, real_batch, real_run = D.OCR_CACHE, D.OCR_BATCH_SIZE, D._paddle_batch
    try:
        paths = _make_images(td, ("first.jpg", "second.jpg"))
        stats = {p: os.stat(p) for p in paths}
        D.OCR_CACHE = os.path.join(td, "cache")
        D.OCR_BATCH_SIZE = 1
        attempts: list[str] = []

        def interrupted(batch, _timeout):
            p = batch[0]
            attempts.append(p)
            if p == paths[1]:
                raise KeyboardInterrupt("부모 회차 종료 흉내")
            return {p: "saved-before-cut"}

        D._paddle_batch = interrupted
        try:
            D.ocr_images(paths, stats=stats)
            raise AssertionError("중단 흉내가 전달되지 않았다")
        except KeyboardInterrupt:
            pass
        assert D._read_cache(paths[0], stats[paths[0]]) == "saved-before-cut"
        assert D._read_cache(paths[1], stats[paths[1]]) is None

        # 재실행은 이미 굳힌 첫 장을 건너뛰고 남은 한 장만 처리한다.
        attempts.clear()
        D._paddle_batch = lambda batch, _timeout: (
            attempts.extend(batch) or {p: "resumed" for p in batch})
        got = D.ocr_images(paths, stats=stats)
        assert attempts == [paths[1]], attempts
        assert got[paths[0]] == "saved-before-cut" and got[paths[1]] == "resumed"
    finally:
        D.OCR_CACHE = real_cache
        D.OCR_BATCH_SIZE, D._paddle_batch = real_batch, real_run
        shutil.rmtree(td, ignore_errors=True)


def test_crosscheck_engine_reuses_completed_cache() -> None:
    td = tempfile.mkdtemp(prefix="ocr205_cross_")
    real_cross, real_engine = X.XCACHE, X.ENGINES["windows"]
    try:
        paths = _make_images(td, ("done.jpg", "new1.jpg", "new2.jpg"))
        stats = {p: os.stat(p) for p in paths}
        X.XCACHE = os.path.join(td, "cross")
        X._write_engine_cache("windows", paths[0], "done", stats[paths[0]])
        calls: list[list[str]] = []

        def fake_windows(batch, _timeout):
            calls.append(list(batch))
            return {p: "win:" + os.path.basename(p) for p in batch}

        X.ENGINES["windows"] = ("test", lambda: True, fake_windows, "test")
        first = X.read_texts("windows", paths, stats=stats)
        assert calls == [[paths[1]], [paths[2]]], calls
        calls.clear()
        second = X.read_texts("windows", paths, stats=stats)
        assert calls == [], "교차검증 완료 사진을 다시 OCR 했다: %r" % calls
        assert first == second
    finally:
        X.XCACHE, X.ENGINES["windows"] = real_cross, real_engine
        shutil.rmtree(td, ignore_errors=True)


if __name__ == "__main__":
    test_stat_is_reused_and_only_new_two_are_read()
    test_interruption_keeps_the_last_completed_image()
    test_crosscheck_engine_reuses_completed_cache()
    print("OCR [205] OK — stat 1회 재사용 · 완료 0장 · 새 2장 · 중단 뒤 이어받기")
