# -*- coding: utf-8 -*-
"""Canonical source repository full audit.

The audit is read-only for ``0. 원본 자료``.  ``robocopy /L`` is used only as
an efficient network-drive directory enumerator; no source file is copied,
moved, renamed, or deleted.
"""
from __future__ import annotations

import argparse
import collections
import concurrent.futures
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

from source_dirs import ORIGIN_ROOT

ROOT = Path(__file__).resolve().parent
REPORT_DIR = ROOT / "reports"
ROBOCOPY_FILE = re.compile(r"^\s*(\d+)\s+(.+)$")


def enumerate_files(root: str) -> list[dict]:
    """Return every file without recursively stat-ing slow SMB paths."""
    null_target = os.path.join(tempfile.gettempdir(), "codex-source-audit-null")
    command = [
        "robocopy", root, null_target, "/L", "/E", "/FP", "/BYTES",
        "/NJH", "/NJS", "/NC", "/R:0", "/W:0",
    ]
    result = subprocess.run(
        command, capture_output=True, text=True, errors="replace", timeout=600,
        check=False,
    )
    # Robocopy 0..7 are non-fatal result bitmasks.  /L never copies a file.
    if result.returncode > 7:
        raise RuntimeError(
            f"robocopy enumeration failed ({result.returncode}): "
            f"{result.stderr[-500:]}"
        )
    files = []
    for line in result.stdout.splitlines():
        match = ROBOCOPY_FILE.match(line)
        if not match:
            continue
        size, path = int(match.group(1)), match.group(2).strip()
        if path.endswith("\\"):
            continue
        files.append({
            "path": path,
            "relative_path": os.path.relpath(path, root),
            "size": size,
            "extension": Path(path).suffix.lower() or "(none)",
        })
    return files


def _read_prefix(path: str, size: int = 32) -> bytes:
    with open(path, "rb") as stream:
        return stream.read(size)


def validate_file(item: dict) -> dict:
    """Perform a lightweight format-aware integrity check."""
    path = item["path"]
    ext = item["extension"]
    try:
        if item["size"] == 0:
            return {"status": "error", "detail": "0-byte file"}
        if ext == ".xlsx":
            if not zipfile.is_zipfile(path):
                return {"status": "error", "detail": "invalid xlsx container"}
            with zipfile.ZipFile(path) as archive:
                bad_member = archive.testzip()
                if bad_member:
                    return {"status": "error", "detail": f"bad zip member: {bad_member}"}
                required = {"[Content_Types].xml", "xl/workbook.xml"}
                missing = sorted(required.difference(archive.namelist()))
                if missing:
                    return {"status": "error", "detail": f"missing: {', '.join(missing)}"}
            return {"status": "ok", "detail": "xlsx zip/xml"}
        if ext == ".json":
            with open(path, encoding="utf-8-sig") as stream:
                json.load(stream)
            return {"status": "ok", "detail": "json"}
        if ext in {".txt", ".csv"}:
            raw = Path(path).read_bytes()
            for encoding in ("utf-8-sig", "cp949"):
                try:
                    raw.decode(encoding)
                    return {"status": "ok", "detail": encoding}
                except UnicodeDecodeError:
                    continue
            return {"status": "error", "detail": "text decoding failed"}
        if ext in {".jpg", ".jpeg"}:
            head = _read_prefix(path)
            ok = head.startswith(b"\xff\xd8\xff")
            return {"status": "ok" if ok else "error", "detail": "jpeg signature"}
        if ext == ".png":
            head = _read_prefix(path)
            ok = head.startswith(b"\x89PNG\r\n\x1a\n")
            return {"status": "ok" if ok else "error", "detail": "png signature"}
        if ext == ".pdf":
            with open(path, "rb") as stream:
                head = stream.read(8)
                stream.seek(max(0, item["size"] - 4096))
                tail = stream.read()
            if not head.startswith(b"%PDF-"):
                return {"status": "error", "detail": "pdf header missing"}
            if b"%%EOF" not in tail:
                return {"status": "error", "detail": "pdf EOF missing"}
            return {"status": "ok", "detail": "pdf header/eof"}
        if ext == ".db":
            return {"status": "skipped", "detail": "system cache"}
        return {"status": "skipped", "detail": "unsupported extension"}
    except Exception as exc:  # report the file; never stop the full audit
        return {"status": "error", "detail": f"{type(exc).__name__}: {exc}"}


def sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def exact_duplicates(files: list[dict], workers: int, scope: str) -> list[dict]:
    """Hash equal-size candidates and return exact duplicate groups.

    ``collision`` limits hashing to size groups containing a preserved
    ``__dup_`` collision copy.  This is the practical default for the network
    repository: hashing every same-size PO/PDF and image can hold the SMB share
    open for many minutes while adding no safe automatic action.
    """
    by_size: dict[int, list[dict]] = collections.defaultdict(list)
    for item in files:
        if item["size"] > 0:
            by_size[item["size"]].append(item)
    groups = [group for group in by_size.values() if len(group) > 1]
    if scope == "collision":
        groups = [
            group for group in groups
            if any("__dup_" in item["relative_path"].lower() for item in group)
        ]
    candidates = [item for group in groups for item in group]
    if not candidates:
        return []
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
        hashes = list(pool.map(lambda item: sha256(item["path"]), candidates))
    by_hash: dict[tuple[int, str], list[dict]] = collections.defaultdict(list)
    for item, digest in zip(candidates, hashes):
        by_hash[(item["size"], digest)].append(item)
    return [
        {
            "sha256": digest,
            "size": size,
            "count": len(group),
            "files": sorted(item["relative_path"] for item in group),
        }
        for (size, digest), group in by_hash.items()
        if len(group) > 1
    ]


def audit(root: str, workers: int, hash_scope: str = "collision") -> dict:
    files = enumerate_files(root)
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
        checks = list(pool.map(validate_file, files))
    for item, check in zip(files, checks):
        item["validation"] = check

    extension_counts = collections.Counter(item["extension"] for item in files)
    folder_counts = collections.Counter(item["relative_path"].split(os.sep, 1)[0] for item in files)
    validation_counts = collections.Counter(
        item["validation"]["status"] for item in files
    )
    errors = [
        {
            "relative_path": item["relative_path"],
            "detail": item["validation"]["detail"],
        }
        for item in files
        if item["validation"]["status"] == "error"
    ]
    duplicates = exact_duplicates(files, workers, hash_scope) if hash_scope != "none" else []
    return {
        "schema_version": 1,
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "root": root,
        "read_only": True,
        "summary": {
            "files": len(files),
            "bytes": sum(item["size"] for item in files),
            "zero_byte": sum(item["size"] == 0 for item in files),
            "extensions": dict(sorted(extension_counts.items())),
            "top_level": dict(sorted(folder_counts.items())),
            "validation": dict(sorted(validation_counts.items())),
            "errors": len(errors),
            "duplicate_hash_scope": hash_scope,
            "exact_duplicate_groups": len(duplicates),
            "exact_duplicate_extra_files": sum(group["count"] - 1 for group in duplicates),
        },
        "errors": errors,
        "exact_duplicates": sorted(
            duplicates, key=lambda group: (-group["size"], group["sha256"])
        ),
        "files": files,
    }


def markdown(report: dict) -> str:
    summary = report["summary"]
    lines = [
        "# 원본 자료 전수 감사",
        "",
        f"- 생성: {report['generated_at']}",
        f"- 원본: `{report['root']}`",
        f"- 파일: **{summary['files']:,}개** / **{summary['bytes']:,} bytes**",
        f"- 형식검증 오류: **{summary['errors']}개** / 빈 파일: **{summary['zero_byte']}개**",
        f"- 완전중복: **{summary['exact_duplicate_groups']}그룹** "
        f"(추가 사본 {summary['exact_duplicate_extra_files']}개, "
        f"해시범위: {summary['duplicate_hash_scope']})",
        "",
        "## 최상위 분류",
        "",
        "| 분류 | 파일 |",
        "|---|---:|",
    ]
    lines.extend(f"| {name} | {count:,} |" for name, count in summary["top_level"].items())
    lines += ["", "## 확장자", "", "| 확장자 | 파일 |", "|---|---:|"]
    lines.extend(f"| {ext} | {count:,} |" for ext, count in summary["extensions"].items())
    lines += ["", "## 형식검증 오류", ""]
    if report["errors"]:
        lines += ["| 파일 | 오류 |", "|---|---|"]
        lines.extend(
            f"| {row['relative_path']} | {row['detail']} |"
            for row in report["errors"]
        )
    else:
        lines.append("- 없음")
    lines += ["", "## 완전중복", ""]
    if report["exact_duplicates"]:
        for group in report["exact_duplicates"]:
            lines.append(
                f"- {group['size']:,} bytes · {group['count']}개 · "
                f"`{group['sha256'][:12]}`"
            )
            lines.extend(f"  - `{name}`" for name in group["files"])
    else:
        lines.append("- 없음")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", default=ORIGIN_ROOT)
    parser.add_argument("--workers", type=int, default=16)
    parser.add_argument("--skip-hash", action="store_true")
    parser.add_argument(
        "--hash-all", action="store_true",
        help="모든 동일 크기 후보를 해시(네트워크 원본에서는 매우 느릴 수 있음)",
    )
    parser.add_argument("--json", default=str(REPORT_DIR / "source_full_audit_latest.json"))
    parser.add_argument("--md", default=str(REPORT_DIR / "source_full_audit_latest.md"))
    args = parser.parse_args()

    hash_scope = "none" if args.skip_hash else ("all" if args.hash_all else "collision")
    report = audit(args.root, max(1, args.workers), hash_scope)
    Path(args.json).parent.mkdir(parents=True, exist_ok=True)
    Path(args.json).write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    Path(args.md).write_text(markdown(report), encoding="utf-8")
    summary = report["summary"]
    print(
        f"원본 전수감사: {summary['files']:,}개 · "
        f"형식오류 {summary['errors']} · 빈 파일 {summary['zero_byte']} · "
        f"완전중복 {summary['exact_duplicate_groups']}그룹/"
        f"{summary['exact_duplicate_extra_files']}추가사본"
    )
    print(args.md)
    return 1 if summary["errors"] else 0


if __name__ == "__main__":
    sys.exit(main())
