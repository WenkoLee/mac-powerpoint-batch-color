#!/usr/bin/env python3
"""Audit public project files, build inputs, third-party snapshot and OOXML links."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path, PurePosixPath
import posixpath
import sys
import warnings
from xml.etree import ElementTree as ET
from zipfile import ZipFile, is_zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "vendor"))
warnings.filterwarnings("ignore", message="No Python codec for VBA code page 10025.*")
from pyopenvba.cfb import CFB
from pyopenvba.vba import parse_vba_project

EXCLUDED = {".git", ".venv", "__pycache__", "build", "dist"}
PRIVATE_MARKERS = ["install-" + "brightslide" + "-1-1-2"]


def public_files(root: Path = ROOT) -> list[Path]:
    return [p for p in sorted(root.rglob("*"))
            if p.is_file() and not any(part in EXCLUDED for part in p.relative_to(root).parts)
            and not p.name.startswith("~$") and p.name != ".DS_Store"]


def source_modules(package: Path) -> dict[str, str]:
    with ZipFile(package) as archive:
        if "ppt/vbaProject.bin" not in archive.namelist():
            return {}
        project = parse_vba_project(CFB.from_bytes(archive.read("ppt/vbaProject.bin")))
        return {module.name: module.source for module in project.modules}


def check_ooxml(package: Path) -> list[str]:
    problems = []
    with ZipFile(package) as archive:
        names = set(archive.namelist())
        if archive.testzip() is not None:
            problems.append("ZIP integrity failure")
        for name in names:
            if name.endswith((".xml", ".rels")):
                try:
                    tree = ET.fromstring(archive.read(name))
                except ET.ParseError as error:
                    problems.append(f"{name}: {error}")
                    continue
                if name.endswith(".rels"):
                    base = "" if name == "_rels/.rels" else posixpath.dirname(posixpath.dirname(name))
                    for rel in tree:
                        target = rel.get("Target", "")
                        if rel.get("TargetMode") == "External":
                            problems.append(f"External relationship: {name}")
                            continue
                        resolved = posixpath.normpath(posixpath.join(base, target)) if not target.startswith("/") else target[1:]
                        if resolved not in names:
                            problems.append(f"Missing relationship target: {name} -> {resolved}")
        # UTF-8 / UTF-16 text leaks are checked in decompressed XML and VBA source.
        text = "\n".join(archive.read(n).decode("utf-8", "ignore") for n in names if n.endswith((".xml", ".rels")))
        text += "\n" + "\n".join(source_modules(package).values())
        for marker in PRIVATE_MARKERS:
            if marker in text:
                problems.append("Private source marker in package")
    return problems


def audit(root: Path = ROOT) -> dict:
    files, issues = public_files(root), []
    manifest = json.loads((root / "vendor/SOURCE.json").read_text())
    for name, expected in manifest["files"].items():
        path = root / "vendor" / name
        if not path.exists() or hashlib.sha256(path.read_bytes()).hexdigest() != expected:
            issues.append(f"Dependency snapshot mismatch: {name}")
    for path in files:
        relative = path.relative_to(root).as_posix()
        if path.is_symlink():
            issues.append(f"Symlink in public tree: {relative}")
        if path.suffix.lower() in {".pptx", ".pptm", ".ppam"}:
            allowed = (relative == "templates/powerpoint-base.pptm" or relative == "templates/qa-fixture.pptx" or
                       relative == "vendor/pyopenvba/_templates/blank_files/blank_presentation.pptm" or
                       relative.startswith("release/MacPowerPointBatchColor_") or relative.startswith("release/一键改色_") or
                       relative.startswith("examples/"))
            if not allowed:
                issues.append(f"Unexpected presentation: {relative}")
            issues.extend(f"{relative}: {problem}" for problem in check_ooxml(path))
        if path.suffix in {".md", ".bas", ".py", ".json", ".xml"} and not relative.startswith("vendor/"):
            text = path.read_text()
            if any(marker in text for marker in PRIVATE_MARKERS):
                issues.append(f"Private source marker: {relative}")
            if "/" + "Users/" in text or "/" + "private/var/" in text:
                issues.append(f"Personal absolute path: {relative}")
        if relative.startswith("release/") and path.suffix == ".ppam":
            if set(source_modules(path)) != {"ColorReplaceInPlace", "ColorPickerForm"}:
                issues.append(f"Unexpected VBA modules in release: {relative}")
    result = {"status": "pass" if not issues else "fail", "public_file_count": len(files),
              "dependency_version": manifest["version"], "issues": issues,
              "presentations": [p.relative_to(root).as_posix() for p in files if p.suffix in {".pptx", ".pptm", ".ppam"}]}
    return result


if __name__ == "__main__":
    result = audit()
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if result["issues"]:
        raise SystemExit(1)
