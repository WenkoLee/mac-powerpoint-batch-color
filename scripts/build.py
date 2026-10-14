#!/usr/bin/env python3
"""Build the Mac add-in offline from the source tree; no PowerPoint required."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import sys
import uuid
import warnings
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "vendor"))
warnings.filterwarnings("ignore", message="No Python codec for VBA code page 10025.*")
from pyopenvba import PowerPointFile

FIXED_TIME = (2026, 1, 1, 0, 0, 0)


def encode_vba(source: str) -> str:
    """Keep UTF-8 source readable; encode non-ASCII string literals for Mac VBA."""
    def replace(match: re.Match) -> str:
        value = match[0][1:-1].replace('""', '"')
        if value.isascii():
            return match[0]
        parts, segment = [], ""
        for char in value:
            if ord(char) < 128:
                segment += char
            else:
                if segment:
                    parts.append('"' + segment.replace('"', '""') + '"')
                    segment = ""
                parts.append(f"ChrW$({ord(char)})")
        if segment:
            parts.append('"' + segment.replace('"', '""') + '"')
        return " & ".join(parts) or '""'

    encoded = re.sub(r'"(?:[^"\r\n]|"")*"', replace, source)
    # Comments are documentation, never executable Unicode identifiers.
    encoded = encoded.encode("ascii", "backslashreplace").decode("ascii")
    if max(map(len, encoded.splitlines()), default=0) >= 1024:
        raise ValueError("A VBA source line exceeds the supported line length")
    return encoded


def stable_zip(path: Path, parts: dict[str, bytes]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(path, "w") as package:
        for name, data in sorted(parts.items()):
            info = ZipInfo(name, FIXED_TIME)
            info.compress_type = ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            package.writestr(info, data)


def add_panel(presentation: PowerPointFile) -> None:
    form = presentation.add_form("ColorPickerForm", caption="", width=680, height=418)
    for control in json.loads((ROOT / "src/form-controls.json").read_text()):
        item = form.add_control(control["type"], control["name"], left=0, top=0, width=10, height=10)
        if "caption" in control:
            # The runtime uses ChrW for all Chinese captions.
            item.set_property("Caption", "")
    module = presentation.vba_project().get_module("ColorPickerForm")
    identities = iter(str(uuid.uuid5(uuid.NAMESPACE_URL, "mac-powerpoint-batch-color/" + role)).upper()
                      for role in ["form", "designer"])
    module.attribute_header = re.sub(r"\{[0-9A-F-]{36}\}", lambda _: "{" + next(identities) + "}",
                                     module.attribute_header, count=2)
    presentation.set_module("ColorPickerForm", encode_vba((ROOT / "src/ColorPickerForm.bas").read_text()))


def build(output: Path, with_qa: bool = False) -> Path:
    version = (ROOT / "VERSION").read_text().strip()
    build_dir = ROOT / "build"
    build_dir.mkdir(exist_ok=True)
    mother = build_dir / "MacPowerPointBatchColor.pptm"
    engine = encode_vba((ROOT / "src/ColorReplaceInPlace.bas").read_text())
    with PowerPointFile(ROOT / "templates/powerpoint-base.pptm") as presentation:
        presentation.set_module("ColorReplaceInPlace", engine)
        add_panel(presentation)
        presentation.save(mother)

    with PowerPointFile(mother) as check:
        assert set(check.module_names()) == {"ColorReplaceInPlace", "ColorPickerForm"}
        assert len(check.forms()) == 1 and len(check.forms()[0].walk()) == len(json.loads((ROOT / "src/form-controls.json").read_text()))
        if check.validate():
            raise ValueError(check.validate())

    with ZipFile(mother) as source, ZipFile(ROOT / "templates/addin-container.zip") as template:
        parts = {name: template.read(name) for name in template.namelist()}
        parts["ppt/vbaProject.bin"] = source.read("ppt/vbaProject.bin")
    parts["customUI/customUI14.xml"] = (ROOT / "src/ribbon.xml").read_bytes()
    package = output / ("一键改色_" + re.sub(r"[^0-9A-Za-z]", "_", version) + ".ppam")
    stable_zip(package, parts)
    with ZipFile(package) as check:
        assert check.testzip() is None
        ET.fromstring(check.read("customUI/customUI14.xml"))
        content_types = ET.fromstring(check.read("[Content_Types].xml"))
        assert any(node.get("ContentType") == "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"
                   for node in content_types)

    if with_qa:
        with PowerPointFile(mother) as presentation:
            source = encode_vba((ROOT / "tests/AcceptanceRunner.bas").read_text())
            presentation.vba_project().add_module("AcceptanceRunner", source)
            qa_path = build_dir / "AcceptanceHarness.pptm"
            presentation.save(qa_path)
        # Mac does not expose Shapes.AddChart/AddChart2. Use a synthetic native
        # chart fixture, then attach the QA VBA project using standard OOXML.
        with ZipFile(qa_path) as macros, ZipFile(ROOT / "templates/qa-fixture.pptx") as fixture:
            qa_parts = {name: fixture.read(name) for name in fixture.namelist()}
            qa_parts["ppt/vbaProject.bin"] = macros.read("ppt/vbaProject.bin")
        ct = ET.fromstring(qa_parts["[Content_Types].xml"])
        ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        for node in ct:
            if node.get("PartName") == "/ppt/presentation.xml":
                node.set("ContentType", "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml")
        ET.SubElement(ct, "{" + ns + "}Override", PartName="/ppt/vbaProject.bin",
                      ContentType="application/vnd.ms-office.vbaProject")
        qa_parts["[Content_Types].xml"] = ET.tostring(ct, encoding="utf-8", xml_declaration=True)
        rel_name = "ppt/_rels/presentation.xml.rels"
        rels = ET.fromstring(qa_parts[rel_name])
        ET.SubElement(rels, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                      Id="rIdQAVba", Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject", Target="vbaProject.bin")
        qa_parts[rel_name] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)
        stable_zip(qa_path, qa_parts)

    digest = hashlib.sha256(package.read_bytes()).hexdigest()
    (output / "SHA256SUMS.txt").write_text(f"{digest}  {package.name}\n")
    manifest = {"project": "一键改色 / One-Click Color for Mac PowerPoint", "version": version,
                "status": "release", "package": package.name, "sha256": digest,
                "source_sha256": {p.relative_to(ROOT).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
                                  for p in sorted((ROOT / "src").iterdir()) if p.is_file()}}
    (output / "build-manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")
    print(package)
    print("SHA256:", digest)
    return package


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=ROOT / "release")
    parser.add_argument("--with-qa", action="store_true", help="also create the local test harness")
    args = parser.parse_args()
    build(args.output_dir.resolve(), args.with_qa)
