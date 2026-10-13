from __future__ import annotations

import importlib.util
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("project_build", ROOT / "scripts/build.py")
builder = importlib.util.module_from_spec(spec)
spec.loader.exec_module(builder)
audit_spec = importlib.util.spec_from_file_location("project_audit", ROOT / "scripts/audit.py")
auditor = importlib.util.module_from_spec(audit_spec)
audit_spec.loader.exec_module(auditor)


class BuildTests(unittest.TestCase):
    def test_independent_offline_and_reproducible_build(self):
        with tempfile.TemporaryDirectory() as directory:
            isolated = Path(directory) / "independent-project"
            isolated.mkdir()
            for name in ["src", "scripts", "templates", "vendor"]:
                shutil.copytree(ROOT / name, isolated / name, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
            shutil.copyfile(ROOT / "VERSION", isolated / "VERSION")
            for destination in ["first", "second"]:
                subprocess.run([sys.executable, str(isolated / "scripts/build.py"), "--output-dir", str(isolated / destination)],
                               cwd=directory, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            first = next((isolated / "first").glob("*.ppam"))
            second = next((isolated / "second").glob("*.ppam"))
            self.assertEqual(first.read_bytes(), second.read_bytes())
            self.assertEqual(set(auditor.source_modules(first)), {"ColorReplaceInPlace", "ColorPickerForm"})
            self.assertEqual(auditor.check_ooxml(first), [])
            with ZipFile(first) as archive:
                ribbon = ET.fromstring(archive.read("customUI/customUI14.xml"))
                labels = [node.get("label", "") for node in ribbon.iter()]
                self.assertTrue(any("测试版" in label for label in labels))
                self.assertFalse(any(name.startswith("ppt/slides/") for name in archive.namelist()))

    def test_unicode_source_encoding(self):
        encoded = builder.encode_vba('MsgBox "颜色 " & "预览"\nMsgBox "112233"\n')
        self.assertTrue(encoded.isascii())
        self.assertIn('ChrW$(39068)', encoded)
        self.assertIn('"112233"', encoded)

    def test_third_party_snapshot_and_public_tree(self):
        self.assertEqual(auditor.audit()["issues"], [])


if __name__ == "__main__":
    unittest.main()
