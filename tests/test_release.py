"""Verify packaged skills and the coursework helpers using synthetic fixtures."""
from hashlib import sha256
import json
from pathlib import Path
import re
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from PIL import Image
import yaml

ROOT = Path(__file__).resolve().parents[1]
SKILLS = ROOT / "skills"
SCRIPTS = SKILLS / "hhy-coursework-delivery" / "scripts"
FRONTMATTER = re.compile(r"\A---\n(.*?)\n---(?:\n|$)", re.DOTALL)
LINK = re.compile(r"\]\(([^)\s]+)\)")


class PackageTest(unittest.TestCase):
    def test_skill_frontmatter_and_interface(self):
        self.assertEqual(
            {item.name for item in SKILLS.iterdir() if item.is_dir()},
            {"hhy-learning-coach", "hhy-coursework-delivery"},
        )
        for entry in SKILLS.glob("*/SKILL.md"):
            with self.subTest(skill=entry.parent.name):
                skill = yaml.safe_load(FRONTMATTER.match(entry.read_text(encoding="utf-8"))[1])
                metadata = yaml.safe_load(
                    (entry.parent / "agents/openai.yaml").read_text(encoding="utf-8")
                )
                self.assertEqual(skill["name"], entry.parent.name)
                self.assertTrue(skill["description"])
                self.assertIn("$" + skill["name"], metadata["interface"]["default_prompt"])
                self.assertTrue((entry.parent / "references/upstream-licenses.txt").is_file())

    def test_markdown_relative_links_exist(self):
        for note in ROOT.rglob("*.md"):
            if ".git" in note.relative_to(ROOT).parts:
                continue
            for url in LINK.findall(note.read_text(encoding="utf-8")):
                target = url.split("#", 1)[0]
                if target and "://" not in target:
                    resolved = (note.parent / target).resolve()
                    self.assertTrue(resolved.is_relative_to(ROOT), (note, url))
                    self.assertTrue(resolved.exists(), (note, url))


class CourseworkHelperTest(unittest.TestCase):
    def run_helper(self, script, *args):
        return subprocess.run(
            [sys.executable, "-X", "utf8", "-B", str(SCRIPTS / script), *map(str, args)],
            capture_output=True,
            encoding="utf-8",
            check=False,
        )

    def test_audit_keeps_source_unchanged_and_will_not_replace_report(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "report.docx"
            report = Path(tmp) / "audit.json"
            with ZipFile(source, "w") as package:
                package.writestr(
                    "word/document.xml",
                    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                    '<w:body><w:p><w:r><w:t>Sample report</w:t></w:r></w:p></w:body></w:document>',
                )
            original = sha256(source.read_bytes()).hexdigest()
            result = self.run_helper("audit_docx.py", source, "--out", report)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIsInstance(json.loads(report.read_text(encoding="utf-8")), dict)
            first = report.read_bytes()
            result = self.run_helper("audit_docx.py", source, "--out", report)
            self.assertNotEqual(result.returncode, 0)
            self.assertEqual(report.read_bytes(), first)
            self.assertEqual(sha256(source.read_bytes()).hexdigest(), original)

    def test_compare_reports_pixel_changes_and_rejects_wrong_page_count(self):
        with TemporaryDirectory() as tmp:
            before, after = Path(tmp) / "before", Path(tmp) / "after"
            before.mkdir()
            after.mkdir()
            image = Image.new("RGBA", (8, 8), (255, 255, 255, 255))
            image.save(before / "page-1.png")
            image.putpixel((4, 5), (255, 0, 0, 255))
            image.save(after / "page-1.png")
            result = self.run_helper(
                "compare_pages.py", before, after, "--out", Path(tmp) / "comparison.json"
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertEqual(json.loads(result.stdout)["needs_review"], [1])
            rejected = Path(tmp) / "rejected.json"
            result = self.run_helper(
                "compare_pages.py", before, after, "--after-pages", 2, "--out", rejected
            )
            self.assertNotEqual(result.returncode, 0)
            self.assertFalse(rejected.exists())


if __name__ == "__main__":
    unittest.main()
