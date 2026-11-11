"""Tests for arabic-first-docs. Run with: python3 -m unittest -v"""

import json
import os
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

import arabic_first_docs as afd  # noqa: E402


SPEC = {
    "title": "التقرير الشهري",
    "sections": [
        {"heading": "أولاً", "paragraphs": ["نص عربي للاختبار.", "نص ثانٍ."],
         "table": [["العمود", "القيمة"], ["الأول", "١"], ["الثاني", "٢"]]},
    ],
}


def codes(findings):
    return [f.code for f in findings]


class BuildTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.out = os.path.join(self.tmp.name, "out.docx")
        self.findings = afd.build(SPEC, self.out)

    def tearDown(self):
        self.tmp.cleanup()

    def test_built_document_passes_check(self):
        self.assertEqual(codes(self.findings), [])

    def test_doc_defaults_carry_bidi(self):
        doc = Document(self.out)
        defaults = doc.styles.element.find(qn("w:docDefaults"))
        self.assertIsNotNone(defaults)
        path = f"{qn('w:pPrDefault')}/{qn('w:pPr')}/{qn('w:bidi')}"
        self.assertIsNotNone(defaults.find(path))

    def test_every_arabic_paragraph_has_bidi(self):
        doc = Document(self.out)
        missing = [p.text for p in doc.paragraphs
                   if afd.ARABIC.search(p.text or "")
                   and (p._p.pPr is None or p._p.pPr.find(qn("w:bidi")) is None)]
        self.assertEqual(missing, [])

    def test_runs_carry_rtl_and_complex_script_font(self):
        doc = Document(self.out)
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                rpr = run._element.rPr
                self.assertIsNotNone(rpr.find(qn("w:rtl")))
                self.assertTrue(rpr.find(qn("w:rFonts")).get(qn("w:cs")))

    def test_section_and_table_are_rtl(self):
        doc = Document(self.out)
        self.assertIsNotNone(doc.sections[0]._sectPr.find(qn("w:bidi")))
        self.assertIsNotNone(doc.tables[0]._tbl.tblPr.find(qn("w:bidiVisual")))

    def test_bidi_is_in_valid_position(self):
        """w:bidi must precede w:spacing, w:ind and w:jc inside w:pPr."""
        doc = Document(self.out)
        for paragraph in doc.paragraphs:
            ppr = paragraph._p.pPr
            if ppr is None or ppr.find(qn("w:bidi")) is None:
                continue
            order = [child.tag for child in ppr]
            bidi_at = order.index(qn("w:bidi"))
            for tag in ("w:spacing", "w:ind", "w:jc"):
                if qn(tag) in order:
                    self.assertLess(bidi_at, order.index(qn(tag)))


class DetectionTest(unittest.TestCase):
    def test_plain_document_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "plain.docx")
            doc = Document()
            doc.add_paragraph("نص عربي بلا ضبط للاتجاه.")
            doc.save(path)
            findings = afd.check(path)
        found = codes(findings)
        self.assertIn("NO_DOC_DEFAULTS", found)
        self.assertIn("MISSING_PARA_BIDI", found)
        self.assertIn("MISSING_RUN_RTL", found)

    def test_fix_repairs_a_plain_document(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "plain.docx")
            fixed = os.path.join(tmp, "fixed.docx")
            doc = Document()
            doc.add_paragraph("نص عربي بلا ضبط للاتجاه.")
            doc.save(source)
            self.assertNotEqual(codes(afd.check(source)), [])
            self.assertEqual(codes(afd.fix(source, fixed)), [])

    def test_english_document_is_not_flagged(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "en.docx")
            findings = afd.build({"title": "Monthly report",
                                  "sections": [{"heading": "Purpose",
                                                "paragraphs": ["English only."]}]},
                                 path, lang="en")
            self.assertEqual(codes(findings), [])

    def test_unreadable_file_is_a_finding_not_a_crash(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "not-a-docx.docx")
            with open(path, "w", encoding="utf-8") as fh:
                fh.write("not a zip")
            self.assertEqual(codes(afd.check(path)), ["UNREADABLE"])


class CliTest(unittest.TestCase):
    def run_cli(self, *args):
        return subprocess.run([sys.executable, os.path.join(HERE, "arabic_first_docs.py"), *args],
                              capture_output=True, text=True)

    def test_build_then_check_exit_zero(self):
        with tempfile.TemporaryDirectory() as tmp:
            spec = os.path.join(tmp, "spec.json")
            out = os.path.join(tmp, "out.docx")
            with open(spec, "w", encoding="utf-8") as fh:
                json.dump(SPEC, fh, ensure_ascii=False)
            built = self.run_cli("build", spec, out)
            self.assertEqual(built.returncode, 0, built.stdout + built.stderr)
            self.assertIn("PASS", built.stdout)
            checked = self.run_cli("check", out)
            self.assertEqual(checked.returncode, 0, checked.stdout)

    def test_check_json_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            spec = os.path.join(tmp, "spec.json")
            out = os.path.join(tmp, "out.docx")
            with open(spec, "w", encoding="utf-8") as fh:
                json.dump(SPEC, fh, ensure_ascii=False)
            self.run_cli("build", spec, out)
            result = self.run_cli("check", out, "--json")
        payload = json.loads(result.stdout)
        self.assertEqual(payload["result"], "PASS")

    def test_missing_spec_exits_two(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = self.run_cli("build", os.path.join(tmp, "nope.json"),
                                  os.path.join(tmp, "out.docx"))
        self.assertEqual(result.returncode, 2)


if __name__ == "__main__":
    unittest.main()
