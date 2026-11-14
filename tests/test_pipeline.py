# -*- coding: utf-8 -*-
"""اختبارات تراجع لمحركات التفريغ الثلاثة.

تشغيل محلياً:
    PYTHONPATH=app/src/main/python python3 -m unittest discover -s tests -v
"""
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from docx import Document

import bridge
import expression
import generate
import language
import linkimport
import reading_general
import styles

ROOT = Path(__file__).resolve().parents[1]
PY = ROOT / "app" / "src" / "main" / "python"
MODELS = ROOT / "app" / "src" / "main" / "assets" / "models"


def _stage(name, suffix=""):
    return {
        "name": name,
        "capabilities": [f"- قدرة {suffix}".strip()],
        "teacher": [f"- نشاط المدرس {suffix}".strip()],
        "content": [f"- محتوى {suffix}".strip()],
        "evaluation": [f"- تقويم {suffix}".strip()],
    }


def make_source(path: Path, component: str, stages, topic="قيم العمل", unit="الوحدة الأولى"):
    data = {
        "header": {
            "unit": unit,
            "domain": "مجال القيم",
            "topic": topic,
            "duration": "ساعة واحدة",
            "competencies": {"تواصلية": "التواصل", "منهجية": "التحليل", "ثقافية": "الاعتزاز"},
        },
        "stages": [_stage(s, str(i + 1)) for i, s in enumerate(stages)],
    }
    return generate.write_source_docx(data, component, topic, "المختار", "الثالثة إعدادي", path)


def all_text(path: Path) -> str:
    doc = Document(str(path))
    parts = [p.text for p in doc.paragraphs]
    parts += [c.text for t in doc.tables for row in t.rows for c in row.cells]
    return "\n".join(parts)


class ReadingTests(unittest.TestCase):
    def test_generated_header_and_no_stale_template(self):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            src = make_source(td / "reading.docx", "reading", generate.STAGES["reading"])
            parsed = reading_general.Source(str(src))
            self.assertEqual(parsed.info.get("unit"), "الأولى")
            header = reading_general.find_header(parsed.grids[0])
            self.assertIsNotNone(header)
            self.assertEqual(header[1], 1)
            self.assertEqual(header[2].get(3), "content")

            out = json.loads(bridge.run("reading", str(MODELS / "reading.docx"), str(src), str(td / "out"),
                                        out_name="قراءة.docx"))
            self.assertTrue(out["ok"], out)
            self.assertEqual(out["target"]["unit"], "الوحدة الأولى")
            text = all_text(Path(out["path"]))
            self.assertNotIn("روم لاندو", text)
            self.assertNotIn("محنة نفي الملك", text)
            self.assertIn("قدرة 1", text)
            self.assertIn("محتوى 1", text)

    def test_missing_stage_does_not_keep_template_example(self):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            src = make_source(td / "partial.docx", "reading", ["تقويم تشخيصي", "الاستثمار"])
            out = json.loads(bridge.run("reading", str(MODELS / "reading.docx"), str(src), str(td / "out"),
                                        out_name="ناقص.docx"))
            self.assertTrue(out["ok"], out)
            text = all_text(Path(out["path"]))
            self.assertNotIn("روم لاندو", text)
            self.assertNotIn("المنفى", text)


class LanguageTests(unittest.TestCase):
    def test_generated_source_roundtrip(self):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            src = make_source(td / "lang.docx", "language", generate.STAGES["language"], topic="اسم الفاعل")
            out = json.loads(bridge.run("language", str(MODELS / "language.docx"), str(src), str(td / "out"),
                                        out_name="لغة.docx"))
            self.assertTrue(out["ok"], out)
            self.assertEqual(out["target"]["unit"], "الوحدة الأولى")
            text = all_text(Path(out["path"]))
            self.assertIn("اسم الفاعل", text)
            self.assertIn("محتوى 2", text)

    def test_dotenv_api_key(self):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            old = Path.cwd()
            try:
                os.chdir(td)
                (td / ".env").write_text("GEMINI_API_KEY=test-key-123\n", encoding="utf-8")
                with mock.patch.dict(os.environ, {}, clear=True):
                    self.assertEqual(language.resolve_api_key(language.Settings()), "test-key-123")
            finally:
                os.chdir(old)


class ExpressionTests(unittest.TestCase):
    def test_all_source_rows_survive(self):
        stages = generate.EXPRESSION_STAGES["الاكتساب"]
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            src = make_source(td / "expr.docx", "expression", stages, topic="السرد")
            with mock.patch.object(expression, "source_kind", wraps=expression.source_kind) as sk:
                out = json.loads(bridge.run("expression", str(MODELS / "expression.docx"), str(src),
                                            str(td / "out"), out_name="تعبير.docx"))
            self.assertTrue(out["ok"], out)
            self.assertNotIn("نشاط واحد", out["log"])
            doc = Document(out["path"])
            main = doc.tables[-1]
            self.assertEqual(len(main.rows), len(stages) + 1)
            text = all_text(Path(out["path"]))
            for stage in stages:
                self.assertIn(stage, text)


class LinkImportTests(unittest.TestCase):
    def test_rejects_http_and_private_hosts(self):
        with self.assertRaises(linkimport.LinkError):
            linkimport.fetch_text("http://example.com")
        with self.assertRaises(linkimport.LinkError):
            linkimport.fetch_text("https://127.0.0.1/x")
        with self.assertRaises(linkimport.LinkError):
            linkimport.fetch_text("https://localhost/x")

    def test_source_text_is_wrapped_as_data(self):
        prompt = generate._structure_prompt("reading", "نص التحضير", generate.STAGES["reading"])
        self.assertIn("<source_text>", prompt)
        self.assertIn("</source_text>", prompt)
        self.assertIn("بيانات تحضير فقط", prompt)


class StylesTests(unittest.TestCase):
    def test_every_style_applies_without_error(self):
        keys = list(styles.load_styles().keys())
        self.assertGreaterEqual(len(keys), 10)
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            src = make_source(td / "reading.docx", "reading", generate.STAGES["reading"])
            for key in keys:
                out = json.loads(bridge.run("reading", str(MODELS / "reading.docx"), str(src),
                                            str(td / "out"), out_name=f"{key}.docx", style=key))
                self.assertTrue(out["ok"], (key, out))
                self.assertNotIn("تعذّر تطبيق النمط", out["log"])

    def test_classic_is_noop_and_unknown_style_raises(self):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            src = make_source(td / "reading.docx", "reading", generate.STAGES["reading"])
            self.assertFalse(styles.apply_style(str(src), "classic"))
            with self.assertRaises(ValueError):
                styles.apply_style(str(src), "no_such_style")


if __name__ == "__main__":
    unittest.main()
