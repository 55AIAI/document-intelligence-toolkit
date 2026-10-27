import tempfile
import unittest
from test_fallback import _make_docs, CHECKS
from converters.word_to_excel import read_docx, extract_fallback


class ExtractionFormats(unittest.TestCase):
    def test_six_synthetic_formats(self):
        from pathlib import Path
        with tempfile.TemporaryDirectory() as folder:
            _make_docs(folder)
            for name, check in CHECKS.items():
                with self.subTest(format=name):
                    self.assertTrue(check(extract_fallback(read_docx(str(Path(folder) / (name + ".docx"))))))


    def test_word_to_excel_to_word_without_llm(self):
        from pathlib import Path
        from contextlib import redirect_stdout
        from io import StringIO
        from docx import Document
        from converters.word_to_excel import run as word_run
        from converters.excel_to_word import run as excel_run
        with tempfile.TemporaryDirectory() as folder, redirect_stdout(StringIO()):
            folder = Path(folder)
            _make_docs(str(folder))
            word_run(str(folder / "A.docx"), use_llm=False, graph=True, html=True)
            self.assertTrue((folder / "A_graph.svg").exists())
            self.assertTrue((folder / "A_editor.html").exists())
            excel_run(str(folder / "A_scheme.xlsx"), use_llm=False)
            report = Document(folder / "A_scheme_spravka.docx")
            self.assertIn("Альфа", "\\n".join(p.text for p in report.paragraphs))

    def test_anonymized_conversion_does_not_load_photos(self):
        from pathlib import Path
        from contextlib import redirect_stdout
        from io import StringIO
        from unittest.mock import patch
        from converters.word_to_excel import run
        with tempfile.TemporaryDirectory() as folder, redirect_stdout(StringIO()):
            _make_docs(folder)
            with patch("converters.word_to_excel.load_photos", side_effect=AssertionError("photos accessed")):
                run(str(Path(folder) / "A.docx"), use_llm=False, anonymize=True)
            data = (Path(folder) / "A_ANON_entities.json").read_text(encoding="utf-8")
            self.assertNotIn("Иванов", data)
            self.assertNotIn("500", data)
