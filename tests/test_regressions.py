"""Synthetic regression cases; no model downloads or external services."""
import copy
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import Mock, patch

import requests
from anon import anonymize_data
from converters.word_to_excel import validate_and_dedup, graph_to_tree, extract_llm
from llm import qwen
from validation import validate_document
from writers.graph_html import build_html


def org(name="Demo Contractor", amount="100 млн тг", **kwargs):
    code = "000000000000" if name == "Demo Contractor" else f"{int(hashlib.sha256(name.encode()).hexdigest()[:12], 16) % 10**12:012d}"
    return {"name": name, "bin": code, "amount": amount,
            "subcontractors": [], **kwargs}


def document(*nodes):
    return {"project_name": "Synthetic demo", "customer": {"name": "Demo City"},
            "contractors": list(nodes)}


class ContractTests(unittest.TestCase):
    def test_distinct_contract_amounts_survive(self):
        data, _ = validate_and_dedup(document(org(amount="100 млн тг"), org(amount="200 млн тг")))
        self.assertEqual([n["amount"] for n in data["contractors"]], ["100 млн тг", "200 млн тг"])

    def test_shared_subcontractor_keeps_both_parents(self):
        data = document(org("Parent A", subcontractors=[org("Shared")]),
                        org("Parent B", subcontractors=[org("Shared")]))
        data, _ = validate_and_dedup(data)
        self.assertEqual([len(n["subcontractors"]) for n in data["contractors"]], [1, 1])

    def test_exact_sibling_duplicates_merge_children(self):
        first, second = org(), org()
        first["subcontractors"] = [org("Child A")]
        second["subcontractors"] = [org("Child B")]
        data, warnings = validate_and_dedup(document(first, second))
        self.assertEqual(len(data["contractors"]), 1)
        self.assertEqual(len(data["contractors"][0]["subcontractors"]), 2)
        self.assertTrue(warnings)

    def test_distinct_nested_contracts_survive(self):
        a, b = org(contract={"amount": "10 млн"}), org(contract={"amount": "20 млн"})
        self.assertEqual(len(validate_and_dedup(document(a, b))[0]["contractors"]), 2)

    def test_parent_and_child_same_company_stay_distinct(self):
        parent = org(subcontractors=[org(amount="50 млн тг")])
        self.assertEqual(len(validate_and_dedup(document(parent))[0]["contractors"][0]["subcontractors"]), 1)


class ValidationTests(unittest.TestCase):
    def test_invalid_shapes_rejected(self):
        for data in ([], {}, {"contractors": "wrong"}, document({"name": 123}),
                     document({"name": "Demo", "director": []}),
                     document({"name": "Demo", "subcontractors": "wrong"})):
            with self.subTest(data=data), self.assertRaises(ValueError):
                validate_document(data)

    def test_python_cycle_rejected(self):
        child = org()
        child["subcontractors"].append(child)
        with self.assertRaises(ValueError):
            validate_document(document(child))

    def test_embedded_graph_cycle_rejected(self):
        graph = {"nodes": [{"id": "a", "type": "customer", "name": "A"},
                           {"id": "b", "type": "contractor", "name": "B"}],
                 "edges": [{"from": "a", "to": "b"}, {"from": "b", "to": "a"}]}
        with self.assertRaises(ValueError):
            graph_to_tree(graph)

    def test_hallucinated_amount_rejects_llm_output(self):
        with patch("llm.qwen.generate", return_value=json.dumps(document(org(amount="999 млн тг")))):
            self.assertIsNone(extract_llm("Demo City Demo Contractor 000000000000 100 млн тг"))

    def test_supported_llm_output_accepted(self):
        with patch("llm.qwen.generate", return_value=json.dumps(document(org()))):
            self.assertIsNotNone(extract_llm("Demo City Demo Contractor 000000000000 100 млн тг"))

    def test_large_input_is_not_silently_truncated(self):
        with patch("llm.qwen.generate") as generate:
            self.assertIsNone(extract_llm("x" * 7001))
            generate.assert_not_called()


class PrivacyAndEditorTests(unittest.TestCase):
    def test_anonymization_masks_flat_fields_and_removes_canvas(self):
        data = document(org(amount="123 млн", year="2025", description="SECRET DESCRIPTION",
                            _geom={"text": "SECRET"}, director={"fio": "SECRET PERSON"}))
        data["_editor"] = {"blocks": [{"photo": "SECRET PHOTO", "text": "SECRET"}]}
        masked = json.dumps(anonymize_data(copy.deepcopy(data)), ensure_ascii=False)
        for value in ("123", "2025", "SECRET", "_editor", "_geom"):
            self.assertNotIn(value, masked)

    def test_script_breakout_is_escaped(self):
        payload = "</script><script>globalThis.compromised=true</script>"
        html = build_html(document(org(payload)))
        self.assertNotIn(payload, html)
        self.assertIn("\\u003c/script\\u003e", html)

    @unittest.skipUnless(shutil.which("node"), "Node is optional for editor JavaScript tests")
    def test_javascript_parses_and_cycle_detection_finishes(self):
        html = build_html(document(org()))
        script = re.search(r"<script>(.*)</script>", html, re.S).group(1)
        function = script[script.index("function graphDepth("):script.index("function analyze(){")]
        result = subprocess.run(["node", "-e", function +
            ';console.log(JSON.stringify(graphDepth({a:["b"],b:["a"]},["a"])))'],
            capture_output=True, text=True, timeout=5)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue(json.loads(result.stdout)["cycle"])
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "editor.js"
            target.write_text(script, encoding="utf-8")
            checked = subprocess.run(["node", "--check", str(target)], capture_output=True, text=True, timeout=5)
            self.assertEqual(checked.returncode, 0, checked.stderr)

    def test_rebuild_creates_valid_openxml_and_svg(self):
        from main import rebuild_from_json
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "sample.json"
            source.write_text(json.dumps(document(org())), encoding="utf-8")
            rebuild_from_json(str(source))
            with zipfile.ZipFile(Path(folder) / "sample_scheme.xlsx") as archive:
                self.assertIsNone(archive.testzip())
                self.assertIn("xl/worksheets/sheet1.xml", archive.namelist())
            self.assertIn("<svg", (Path(folder) / "sample_graph.svg").read_text(encoding="utf-8"))
            self.assertTrue((Path(folder) / "sample_editor.html").exists())


class SavedCanvasTests(unittest.TestCase):
    def test_unsafe_selector_id_rejected(self):
        with self.assertRaises(ValueError):
            build_html({"_editor": {"blocks": [{"id": 'b1"]><img>', "type": "org",
                                                "role": "con", "lines": []}], "edges": []}})

    def test_saved_canvas_restores_without_tree(self):
        from main import rebuild_from_json
        data = {"_editor": {"blocks": [{"id": "b1", "type": "org", "role": "con",
                                        "x": 1, "y": 1, "lines": []}], "edges": []}}
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "schema_state.json"
            source.write_text(json.dumps(data), encoding="utf-8")
            rebuild_from_json(str(source))
            self.assertTrue((Path(folder) / "schema_state_editor.html").exists())
            with self.assertRaises(ValueError):
                rebuild_from_json(str(source), anonymize=True)


class TransportTests(unittest.TestCase):
    def test_http_error_returns_fallback(self):
        response = Mock()
        response.raise_for_status.side_effect = requests.HTTPError("503")
        with patch("llm.qwen.requests.post", return_value=response):
            self.assertEqual(qwen.generate("demo"), "")

    def test_availability_checks_selected_model(self):
        response = Mock()
        response.json.return_value = {"models": [{"name": "unrelated:latest"}]}
        with patch("llm.qwen.requests.get", return_value=response):
            self.assertFalse(qwen.is_available())

    def test_json_mode_and_text_response(self):
        response = Mock()
        response.json.return_value = {"response": "{}"}
        with patch("llm.qwen.requests.post", return_value=response) as request:
            self.assertEqual(qwen.generate("demo", json_mode=True), "{}")
            self.assertEqual(request.call_args.kwargs["json"]["format"], "json")
