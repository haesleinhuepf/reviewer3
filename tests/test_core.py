from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from docx import Document

from reviewer3.core import default_output_path, parse_feedback, review_docx


class StubReviewer:
    def __init__(self) -> None:
        self.calls: list[dict] = []

    def review_paragraph(self, paragraph_text, previous_summaries, introduced_terms_so_far, previous_feedback):
        self.calls.append(
            {
                "paragraph_text": paragraph_text,
                "previous_summaries": list(previous_summaries),
                "introduced_terms_so_far": list(introduced_terms_so_far),
                "previous_feedback": list(previous_feedback),
            }
        )
        if len(self.calls) == 1:
            return parse_feedback(
                '{"summary":"Intro paragraph","explained":["goal"],'
                '"introduced_terms":["LLM"],"comment_question":"Can you define LLM?","suggested_text":null}'
            )
        return parse_feedback(
            '{"summary":"Methods paragraph","explained":["method"],'
            '"introduced_terms":["Ollama"],"comment_question":null,'
            '"suggested_text":"This is a clearer rewrite."}'
        )


class TestCore(unittest.TestCase):
    def test_default_output_path(self):
        self.assertEqual(
            default_output_path(Path("/tmp/paper.docx")),
            Path("/tmp/paper_rv3.docx"),
        )

    def test_review_docx_adds_feedback_and_uses_running_context(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = Path(tmpdir) / "paper.docx"
            output_path = Path(tmpdir) / "paper_rv3.docx"
            doc = Document()
            doc.add_paragraph("Paragraph one about abbreviations.")
            doc.add_paragraph("Paragraph two explains the method.")
            doc.save(str(input_path))

            stub = StubReviewer()
            saved = review_docx(input_docx=input_path, output_docx=output_path, reviewer=stub)

            self.assertEqual(saved, output_path)
            self.assertTrue(output_path.exists())
            self.assertTrue(input_path.exists())
            self.assertEqual(2, len(stub.calls))
            self.assertEqual([], stub.calls[0]["previous_summaries"])
            self.assertEqual(["Intro paragraph"], stub.calls[1]["previous_summaries"])

            reviewed = Document(str(output_path))
            comments = list(reviewed.comments)
            self.assertEqual(1, len(comments))
            self.assertIn("define LLM", comments[0].text)

            xml = reviewed.part.element.xml
            self.assertIn("w:ins", xml)
            self.assertIn("reviewer3 suggestion", xml)


if __name__ == "__main__":
    unittest.main()
