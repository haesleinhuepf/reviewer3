from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from docx import Document

from reviewer3.core import FeedbackEntry, PrefaceSummary, default_output_path, review_docx


class StubReviewer:
    def __init__(self) -> None:
        self.review_calls: list[dict] = []
        self.preface_calls: list[dict] = []

    def review_document(self, document_text):
        self.review_calls.append({"document_text": document_text})
        return [
            FeedbackEntry(
                feedback_type="local_modification",
                word_group="abbreviations",
                comment=None,
                suggestion="abbreviations (LLM)",
            ),
            FeedbackEntry(
                feedback_type="regional_question",
                word_group="explains the method",
                comment="Could you provide one concrete example of the method output?",
                suggestion=None,
            ),
            FeedbackEntry(
                feedback_type="global_feedback",
                word_group=None,
                comment="Some method details and interpretation appear in different sections and should be connected explicitly.",
                suggestion=None,
            ),
        ]

    def compose_preface(self, document_text, feedback_entries):
        self.preface_calls.append(
            {
                "document_text": document_text,
                "feedback_entries": list(feedback_entries),
            }
        )
        return PrefaceSummary(
            about_document="This document introduces the study goals and outlines the applied method.",
            scientific_importance="The topic is valuable because it improves methodological clarity for reproducible research.",
            global_feedback="Open questions remain around how results are interpreted across sections, and related claims should be grouped more tightly.",
            feedback_summary="Local wording and punctuation were improved, and one regional clarification question was raised.",
        )


class TestCore(unittest.TestCase):
    def test_default_output_path(self):
        self.assertEqual(
            default_output_path(Path("/tmp/paper.docx")),
            Path("/tmp/paper_rv3.docx"),
        )

    def test_review_docx_adds_structured_feedback_and_preface(self):
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
            self.assertEqual(1, len(stub.review_calls))
            self.assertEqual(1, len(stub.preface_calls))
            self.assertIn("Paragraph one about abbreviations.", stub.review_calls[0]["document_text"])
            self.assertIn("Paragraph two explains the method.", stub.review_calls[0]["document_text"])

            reviewed = Document(str(output_path))
            self.assertEqual("Reviewer3 Preface", reviewed.paragraphs[0].text)
            self.assertIn("What this document is about:", reviewed.paragraphs[1].text)
            self.assertIn("Scientific importance:", reviewed.paragraphs[2].text)
            self.assertIn("Global feedback:", reviewed.paragraphs[3].text)
            self.assertIn("Short feedback summary:", reviewed.paragraphs[4].text)

            comments = list(reviewed.comments)
            self.assertEqual(1, len(comments))
            self.assertIn("concrete example", comments[0].text)

            xml = reviewed.part.element.xml
            self.assertIn("w:ins", xml)
            self.assertIn("[reviewer3 suggestion]", xml)
            self.assertIn("abbreviations (LLM)", xml)
            self.assertIn("Open questions remain", xml)


if __name__ == "__main__":
    unittest.main()
