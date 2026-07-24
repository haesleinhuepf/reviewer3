from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openai import OpenAI

REVIEWER_AUTHOR = "reviewer3"
REVIEWER_INITIALS = "rv3"


@dataclass
class ParagraphFeedback:
    summary: str
    explained: list[str]
    introduced_terms: list[str]
    comment_question: str | None
    suggested_text: str | None


class LLMReviewer:
    def __init__(self, client: OpenAI, model: str, temperature: float = 0.2) -> None:
        self._client = client
        self._model = model
        self._temperature = temperature

    @classmethod
    def from_env(cls) -> "LLMReviewer":
        base_url = os.getenv("REVIEWER3_OPENAI_BASE_URL", "http://localhost:11434/v1")
        api_key = os.getenv("REVIEWER3_OPENAI_API_KEY", "ollama")
        model = os.getenv("REVIEWER3_OPENAI_MODEL", "gpt-oss:20b")
        temperature = float(os.getenv("REVIEWER3_OPENAI_TEMPERATURE", "0.2"))
        return cls(client=OpenAI(base_url=base_url, api_key=api_key), model=model, temperature=temperature)

    def review_paragraph(
        self,
        paragraph_text: str,
        previous_summaries: list[str],
        introduced_terms_so_far: list[str],
        previous_feedback: list[str],
    ) -> ParagraphFeedback:
        messages = [
            {
                "role": "system",
                "content": (
                    "You are reviewer3, a constructive reviewer for scientific writing. "
                    "Provide concise, actionable, and respectful feedback. "
                    "If no feedback is needed, explicitly indicate that. "
                    "Return JSON only with keys: summary, explained, introduced_terms, "
                    "comment_question, suggested_text, no_feedback."
                ),
            },
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "paragraph": paragraph_text,
                        "previous_summaries": previous_summaries,
                        "introduced_terms_so_far": introduced_terms_so_far,
                        "previous_feedback": previous_feedback,
                        "instructions": {
                            "summary": "One-sentence summary of current paragraph.",
                            "explained": "List what is explained in this paragraph.",
                            "introduced_terms": "List introduced terms or abbreviations.",
                            "comment_question": "A constructive question if clarification is needed, else null.",
                            "suggested_text": "Improved replacement text only if useful, else null.",
                            "no_feedback": "true if comment_question and suggested_text are both unnecessary.",
                        },
                    },
                    ensure_ascii=False,
                ),
            },
        ]
        response = self._client.chat.completions.create(
            model=self._model,
            temperature=self._temperature,
            messages=messages,
        )
        content = response.choices[0].message.content or "{}"
        return parse_feedback(content, paragraph_text=paragraph_text)


def parse_feedback(raw: str, paragraph_text: str = "") -> ParagraphFeedback:
    parsed = _parse_json_object(raw)
    summary = str(parsed.get("summary") or "").strip()
    explained = _to_str_list(parsed.get("explained"))
    introduced_terms = _to_str_list(parsed.get("introduced_terms") or parsed.get("terms") or parsed.get("abbreviations"))
    comment_question = _to_optional_str(
        parsed.get("comment_question") or parsed.get("comment") or parsed.get("question")
    )
    suggested_text = _to_optional_str(
        parsed.get("suggested_text") or parsed.get("suggestion") or parsed.get("replacement")
    )
    no_feedback = bool(parsed.get("no_feedback"))
    if no_feedback:
        comment_question = None
        suggested_text = None
    if not summary:
        summary = (paragraph_text or "")[:240].strip()
    return ParagraphFeedback(
        summary=summary,
        explained=explained,
        introduced_terms=introduced_terms,
        comment_question=comment_question,
        suggested_text=suggested_text,
    )


def _parse_json_object(raw: str) -> dict[str, Any]:
    text = raw.strip()
    if not text:
        return {}
    try:
        parsed = json.loads(text)
        return parsed if isinstance(parsed, dict) else {}
    except json.JSONDecodeError:
        pass
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        return {}
    try:
        parsed = json.loads(text[start : end + 1])
        return parsed if isinstance(parsed, dict) else {}
    except json.JSONDecodeError:
        return {}


def _to_str_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    text = str(value).strip()
    return [text] if text else []


def _to_optional_str(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def default_output_path(input_docx: Path) -> Path:
    return input_docx.with_name(f"{input_docx.stem}_rv3.docx")


def add_comment_to_paragraph(document: Document, paragraph: Any, comment_text: str) -> None:
    if not comment_text.strip():
        return
    if not paragraph.runs:
        paragraph.add_run(" ")
    document.add_comment(paragraph.runs, text=comment_text, author=REVIEWER_AUTHOR, initials=REVIEWER_INITIALS)


def add_tracked_suggestion(paragraph: Any, suggestion_text: str, change_id: int) -> int:
    suggestion = suggestion_text.strip()
    if not suggestion:
        return change_id
    ins = OxmlElement("w:ins")
    ins.set(qn("w:id"), str(change_id))
    ins.set(qn("w:author"), REVIEWER_AUTHOR)
    ins.set(qn("w:date"), datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"))

    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = f"[reviewer3 suggestion] {suggestion}"
    run.append(text)
    ins.append(run)

    paragraph._p.append(ins)
    return change_id + 1


def review_docx(input_docx: Path, output_docx: Path | None = None, reviewer: Any | None = None) -> Path:
    input_path = Path(input_docx)
    if not input_path.exists():
        raise FileNotFoundError(input_path)
    if input_path.suffix.lower() != ".docx":
        raise ValueError("Input file must be a .docx document.")

    output_path = Path(output_docx) if output_docx else default_output_path(input_path)
    if input_path.resolve() == output_path.resolve():
        raise ValueError("Output path must differ from input path.")
    reviewer = reviewer or LLMReviewer.from_env()

    doc = Document(str(input_path))

    previous_summaries: list[str] = []
    introduced_terms_so_far: list[str] = []
    previous_feedback: list[str] = []
    change_id = 1

    for paragraph in doc.paragraphs:
        paragraph_text = paragraph.text.strip()
        if not paragraph_text:
            continue

        feedback = reviewer.review_paragraph(
            paragraph_text=paragraph_text,
            previous_summaries=previous_summaries,
            introduced_terms_so_far=introduced_terms_so_far,
            previous_feedback=previous_feedback,
        )

        previous_summaries.append(feedback.summary)
        for term in feedback.introduced_terms:
            if term not in introduced_terms_so_far:
                introduced_terms_so_far.append(term)

        feedback_entry = {
            "paragraph": paragraph_text,
            "comment_question": feedback.comment_question,
            "suggested_text": feedback.suggested_text,
        }
        previous_feedback.append(json.dumps(feedback_entry, ensure_ascii=False))

        if feedback.comment_question:
            add_comment_to_paragraph(doc, paragraph, feedback.comment_question)
        if feedback.suggested_text and feedback.suggested_text.strip() != paragraph_text:
            change_id = add_tracked_suggestion(paragraph, feedback.suggested_text, change_id=change_id)

    doc.save(str(output_path))
    return output_path
