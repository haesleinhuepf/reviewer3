from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from openai import OpenAI

REVIEWER_AUTHOR = "reviewer3"
REVIEWER_INITIALS = "rv3"


@dataclass
class FeedbackEntry:
    feedback_type: str
    word_group: str | None
    comment: str | None
    suggestion: str | None


@dataclass
class PrefaceSummary:
    about_document: str
    scientific_importance: str
    global_feedback: str
    feedback_summary: str


class LLMReviewer:
    def __init__(self, client: OpenAI, model: str, temperature: float = 0.2) -> None:
        self._client = client
        self._model = model
        self._temperature = temperature

    @classmethod
    def from_env(cls) -> "LLMReviewer":
        base_url = os.getenv("REVIEWER3_BASE_URL", "http://localhost:11434/v1")
        api_key = os.getenv("REVIEWER3_API_KEY", "ollama")
        model = os.getenv("REVIEWER3_MODEL", "gpt-oss:20b")
        temperature = float(os.getenv("REVIEWER3_TEMPERATURE", "0.2"))
        return cls(client=OpenAI(base_url=base_url, api_key=api_key), model=model, temperature=temperature)

    def review_document(self, document_text: str) -> list[FeedbackEntry]:
        messages = [
            {
                "role": "system",
                "content": (
                    "You are reviewer3, a constructive reviewer for scientific writing. "
                    "Return JSON only. "
                    "Review the full document and return a JSON list called feedback_entries. "
                    "Each feedback entry must include: feedback_type, word_group, comment, suggestion. "
                    "Allowed feedback_type values are: local_modification, regional_question, global_feedback. "
                    "local_modification: use word_group and suggestion (small edits only, such as wording, punctuation, or rephrasing). "
                    "regional_question: use word_group and comment as a constructive question. "
                    "global_feedback: document-level feedback with comment, word_group may be null, suggestion must be null. "
                    "Do not rewrite entire sections if a local edit is enough. "
                    "Use concise, actionable, and respectful feedback. "
                    "For every comment field, write in first-person reviewer voice, for example: 'I read this and now I think ...'. "
                    "Keep this tone friendly, professional, and constructive."
                ),
            },
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "document": document_text,
                        "instructions": {
                            "output_schema": {
                                "feedback_entries": [
                                    {
                                        "feedback_type": "local_modification | regional_question | global_feedback",
                                        "word_group": "exact text span this feedback refers to, or null for global_feedback",
                                        "comment": "question/comment text or null",
                                        "suggestion": "replacement for word_group or null",
                                    }
                                ]
                            }
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
        content = response.choices[0].message.content or "[]"
        return parse_feedback_entries(content)

    def compose_preface(self, document_text: str, feedback_entries: list[FeedbackEntry]) -> PrefaceSummary:
        global_feedback = [entry.comment for entry in feedback_entries if entry.feedback_type == "global_feedback" and entry.comment]
        local_count = sum(1 for entry in feedback_entries if entry.feedback_type == "local_modification")
        regional_count = sum(1 for entry in feedback_entries if entry.feedback_type == "regional_question")

        messages = [
            {
                "role": "system",
                "content": (
                    "You are reviewer3, a constructive reviewer for scientific writing. "
                    "Create a concise preface with exactly the requested sections. "
                    "Return JSON only with keys: about_document, scientific_importance, global_feedback, feedback_summary. "
                    "Write 2-3 sentences for about_document and 2-3 sentences for scientific_importance. "
                    "Use global_feedback for open questions and cross-section coherence points. "
                    "Use feedback_summary for a very short summary of local and regional feedback. "
                    "Write every section in first-person reviewer voice, for example: 'I read this and now I think ...'. "
                    "Keep this tone friendly, professional, and constructive."
                ),
            },
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "document": document_text,
                        "global_feedback_entries": global_feedback,
                        "local_modification_count": local_count,
                        "regional_question_count": regional_count,
                        "instructions": {
                            "about_document": "What the document is about in 2-3 sentences.",
                            "scientific_importance": "Why it is outstanding/important for science or the scientific community in 2-3 sentences.",
                            "global_feedback": "General open questions and points that should be connected across sections.",
                            "feedback_summary": "Very short summary of collected local and regional feedback.",
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
        return parse_preface_summary(
            content,
            fallback_global_feedback=global_feedback,
            fallback_local_count=local_count,
            fallback_regional_count=regional_count,
        )


def parse_preface_summary(
    raw: str,
    fallback_global_feedback: list[str],
    fallback_local_count: int,
    fallback_regional_count: int,
) -> PrefaceSummary:
    parsed = _parse_json_object(raw)
    about_document = _to_optional_str(parsed.get("about_document"))
    scientific_importance = _to_optional_str(parsed.get("scientific_importance"))
    global_feedback = _to_optional_str(parsed.get("global_feedback"))
    feedback_summary = _to_optional_str(parsed.get("feedback_summary"))

    if not about_document:
        about_document = (
            "I read this and now I think the document presents a clear scientific contribution and outlines its methods and findings."
        )
    if not scientific_importance:
        scientific_importance = (
            "I read this and now I think the work is relevant to the scientific community because it clarifies the research context and provides reproducible insights."
        )
    if not global_feedback:
        global_feedback = (
            "I read this and now I think there are no major global issues, but the narrative can be strengthened by linking related concepts across sections."
        )
        if fallback_global_feedback:
            global_feedback = " ".join(fallback_global_feedback)
    if not feedback_summary:
        feedback_summary = (
            "I read this and now I think the collected feedback is focused and actionable: "
            f"{fallback_local_count} local modifications and {fallback_regional_count} regional questions."
        )

    about_document = _ensure_reviewer_voice(about_document)
    scientific_importance = _ensure_reviewer_voice(scientific_importance)
    global_feedback = _ensure_reviewer_voice(global_feedback)
    feedback_summary = _ensure_reviewer_voice(feedback_summary)

    return PrefaceSummary(
        about_document=about_document,
        scientific_importance=scientific_importance,
        global_feedback=global_feedback,
        feedback_summary=feedback_summary,
    )


def parse_feedback_entries(raw: str) -> list[FeedbackEntry]:
    parsed = _parse_json_array_or_wrapped(raw)
    entries: list[FeedbackEntry] = []
    for item in parsed:
        if not isinstance(item, dict):
            continue

        feedback_type = _normalize_feedback_type(item.get("feedback_type") or item.get("type") or item.get("kind"))
        word_group = _to_optional_str(item.get("word_group") or item.get("text_span") or item.get("target"))
        comment = _to_optional_str(item.get("comment") or item.get("question"))
        suggestion = _to_optional_str(item.get("suggestion") or item.get("replacement") or item.get("change_to"))

        if feedback_type == "local_modification":
            if not word_group or not suggestion:
                continue
            entries.append(
                FeedbackEntry(
                    feedback_type=feedback_type,
                    word_group=word_group,
                    comment=comment,
                    suggestion=suggestion,
                )
            )
            continue

        if feedback_type == "regional_question":
            if not word_group or not comment:
                continue
            entries.append(
                FeedbackEntry(
                    feedback_type=feedback_type,
                    word_group=word_group,
                    comment=_ensure_reviewer_voice(comment),
                    suggestion=None,
                )
            )
            continue

        if feedback_type == "global_feedback":
            if not comment:
                comment = suggestion
            if not comment:
                continue
            entries.append(
                FeedbackEntry(
                    feedback_type=feedback_type,
                    word_group=None,
                    comment=_ensure_reviewer_voice(comment),
                    suggestion=None,
                )
            )

    return entries


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


def _parse_json_array_or_wrapped(raw: str) -> list[dict[str, Any]]:
    text = raw.strip()
    if not text:
        return []

    candidates: list[Any] = []
    try:
        candidates.append(json.loads(text))
    except json.JSONDecodeError:
        pass

    for start_char, end_char in (("[", "]"), ("{", "}")):
        start = text.find(start_char)
        end = text.rfind(end_char)
        if start == -1 or end == -1 or end <= start:
            continue
        snippet = text[start : end + 1]
        try:
            candidates.append(json.loads(snippet))
        except json.JSONDecodeError:
            continue

    for candidate in candidates:
        if isinstance(candidate, dict) and isinstance(candidate.get("feedback_entries"), list):
            return [item for item in candidate["feedback_entries"] if isinstance(item, dict)]
        if isinstance(candidate, list):
            return [item for item in candidate if isinstance(item, dict)]
    return []


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


def _ensure_reviewer_voice(text: str) -> str:
    cleaned = text.strip()
    if not cleaned:
        return cleaned
    lower = cleaned.lower()
    if lower.startswith("i ") or lower.startswith("i'"):
        return cleaned
    return f"I read this and now I think {cleaned[:1].lower()}{cleaned[1:]}"


def _normalize_feedback_type(value: Any) -> str:
    text = str(value or "").strip().lower()
    mapping = {
        "local_modification": "local_modification",
        "local": "local_modification",
        "regional_question": "regional_question",
        "regional": "regional_question",
        "question": "regional_question",
        "global_feedback": "global_feedback",
        "global": "global_feedback",
    }
    return mapping.get(text, "")


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

    original = paragraph.text
    if suggestion == original:
        return change_id

    original_tokens = _tokenize_for_diff(original)
    suggestion_tokens = _tokenize_for_diff(suggestion)

    if not original_tokens and not suggestion_tokens:
        return change_id

    _clear_paragraph_content(paragraph)

    matcher = SequenceMatcher(a=original_tokens, b=suggestion_tokens)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            text = "".join(original_tokens[i1:i2])
            _append_plain_text_run(paragraph, text)
            continue
        if tag in {"replace", "delete"}:
            deleted_text = "".join(original_tokens[i1:i2])
            if deleted_text:
                _append_tracked_delete(paragraph, deleted_text, change_id)
                change_id += 1
        if tag in {"replace", "insert"}:
            inserted_text = "".join(suggestion_tokens[j1:j2])
            if inserted_text:
                _append_tracked_insert(paragraph, inserted_text, change_id)
                change_id += 1

    return change_id


def _tokenize_for_diff(text: str) -> list[str]:
    # Keep whitespace and punctuation as separate tokens so punctuation-only edits are tracked precisely.
    return re.findall(r"\w+|[^\w\s]|\s+", text, flags=re.UNICODE)


def _clear_paragraph_content(paragraph: Any) -> None:
    p = paragraph._p
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)


def _append_plain_text_run(paragraph: Any, text_value: str) -> None:
    if not text_value:
        return
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    if text_value[:1].isspace() or text_value[-1:].isspace():
        text.set(qn("xml:space"), "preserve")
    text.text = text_value
    run.append(text)
    paragraph._p.append(run)


def _append_tracked_insert(paragraph: Any, inserted_text: str, change_id: int) -> None:
    ins = OxmlElement("w:ins")
    ins.set(qn("w:id"), str(change_id))
    ins.set(qn("w:author"), REVIEWER_AUTHOR)
    ins.set(qn("w:date"), datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"))

    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    if inserted_text[:1].isspace() or inserted_text[-1:].isspace():
        text.set(qn("xml:space"), "preserve")
    text.text = inserted_text
    run.append(text)
    ins.append(run)

    paragraph._p.append(ins)


def _append_tracked_delete(paragraph: Any, deleted_text: str, change_id: int) -> None:
    deletion = OxmlElement("w:del")
    deletion.set(qn("w:id"), str(change_id))
    deletion.set(qn("w:author"), REVIEWER_AUTHOR)
    deletion.set(qn("w:date"), datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"))

    run = OxmlElement("w:r")
    text = OxmlElement("w:delText")
    if deleted_text[:1].isspace() or deleted_text[-1:].isspace():
        text.set(qn("xml:space"), "preserve")
    text.text = deleted_text
    run.append(text)
    deletion.append(run)

    paragraph._p.append(deletion)


def prepend_review_preface(document: Document, final_summary: PrefaceSummary) -> None:
    _insert_paragraph_at_start(document, "")
    feedback_paragraph = _insert_paragraph_at_start(
        document,
        f"Short feedback summary: {final_summary.feedback_summary.strip()}",
    )
    _ = feedback_paragraph
    text_summary_paragraph = _insert_paragraph_at_start(
        document,
        f"Global feedback: {final_summary.global_feedback.strip()}",
    )
    _ = text_summary_paragraph
    _insert_paragraph_at_start(
        document,
        f"Scientific importance: {final_summary.scientific_importance.strip()}",
    )
    _insert_paragraph_at_start(
        document,
        f"What this document is about: {final_summary.about_document.strip()}",
    )
    _insert_paragraph_at_start(document, "Reviewer3 Preface")


def _insert_paragraph_at_start(document: Document, text: str) -> Paragraph:
    if not document.paragraphs:
        return document.add_paragraph(text)
    first_paragraph = document.paragraphs[0]._p
    new_paragraph_xml = OxmlElement("w:p")
    first_paragraph.addprevious(new_paragraph_xml)
    paragraph = Paragraph(new_paragraph_xml, first_paragraph.getparent())
    paragraph.add_run(text)
    return paragraph


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
    change_id = 1

    candidate_paragraphs = [paragraph for paragraph in doc.paragraphs if paragraph.text.strip()]
    document_text = "\n\n".join(paragraph.text.strip() for paragraph in candidate_paragraphs)

    feedback_entries: list[FeedbackEntry] = []
    if document_text and hasattr(reviewer, "review_document"):
        feedback_entries = reviewer.review_document(document_text=document_text)

    for entry in feedback_entries:
        if entry.feedback_type == "global_feedback":
            continue

        if not entry.word_group:
            continue

        paragraph = _find_paragraph_by_word_group(candidate_paragraphs, entry.word_group)
        if paragraph is None:
            continue

        if entry.feedback_type == "local_modification" and entry.suggestion:
            new_paragraph_text = _replace_word_group_once(paragraph.text, entry.word_group, entry.suggestion)
            if new_paragraph_text.strip() and new_paragraph_text != paragraph.text:
                change_id = add_tracked_suggestion(paragraph, new_paragraph_text, change_id=change_id)

        if entry.feedback_type == "regional_question" and entry.comment:
            add_comment_to_paragraph(doc, paragraph, entry.comment)

    if hasattr(reviewer, "compose_preface"):
        final_summary = reviewer.compose_preface(document_text=document_text, feedback_entries=feedback_entries)
    else:
        global_items = [item.comment for item in feedback_entries if item.feedback_type == "global_feedback" and item.comment]
        local_count = sum(1 for item in feedback_entries if item.feedback_type == "local_modification")
        regional_count = sum(1 for item in feedback_entries if item.feedback_type == "regional_question")
        final_summary = PrefaceSummary(
            about_document=(
                "I read this and now I think this document presents a scientific topic with method and interpretation details."
            ),
            scientific_importance=(
                "I read this and now I think the draft is relevant for the scientific community and can become more impactful with targeted revisions."
            ),
            global_feedback=(
                " ".join(global_items)
                if global_items
                else "I read this and now I think no major global issues were detected, but related ideas can be connected more tightly across sections."
            ),
            feedback_summary=(
                "I read this and now I think the collected feedback is focused and actionable: "
                f"{local_count} local modifications and {regional_count} regional questions."
            ),
        )

    prepend_review_preface(doc, final_summary)

    doc.save(str(output_path))
    return output_path


def _find_paragraph_by_word_group(paragraphs: list[Any], word_group: str) -> Any | None:
    target = word_group.strip()
    if not target:
        return None

    target_lower = target.lower()
    for paragraph in paragraphs:
        paragraph_text = paragraph.text.strip()
        if paragraph_text and target_lower in paragraph_text.lower():
            return paragraph
    return None


def _replace_word_group_once(paragraph_text: str, word_group: str, suggestion: str) -> str:
    if not paragraph_text:
        return paragraph_text

    pattern = re.compile(re.escape(word_group), flags=re.IGNORECASE)
    replaced_text, count = pattern.subn(suggestion, paragraph_text, count=1)
    if count > 0:
        return replaced_text

    if paragraph_text.strip().lower() == word_group.strip().lower():
        return suggestion

    return paragraph_text
