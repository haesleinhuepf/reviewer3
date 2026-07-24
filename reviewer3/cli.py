from __future__ import annotations

import argparse
from pathlib import Path

from reviewer3.core import LLMReviewer, default_output_path, review_docx


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="reviewer3",
        description="Review scientific DOCX text paragraph-by-paragraph using an LLM.",
    )
    parser.add_argument("input_docx", type=Path, help="Path to input .docx file")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Optional output path (default: <input>_rv3.docx)",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    input_docx: Path = args.input_docx
    output_docx: Path = args.output or default_output_path(input_docx)

    if input_docx.resolve() == output_docx.resolve():
        parser.error("Output path must differ from input path.")

    reviewer = LLMReviewer.from_env()
    saved_path = review_docx(input_docx=input_docx, output_docx=output_docx, reviewer=reviewer)
    print(saved_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
