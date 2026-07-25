from __future__ import annotations

import argparse
import os
import sys
import traceback
from pathlib import Path

from reviewer3.core import LLMReviewer, default_output_path, review_docx


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="reviewer3",
        description="Review a full scientific DOCX document using an LLM and structured feedback entries.",
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
    def progress(message: str) -> None:
        print(message, file=sys.stderr)

    parser = build_parser()
    try:
        args = parser.parse_args()
        input_docx: Path = args.input_docx
        output_docx: Path = args.output or default_output_path(input_docx)

        reviewer = LLMReviewer.from_env(progress_callback=progress)
        saved_path = review_docx(input_docx=input_docx, output_docx=output_docx, reviewer=reviewer)
        print(saved_path)
        return 0
    except KeyboardInterrupt:
        print("Interrupted by user.", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"Error: {type(exc).__name__}: {exc}", file=sys.stderr)
        if os.getenv("REVIEWER3_DEBUG") == "1":
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
