#!/usr/bin/env python3
"""Extract selected numbered questions from a DOCX, preserving its formatting.

Thin CLI over :mod:`paper_core`. The source package's styles, images, equations,
two-column layout, header and footer are all preserved; only the questions you
name are kept, in the order you list them.

Example:
    python extract_questions.py source.docx out.docx 1 5 12 18 140
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from paper_core import extract_questions_by_number


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Extract selected numbered questions from a DOCX, keeping the "
            "original formatting, two-column layout, header and footer."
        )
    )
    parser.add_argument("source", type=Path, help="source .docx file")
    parser.add_argument("target", type=Path, help="output .docx file")
    parser.add_argument(
        "questions",
        metavar="QUESTION",
        type=int,
        nargs="+",
        help="question numbers to retain, in output order (e.g. 1 5 12 18 140)",
    )
    return parser


def _validate(source: Path, target: Path) -> None:
    if not source.is_file():
        raise FileNotFoundError(f"Source file does not exist: {source}")
    if source.suffix.lower() != ".docx":
        raise ValueError(f"Source file must have a .docx extension: {source}")
    if target.suffix.lower() != ".docx":
        raise ValueError(f"Target file must have a .docx extension: {target}")
    if target.parent and not target.parent.exists():
        raise FileNotFoundError(f"Target directory does not exist: {target.parent}")
    if source.resolve() == target.resolve():
        raise ValueError("Source and target must be different files")


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    try:
        _validate(args.source, args.target)
        detected, missing = extract_questions_by_number(
            args.source, args.questions, args.target
        )
    except (OSError, ValueError, RuntimeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    kept = [n for n in dict.fromkeys(args.questions) if n in detected]
    print(f"Created {args.target} ({len(kept)} question(s) kept).")
    if missing:
        print(
            "warning: requested question number(s) not found: "
            + ", ".join(map(str, missing)),
            file=sys.stderr,
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
