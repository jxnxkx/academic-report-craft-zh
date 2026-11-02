#!/usr/bin/env python3
"""Report likely Typora Markdown convention mistakes without modifying files."""

from __future__ import annotations

import argparse
import re
from pathlib import Path


BACKTICK_MATH = re.compile(r"`([A-Za-z][A-Za-z0-9]*(?:_\{?[A-Za-z0-9+\-]+\}?)+)`")
RAW_SUBSCRIPT = re.compile(r"(?<![`$\\\w])([A-Za-z][A-Za-z0-9]*_[A-Za-z0-9]+)(?![`$\w])")
FENCE = re.compile(r"^\s*(```+|~~~+)(.*)$")


def audit(path: Path) -> int:
    text = path.read_text(encoding="utf-8")
    in_fence = False
    fence_marker = ""
    findings = 0
    display_math_open = False

    for number, line in enumerate(text.splitlines(), 1):
        fence_match = FENCE.match(line)
        if fence_match:
            marker, suffix = fence_match.groups()
            if not in_fence:
                in_fence = True
                fence_marker = marker[0]
                if not suffix.strip():
                    print(f"{path}:{number}: fenced code block has no language")
                    findings += 1
            elif marker[0] == fence_marker:
                in_fence = False
                fence_marker = ""
            continue

        if in_fence:
            continue

        if line.strip() == "$$":
            display_math_open = not display_math_open
            continue

        if display_math_open:
            continue

        for match in BACKTICK_MATH.finditer(line):
            print(
                f"{path}:{number}: review backticked math-like token {match.group(0)}; "
                "use $...$ unless it is a literal field/key"
            )
            findings += 1

        without_inline_code = re.sub(r"`[^`]*`", "", line)
        without_inline_math = re.sub(r"\$[^$]*\$", "", without_inline_code)
        for match in RAW_SUBSCRIPT.finditer(without_inline_math):
            print(
                f"{path}:{number}: raw underscore token {match.group(1)!r}; "
                "use math or code markup according to meaning"
            )
            findings += 1

    if in_fence:
        print(f"{path}: unclosed fenced code block")
        findings += 1
    if display_math_open:
        print(f"{path}: unclosed display-math block")
        findings += 1

    return findings


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("paths", nargs="+", type=Path)
    args = parser.parse_args()

    total = 0
    for path in args.paths:
        total += audit(path)
    print(f"Review candidates: {total}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
