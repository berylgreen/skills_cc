from __future__ import annotations

import argparse
import pathlib
import re


STUDENT_ID_RE = re.compile(r"(124\d{9})")


def normalize_name(text: str) -> str:
    text = re.sub(r"[0-9\s,，:：;；|/\\]+", "", text)
    text = re.sub(r"[（(].*?[)）]", "", text)
    return text.strip()


def extract_pairs(text: str) -> list[tuple[str, str]]:
    rows: list[tuple[str, str]] = []
    seen: set[str] = set()

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        match = STUDENT_ID_RE.search(line)
        if not match:
            continue

        student_id = match.group(1)
        if student_id in seen:
            continue

        remainder = line.replace(student_id, " ", 1)
        name = normalize_name(remainder)
        if not name:
            continue

        seen.add(student_id)
        rows.append((student_id, name))

    return rows


def main() -> int:
    parser = argparse.ArgumentParser(description="Normalize WeChat mini program student text into 学号 姓名.")
    parser.add_argument("--input", type=pathlib.Path, required=True, help="Raw text file from OCR or copied list.")
    parser.add_argument("--output", type=pathlib.Path, required=True, help="Normalized output text file.")
    args = parser.parse_args()

    text = args.input.read_text(encoding="utf-8-sig")
    rows = extract_pairs(text)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        "\n".join(f"{student_id} {name}" for student_id, name in rows) + ("\n" if rows else ""),
        encoding="utf-8",
    )
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
