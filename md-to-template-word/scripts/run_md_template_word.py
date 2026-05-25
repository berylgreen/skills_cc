#!/usr/bin/env python
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


def run(cmd: list[str]) -> None:
    subprocess.run(cmd, check=True)


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def detect_template_has_toc(template_path: Path) -> bool:
    with zipfile.ZipFile(template_path, "r") as zf:
        if "word/document.xml" not in zf.namelist():
            return False
        document_xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
    if re.search(r"<w:instrText[^>]*>\s*TOC\s", document_xml):
        return True
    toc_heading = re.search(
        r"<w:p\b[^>]*>\s*(?:<w:pPr\b[^>]*>.*?</w:pPr>)?\s*<w:r\b[^>]*>.*?<w:t>\s*目录\s*</w:t>.*?</w:r>\s*</w:p>",
        document_xml,
        flags=re.S,
    )
    return bool(toc_heading)


def default_output_path(md_path: Path) -> Path:
    return md_path.with_name(f"{md_path.stem}.docx")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert Markdown into a template-based Word document."
    )
    parser.add_argument("--md", required=True, help="Markdown source path")
    parser.add_argument("--template", required=True, help="DOCX template path")
    parser.add_argument("--out", default=None, help="Output DOCX path")
    parser.add_argument(
        "--toc",
        choices=("auto", "on", "off"),
        default="auto",
        help="TOC mode. Default auto = infer from template",
    )
    parser.add_argument(
        "--config",
        default=None,
        help="Optional base JSON config; insert_toc will still be overridden by TOC mode",
    )
    parser.add_argument(
        "--chapter-prefix-pattern",
        default=r"^%1、$",
        help="Validator regex for Heading 1 numbering",
    )
    parser.add_argument(
        "--multilevel-pattern",
        default=r"(%2\.|（%3）)",
        help="Validator regex for Heading 2/3 numbering",
    )
    parser.add_argument(
        "--expected-table-style",
        choices=("any", "three-line", "full-grid"),
        default="three-line",
        help="Render validator expectation for data tables",
    )
    parser.add_argument(
        "--min-table-captions",
        type=int,
        default=None,
        help="Optional lower bound for table captions during render validation",
    )
    parser.add_argument(
        "--skip-finalize",
        action="store_true",
        help="Skip Word field update/finalization",
    )
    parser.add_argument(
        "--skip-validate",
        action="store_true",
        help="Skip structural/render validation",
    )
    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent
    skill_dir = script_dir.parent
    workspace_root = skill_dir.parent.parent.parent
    translator_scripts = workspace_root / ".agents" / "skills" / "docx-template-translator" / "scripts"
    pipeline = script_dir / "adaptive_md_template_pipeline.py"
    md_path = Path(args.md).resolve()
    template_path = Path(args.template).resolve()

    if args.toc == "auto":
        toc_enabled = detect_template_has_toc(template_path)
    else:
        toc_enabled = args.toc == "on"

    out_path = Path(args.out).resolve() if args.out else default_output_path(md_path)

    base_config_path = (
        Path(args.config).resolve()
        if args.config
        else skill_dir / "presets" / ("default.json" if toc_enabled else "no-toc.json")
    )
    config = load_json(base_config_path)
    config["insert_toc"] = toc_enabled

    out_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="md_template_word_") as tmpdir:
        tmpdir_path = Path(tmpdir)
        runtime_config = tmpdir_path / "runtime_config.json"
        runtime_config.write_text(
            json.dumps(config, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        run(
            [
                sys.executable,
                str(pipeline),
                "--template",
                str(template_path),
                "--source-md",
                str(md_path),
                "--out",
                str(out_path),
                "--config",
                str(runtime_config),
            ]
        )

        if toc_enabled:
            run(
                [
                    sys.executable,
                    str(translator_scripts / "inject_toc_field.py"),
                    str(out_path),
                    "--in-place",
                ]
            )

        if not args.skip_finalize:
            run(
                [
                    sys.executable,
                    str(translator_scripts / "finalize_word_docx.py"),
                    str(out_path),
                    "--prefer-dispatch-ex",
                ]
            )

        if not args.skip_validate:
            conversion_json = tmpdir_path / "validation_conversion.json"
            render_json = tmpdir_path / "validation_render.json"

            conversion_cmd = [
                sys.executable,
                str(translator_scripts / "validate_docx_conversion.py"),
                str(out_path),
                "--out",
                str(conversion_json),
                "--no-default-placeholders",
            ]
            run(conversion_cmd)

            render_cmd = [
                sys.executable,
                str(translator_scripts / "validate_docx_render.py"),
                str(out_path),
                "--out",
                str(render_json),
                "--chapter-prefix-pattern",
                args.chapter_prefix_pattern,
                "--multilevel-pattern",
                args.multilevel_pattern,
                "--expected-table-style",
                args.expected_table_style,
            ]
            if not toc_enabled:
                render_cmd.extend(["--allow", "toc-field"])
            if args.min_table_captions is not None:
                render_cmd.extend(["--min-table-captions", str(args.min_table_captions)])
            run(render_cmd)

    print(out_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
