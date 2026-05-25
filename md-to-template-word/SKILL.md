---
name: md-to-template-word
description: Use when converting Markdown course notices, assignment sheets, reports, or structured teaching documents into a Word .docx that must follow a supplied template, especially when native Word heading numbering, bullet lists, and template-based table formatting must survive the conversion.
---

# Markdown To Template Word

## Overview

Use this skill when the source of truth is a `.md` file but the deliverable must be a template-based Word document, not a plain pandoc export. The bundled workflow rebuilds Markdown blocks into a template-backed `.docx`, preserves native Word heading numbering, maps unordered lists to `List Paragraph`, follows the template's own TOC or no-TOC structure, and can validate the final render.

## When to Use

- Markdown is the editable source and the final delivery format is `.docx`
- The user provides or names a Word template
- The output must use native Word headings, numbering, bullet lists, or template-defined TOC structure
- A plain `pandoc --reference-doc` export is not enough

Do not use this for PDF-only output, spreadsheet work, or cases where the user wants to hand-edit the Word file without a reproducible conversion flow.

## Files

- Runner: `scripts/run_md_template_word.py`
- Pipeline: `scripts/adaptive_md_template_pipeline.py`
- Presets: `presets/default.json`, `presets/no-toc.json`
- Validation/finalization helpers come from the local `docx-template-translator` skill in the same workspace

## Workflow

1. Identify:
   - source Markdown
   - target template `.docx`
   - output path; default is same-name `.docx`
   - whether the chosen template itself contains a TOC
2. Run the bundled command:

```bash
python .agents/skills/md-to-template-word/scripts/run_md_template_word.py ^
  --md "input.md" ^
  --template "template.docx"
```

3. The runner normally follows the template automatically:

- if the template contains a TOC, it injects and refreshes a real Word TOC field
- if the template does not contain a TOC, it keeps the document as a no-TOC output
- only use `--toc on` or `--toc off` when debugging or deliberately overriding template behavior

4. If the document has known table captions or custom style mapping, pass a JSON config via `--config`. Start from one of the bundled presets and override only the needed keys.

## Markdown Expectations

- `#` first heading becomes the document title
- `## / ### / ####` map to `Heading 1 / 2 / 3`
- Manual numbering in headings may stay in Markdown; the pipeline strips common prefixes so Word numbering can take over
- Unordered lists must use normal Markdown bullets like `- item`
- Pipe tables are converted into Word tables

## Validation

The runner can automatically:

- inject a real TOC field when the template indicates a TOC
- finalize the document through Word field refresh
- run structural and render validation

Use `--skip-finalize` or `--skip-validate` only when debugging the pipeline.
