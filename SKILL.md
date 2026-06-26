---
name: chinese-word-pro
description: Use when creating, editing, formatting, or verifying Chinese Word `.docx` files, especially when Chinese filenames, Chinese body text, USTC/中国科学技术大学 thesis formatting, thesis/report formatting, font sizes, tables, references, or anti-garbled UTF-8 workflows matter. Enforces UTF-8 input, East Asian font settings, polished Word typography, USTC thesis standards, and render-based QA.
---

# Chinese Word Pro

Use this skill for Chinese `.docx` work. It complements the built-in `documents` skill: use this skill for Chinese encoding, typography, and workflow rules, then use `documents` for DOCX rendering and visual QA.

## USTC Thesis Standard

When the user asks for a thesis, dissertation, graduation thesis, USTC/中国科学技术大学 format, school official template, or says Word output should follow the official thesis template, read `references/ustc_thesis_format.md` before creating or formatting the document.

Use `assets/ustc_phd_thesis_template.docx` as the reference asset cloned from the user's official template:

`F:\硕士\毕业相关\中国科技大学博士学位论文撰写模板（新）.doc`

For USTC thesis-style output, the USTC rules override the generic defaults below.

## Non-negotiables

1. Never pass Chinese body text through shell pipelines such as `echo 中文 | python ...` or PowerShell redirection.
2. Put Chinese content in a UTF-8 file, preferably JSON or Markdown, and read it with explicit encoding.
3. Use `pathlib.Path` for Chinese file names and paths.
4. In `python-docx`, set both Latin and East Asian fonts. `run.font.name` alone is not enough.
5. Before delivery, inspect the DOCX XML for replacement characters or long runs of question marks.
6. For final DOCX delivery, render to page PNGs with the `documents` skill whenever LibreOffice/Word rendering is available, then visually inspect and iterate.
7. If rendering is blocked by missing software or sandbox permissions, say so plainly and do not claim visual QA passed.
8. If the manuscript uses Zotero or another Word citation-field workflow, DOCX post-processing must not break the citation fields.
9. Never overwrite the user-facing final DOCX directly from a temporary export before structural checks are complete.

## Preferred Build Pattern

Use the bundled script first when starting a new Chinese Word document:

```powershell
& "<bundled-python>" "<skill-dir>\scripts\build_chinese_word.py" --input "<utf8-content.json>" --output "<中文输出.docx>"
```

When writing custom code, keep these rules:

```python
from pathlib import Path
from docx.oxml.ns import qn

text = Path(input_path).read_text(encoding="utf-8-sig")

run.font.name = "Times New Roman"
run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "宋体")
```

For generated JSON, use:

```python
json.dumps(data, ensure_ascii=False)
```

## Default Chinese Document Style

- Page: A4-like defaults, top/bottom 2.54 cm, left 3.0 cm, right 2.7 cm.
- Body: Chinese `宋体`, English `Times New Roman`, 10.5 pt, 1.5 line spacing, first-line indent 2 Chinese characters.
- Heading 1: `黑体`, 15 pt, bold, restrained blue accent, keep with next.
- Heading 2: `黑体`, 13 pt, bold, keep with next.
- References: hanging indent, 10 pt, 1.25 line spacing.
- Tables: intentional column widths, repeated header row, shaded header, cell padding, vertical centering, no fixed row heights.

Adjust these defaults when the user provides a school template, journal style, official thesis format, or explicit font/size requirements.

For USTC thesis documents, do not use these generic report defaults. Use the USTC reference instead:

- A4; margins top/bottom 2.54 cm, left/right 3.17 cm.
- Body `宋体` + `Times New Roman`, 小四 12 pt, fixed 20 pt line spacing, first-line indent 2 Chinese characters.
- Chapter title `黑体` 三号 centered, 24 pt before, 18 pt after.
- Level-1 heading `黑体` 四号 14 pt, fixed 20 pt, 24 pt before, 6 pt after.
- Level-2 heading `黑体` 13 pt, fixed 20 pt, 12 pt before, 6 pt after.
- References 五号 10.5 pt, fixed 20 pt, hanging indent 2.5 Chinese characters.

## Thesis Word Delivery Rules

Use these rules by default when the task is a Chinese thesis, dissertation chapter, thesis-format report, table/figure-heavy academic DOCX, or any `.docx` revision that mentions template formatting, three-line tables, clipped figures, or broken formula display.

### Trigger Patterns

- Chinese thesis/dissertation output
- chapter revision in `.docx`
- template-based formatting
- 三线表 / three-line table
- figure insertion or side-by-side figures
- formula, subscript, Greek-letter, or inline-math corruption in Word

### Required Workflow

1. Edit the UTF-8 source first. Do not treat the generated `.docx` as the only source of truth.
2. Generate a temporary DOCX with the project-approved export chain.
3. If the project uses citation fields, verify that the temporary DOCX still contains them before post-processing.
4. Run DOCX post-processing to repair Word-structure issues such as figure sizing, table borders, paragraph indentation, and inline symbol rendering.
5. If the project uses citation fields, verify again that post-processing preserved them.
6. If rendering is available, run visual QA on page images/PDF before delivery.
7. If rendering is not available, explicitly disclose that structural fixes were applied but automatic page-image QA was not completed.
8. Only overwrite the final user-facing DOCX after the temporary DOCX has passed the structural checks and, when applicable, the citation-field checks.

### Hard Rules

- `Markdown -> pandoc -> docx` is a draft-generation path, not a complete final-delivery path.
- For Chinese academic DOCX, separate content correctness from Word-structure correctness.
- For citation-managed manuscripts, separate citation-field correctness from visual formatting correctness and validate both.
- Do not assume the reference template will automatically fix three-line tables, figure sizing, inline symbols, or paragraph indentation.
- Always inspect table paragraphs for `first_line_indent`, `left_indent`, and `right_indent`; academic tables should normally reset all of them to zero unless the user explicitly requests otherwise.
- When inline formulas or symbol explanations are prone to corruption in Word, prefer native Word runs with explicit subscript formatting over trusting automatic math conversion.
- Before overwriting a user-facing DOCX, write a temporary output first so an open file or a failed pass does not destroy the working draft.
- If post-processing removes or flattens citation fields, treat that as a delivery failure, not a minor defect.

### Default Figure Standards

- Prefer inline figures rather than floating anchors unless the user explicitly needs wrap-around.
- Side-by-side figures must be scaled to a safe page width so no image exceeds its cell width or page margins.
- Captions must remain immediately adjacent to the figure block.
- If a figure risks clipping in Word, reduce figure width or adjust the containing layout; do not force it to fill the page.
- For Chinese academic work, verified PNG is preferred over SVG when Word rendering is unstable.

### Default Three-Line Table Standards

- Default three-line table widths:
  - top rule: `1.5 pt`
  - header separator rule: `0.5 pt`
  - bottom rule: `1.5 pt`
- By default, do not keep vertical borders and do not keep ordinary row gridlines unless the user explicitly requests auxiliary lines.
- Table header cells should be centered.
- Short body fields should usually be centered.
- Long narrative or definition columns should usually be left-aligned.
- Prefer compact academic geometry:
  - fixed column widths
  - small cell margins
  - zero paragraph indentation
  - explicit line spacing
  - no fixed row heights that can clip text

### Formula and Symbol Repair Standards

When these symbols appear in body-text explanations, prefer native Word runs with subscript rather than relying entirely on pandoc math conversion:

- `μ_i`
- `λ_t`
- `β_1`
- `β_2`
- `K_it`
- `PR_kt`
- `N_ikt`

If the generated DOCX shows degraded forms like `(_i)` or missing Greek letters, treat that as a required post-processing fix, not a cosmetic issue.

### Recommended Companion Script

For non-trivial thesis DOCX delivery, use `scripts/postprocess_thesis_docx.py` after draft generation to normalize:

- side-by-side figure sizing
- three-line table borders
- zero-indentation table paragraphs
- compact table spacing
- inline subscript and Greek-letter repair in explanation paragraphs


## Garbling Check

After saving a `.docx`, inspect `word/document.xml` inside the ZIP:

- Fail if expected Chinese strings are missing.
- Fail if XML contains `�`.
- Fail if XML contains suspicious long runs like `????`.

If the XML already has question marks, the content was damaged before Word generation. Go back to the input encoding path rather than changing Word fonts.

## Citation-Field Protection Check

When the manuscript uses Zotero or another Word citation-field workflow, inspect `word/document.xml` inside the ZIP before and after post-processing.

Minimum pass conditions:

- citation-field markers such as `ADDIN ZOTERO_ITEM` or `CSL_CITATION` are still present
- no post-processing step flattened those markers into plain-text references

If this check fails, do not deliver that DOCX as the final manuscript.

## Render QA

Use `documents/render_docx.py` to render pages to PNGs and inspect them. Check:

- no `????` or tofu glyph boxes
- Chinese and English mixed text uses the intended fonts
- headings, body text, tables, and references have consistent spacing
- table cells are not cramped or clipped
- page breaks do not create awkward large blanks

Repeat edits and rendering until the output is visually clean.

## When The User Wants “Perfect Word Output”

Treat “perfect” as a workflow, not a single library:

1. Start from UTF-8 content files.
2. Generate with explicit Word styles.
3. Validate DOCX XML for encoding damage.
4. Validate citation fields when the project uses them.
5. Render to PNG/PDF.
6. Visually inspect and iterate.
7. Deliver only the final `.docx` unless the user asks for QA artifacts.
