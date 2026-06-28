---
name: chinese-word-pro
description: Use when creating, editing, formatting, or verifying Chinese Word `.docx` files, especially when Chinese filenames, Chinese body text, USTC/中国科学技术大学 thesis formatting, thesis/report formatting, font sizes, tables, references, or anti-garbled UTF-8 workflows matter. Enforces UTF-8 input, East Asian font settings, polished Word typography, USTC thesis standards, and render-based QA.
---

# Chinese Word Pro

Use this skill for Chinese `.docx` work and bilingual academic submission finalization. It complements the built-in `documents` skill: use this skill for encoding, typography, Word-structure repair, and delivery-safe post-processing, then use `documents` for DOCX rendering and visual QA.

## General Chinese Report Standard

When the user asks for a Chinese report, stock analysis report, course report, case analysis, management report, or any non-thesis Chinese Word/PDF deliverable with charts or tables, read `references/general_report_delivery.md` before creating or formatting the document.

Use `assets/report_template.docx` as the default reference template for ordinary Chinese reports unless the user provides a more specific template.

For ordinary Chinese report output, the report delivery rules override generic defaults where they conflict:

- Title and headings must be black unless the user explicitly requests accent colors.
- Remove template residue such as blue title borders, underline styles, and theme accent colors.
- Figures must be inline and visually verified; when clipping or text-overlap risk exists, use a borderless one-column figure block.
- Table captions go above tables; figure captions go below figures.
- Three-line tables must be enforced with cell-level borders.
- Report tables should stay compact and structural. Do not put long narrative judgment columns such as `分析判断`, `原因说明`, or `投资含义` into tables by default; move that analysis into paragraphs immediately below the table.
- OOXML cleanup must include `word/document.xml`, `word/styles.xml`, and `word/theme/theme1.xml`. Theme files must be processed before generic `word/*.xml` cleanup, otherwise `theme1.xml` may keep blue `accent1`/`4F81BD` residues.
- PDF export on macOS/LibreOffice must use `scripts/export_pdf_fixed.py` or an equivalent fontconfig-safe path, then render pages for QA.

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
10. For ordinary Chinese reports, do not rely on a template alone to fix figures, three-line tables, title borders, or PDF fonts; run structural post-processing and render QA.
11. For table-heavy reports, keep tables for metrics/categories only; if a cell needs sentence-level analysis, remove that column and write the interpretation as body text below the table.
12. When cleaning blue heading residue, inspect not only visible text but also `word/styles.xml` and `word/theme/theme1.xml`; visual pass alone is insufficient because Word/LibreOffice can reinterpret theme accents on another machine.

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

For ordinary Chinese report PDF export on macOS/LibreOffice, use the bundled fixed exporter:

```bash
python3 "<skill-dir>/scripts/export_pdf_fixed.py" \
  --input-docx "<final.docx>" \
  --output-pdf "<final.pdf>" \
  --render-dir "<rendered_pages>"
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

## Academic Submission Finalization

Use this module when the task is not only "make the Word look better" but "produce the formal submission-safe Chinese and English final DOCX pair."

Core authority:

- bilingual docx finalization
- journal-style academic tables
- figure-caption normalization
- native Word formula repair
- chapter pagination
- citation-safe post-processing

This skill is the Word finalization authority. It owns the final post-processing rules for:

- Chinese and English body-font normalization
- left-aligned academic paragraph layout by default
- three-line table geometry
- inline figure normalization
- figure-caption separation and numbering preservation
- native equation object conversion
- inline pseudo-formula repair in explanation paragraphs
- first-page and chapter pagination rules
- XML-level garbling checks
- citation-field preservation checks

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
- Do not directly hand-edit the previous "healthy" final DOCX and treat that as the new delivery baseline; always re-export from Markdown, then re-run finalization.
- Chinese and English final DOCX outputs must be finalized with the same structural rule set unless the user explicitly requests divergence.
- Academic submission paragraphs must be left-aligned by default. Do not deliver body text, headings, figure captions, table captions, references, or table-cell text as centered or justified unless the user explicitly requests that style. The only default exceptions are pure figure/image paragraphs, native equation blocks, and equation-number paragraphs.

### Mandatory Delivery Audit

Before a final DOCX is allowed to overwrite the main deliverable, the post-processing workflow must confirm all of the following:

- live citation fields still exist in `word/document.xml`
- `m:oMath` or `m:eqArr` objects exist when model equations are present
- explanation paragraphs no longer expose broken pseudo-formula forms such as `Y_it`, `CR_it`, or `z(...)` as raw degraded text
- figures are inline rather than floating anchors
- figure captions are independent paragraphs and retain numbering
- three-line tables remain structurally intact
- body text, headings, captions, references, and table-cell paragraphs are left-aligned unless explicitly overridden
- chapter-opening titles, abstract, and references use page-break-before where required
- temporary DOCX exports are cleaned up after delivery

### Default Figure Standards

- Prefer inline figures rather than floating anchors unless the user explicitly needs wrap-around.
- Side-by-side figures must be scaled to a safe page width so no image exceeds its cell width or page margins.
- Captions must remain immediately adjacent to the figure block.
- If a figure risks clipping in Word, reduce figure width or adjust the containing layout; do not force it to fill the page.
- For Chinese academic work, verified PNG is preferred over SVG when Word rendering is unstable.
- Figure paragraphs must not inherit ordinary body-text first-line indentation or fixed body line spacing.
- Figure captions must be separate paragraphs, not embedded in the interpretive paragraph above the figure.
- Figure captions must be left-aligned in submission finalization unless the user explicitly requests centered captions.
- Preferred caption formats:
  - Chinese: `图 1 标题`
  - English: `Figure 1. Title`

### Default Three-Line Table Standards

- Default three-line table widths:
  - top rule: `1.5 pt`
  - header separator rule: `0.5 pt`
  - bottom rule: `1.5 pt`
- By default, do not keep vertical borders and do not keep ordinary row gridlines unless the user explicitly requests auxiliary lines.
- Table header cells should be left-aligned by default for submission finalization unless the target template explicitly requires centering.
- Short body fields should usually be left-aligned when the user requests globally left-aligned paragraphs.
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

For journal-submission finalization:

- core model equations should be converted to native Word math objects
- multiline equations should remain a single math block rather than fragmented line-by-line text boxes
- equation numbers should be independent right-aligned paragraphs
- inline pseudo-formulas in explanation text should be repaired to true subscript or superscript runs
- Greek letters, subscripts, and squared terms must not be delivered as bare underscore strings

### Recommended Companion Scripts

For non-trivial thesis DOCX delivery, `scripts/postprocess_thesis_docx.py` remains available for legacy thesis-style repairs.

For bilingual journal-submission finalization, use:

`scripts/finalize_submission_docx.py`

Recommended interface:

```bash
python3 finalize_submission_docx.py \
  --input-docx temp_export.docx \
  --output-docx final.docx \
  --lang cn \
  --mode journal_submission
```

This companion script is the preferred execution point for:

- font and paragraph normalization
- equation-object repair
- pseudo-formula explanation repair
- figure-paragraph and figure-caption normalization
- three-line table repair
- chapter pagination
- garbling checks
- citation-field protection checks


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
