---
name: chinese-word-pro
description: Use when creating, editing, formatting, or verifying Chinese Word `.docx` files, especially when Chinese filenames, Chinese body text, USTC/中国科学技术大学 thesis formatting, thesis/report formatting, font sizes, tables, references, or anti-garbled UTF-8 workflows matter. Enforces UTF-8 input, East Asian font settings, polished Word typography, USTC thesis standards, and render-based QA.
---

# Chinese Word Pro

Use this skill for Chinese `.docx` work and bilingual academic submission finalization. It complements the built-in `documents` skill: use this skill for encoding, typography, Word-structure repair, and delivery-safe post-processing, then use `documents` for DOCX rendering and visual QA.

It also owns Word-only academic refinement after a manuscript has matured beyond Markdown-first rebuilding. In that stage, Chinese and English Word files are the active manuscripts; Markdown is archival unless the user explicitly returns to a full rebuild.

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

For ordinary Chinese report theme cleanup before PDF export or delivery, run:

```bash
python3 "<skill-dir>/scripts/clean_report_theme.py" \
  --input-docx "<temp.docx>" \
  --output-docx "<temp_clean.docx>"
```

This cleaner is the default defense against hidden blue heading residue from Word base templates. It audits and repairs `word/document.xml`, `word/styles.xml`, and `word/theme/theme1.xml` together instead of only changing visible run colors.

## Default Chinese Document Style

- Page: A4-like defaults, top/bottom 2.54 cm, left 3.0 cm, right 2.7 cm.
- Body: Chinese `宋体`, English `Times New Roman`, 10.5 pt, 1.5 line spacing, first-line indent 2 Chinese characters.
- Heading 1: `黑体`, 15 pt, bold, restrained blue accent, keep with next.
- Heading 2: `黑体`, 13 pt, bold, keep with next.
- References: hanging indent, 10 pt, 1.25 line spacing.
- Tables: intentional column widths, repeated header row, shaded header, cell padding, vertical centering, no fixed row heights.
- Academic submission tables: centered, fit the available page width, no inherited body first-line indentation inside cells, and centered cell text by default unless the journal profile explicitly requires field-specific alignment.

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

### Academic Word Finalization Triad

Formal academic Word delivery must pass three independent hard gates. These gates are general and cross-project; they are not tied to any single manuscript, journal, variable system, or language pair.

1. `Formula Gate`
   - Display equations must use native Word math objects.
   - The visible formula body must be centered.
   - The equation number must stay on the same formula paragraph and be right-aligned.
   - Inline symbolic explanations must use Word-native subscript/superscript runs rather than raw pseudo-LaTeX or collapsed text.

2. `Figure Gate`
   - Figures must be inline, not floating anchors.
   - Figure aspect ratio must be preserved; do not stretch an image just to fill blank page space.
   - Figure paragraphs must be centered with zero indentation and safe line spacing.
   - Figure captions must be independent centered paragraphs with continuous numbering.

3. `Paragraph and Table Geometry Gate`
   - Ordinary body paragraphs are left-aligned by default.
   - Figure captions and table captions are centered by default.
   - Academic table-cell paragraphs are centered by default and must explicitly reset first-line, left, and right indentation to zero.
   - Table cells must be vertically centered, avoid fixed row-height clipping, and fit the available page width unless a target-journal profile requires another layout.
   - The finalizer must neutralize style-inherited Chinese first-line indentation at OOXML level, including `w:firstLineChars="0"`, not merely through visible style settings.

If any gate fails, the DOCX is not formal-delivery-safe and must not overwrite the main manuscript file.

### Word-only Academic Refinement Mode

Use this mode when the user confirms the manuscript has entered Word refinement, Word final polish, submission polish, or when the manuscript is structurally stable and repeated Markdown-to-Word rebuilds would risk damaging already repaired formulas, figures, tables, citation fields, or pagination.

Purpose:

- preserve a mature Word manuscript while making targeted, auditable improvements
- avoid unnecessary full Markdown-to-Word re-export after layout has stabilized
- keep Chinese and English Word deliverables synchronized without treating Markdown as the live source

Allowed operations:

- targeted formula repair and equation-number alignment
- table-cell indentation removal, centering, width fitting, and three-line-table repair
- figure replacement, sizing, aspect-ratio preservation, and caption normalization
- paragraph alignment, cover-page, abstract, keyword, chapter-pagination, and reference-page polishing
- citation-field-safe local wording edits
- bilingual Word synchronization, where the English Word is a faithful translation-equivalent counterpart of the Chinese Word

Prohibited operations unless the user explicitly requests a full rebuild:

- regenerating the active final Word files from Markdown for a small local edit
- flattening or replacing live citation fields with plain text
- broad style resets that overwrite previously verified equation, table, or figure geometry
- using ordinary visible tables as the default equation-numbering solution
- directly overwriting the active Word file before a temporary output and audit pass exist

Required workflow:

1. Identify the active Chinese and English Word files and confirm the task is localized or final-polish oriented.
2. Write changes to temporary DOCX files first; never modify the active deliverables in place as the first write.
3. Preserve `ADDIN ZOTERO`, `CSL_CITATION`, or other live citation markers when present.
4. Apply the smallest necessary OOXML or `python-docx` changes rather than a whole-document style reset.
5. Run structural audits for the touched feature: citation fields, equations, captions, inline figures, table indentation, or paragraph geometry.
6. Render or visually inspect the touched pages when available; XML success alone is not enough for formulas and figures.
7. Replace the active Word deliverables only after the temporary files pass.
8. If the project is bilingual, mirror the same substantive change into the other language's Word file and audit both.
9. Commit and push when the project rules require versioned delivery.

Mode boundary:

- If the requested change adds a new section, restructures the argument, changes tables or empirical results wholesale, or requires rebuilding the citation architecture, ask whether to return temporarily to Markdown-first drafting.
- If the requested change is citation enrichment, small wording revision, equation repair, figure/table formatting, or final layout polish, stay in Word-only refinement by default.

Markdown handling:

- Markdown files are archival snapshots in this mode.
- Do not block a Word-only refinement delivery merely because Markdown was not updated.
- If the user asks for archival synchronization after the Word polish, update Markdown separately from the stable Word file and clearly label it as archival, not as the layout source.

### Formula Finalization Pipeline

For academic manuscripts, formula handling is a first-class delivery subsystem rather than an afterthought.

The pipeline is mandatory:

1. `Formula Inventory`
2. `Inline Symbol Renderer`
3. `OMML Equation Compiler`
4. `Formula Delivery Audit`

The purpose is to prevent the common failure mode in which a document looks acceptable in one project, but breaks as soon as a new equation, a longer model specification, or a new explanatory paragraph is added in another project.

### Formula Inventory

Before final overwrite, the workflow must treat all formula-bearing content as an audited set, not as ad hoc strings.

The inventory must cover at least:

- numbered empirical models
- measurement formulas
- weighting, network, standardization, and interaction formulas
- long or multiline equations
- explanation paragraphs immediately below formulas
- repeated coefficient symbols, Greek letters, and indexed variables used again in prose

The implementation may infer this inventory automatically from the DOCX and source text, but formal delivery must behave as though this inventory exists and has been checked.

### Inline Symbol Renderer

The inline symbol renderer is responsible for formula-like content that appears in ordinary prose rather than display-equation blocks.

Its default job is to convert raw or collapsed notation into Word-native subscript or superscript runs, including:

- braced forms such as `K_{it}`, `PR_{kt}`, `N_{ikt}`, `L_{jt}`, `Return_{im}`
- simple underscore forms such as `Y_it`, `CR_it`, `AIW_it`, `Outcome_it`
- collapsed residue such as `Resilienceit`, `AIWit`, `Controlsit`, `Outcomeit`, `w1`, `w2`, `μi`, `λt`, `εit`
- squared-term expressions in explanatory prose such as `AIW²`

Hard rule:

- any symbol that appears in a formal equation and is repeated in explanation prose must be rendered as readable Word-native notation in that prose as well; it must not be delivered as raw underscore text, collapsed plain text, or pseudo-LaTeX residue

### OMML Equation Compiler

The OMML equation compiler is responsible for numbered display equations.

Highest-priority visual contract:

- The final visible layout must match the stable Word-native pattern: a centered native equation block with its equation number on the same numbered paragraph, aligned at the right margin.
- Treat the equation body as the "original formula block" rendered as native Word OMML, not as a picture, text box, or ordinary prose string.
- Use a center tab stop for the equation body and a right tab stop for the number. The paragraph may be left-aligned internally to support tab stops, but the visible equation body must be centered and the visible number must be right-aligned.
- For short and medium equations, keep one visual row.
- For long equations, first try formula-specific fit controls such as modest math-run font reduction and tab-stop tuning. If wrapping is unavoidable, keep the formula as native math and keep the number on the final visual row at the right margin.
- Do not use visible or ordinary Word tables as the default equation layout strategy. A table-like equation container is a last-resort recovery device only after explicit user approval or documented journal-template necessity.
- If render QA shows the equation number below the formula, centered under the formula, left-aligned, or separated from the formula paragraph, the formula audit fails even if the XML contains `m:oMath`.

Default policy:

- convert display equations to native Word math objects
- prefer one numbered equation paragraph rather than disconnected equation and number paragraphs
- use a right-aligned tab stop for numbering by default
- keep equation number on the same visual row as the equation
- for short equations, keep a single formula row
- for long equations, keep one equation paragraph but automatically wrap the formula into multiple visual lines inside that paragraph instead of reverting to ordinary text
- if a previously exported file contains a formula paragraph plus a separate orphan number paragraph, merge them back into one numbered equation paragraph and delete the orphan number paragraph

The default general-purpose strategy is therefore:

- `native OMML equation`
- `same numbered paragraph`
- `right-aligned tab stop`
- `formula-specific fitting or controlled native wrapping for overlong equations`

Visible table fallbacks are not the preferred solution for formal delivery.

### Formula Delivery Audit

Formal delivery must fail closed when formula structure is unsafe.

The formula audit must verify all of the following:

- display equations that should be native math contain `m:oMath` or `m:eqArr`
- no orphan equation-number-only paragraphs remain after finalization
- equation numbers stay on the same numbered equation paragraph and are right-aligned through tab-stop layout
- tab-stop layout must be audited structurally: each normal numbered equation paragraph should contain a leading tab before the native OMML equation object and a second tab before the equation number; the audit must treat an `m:oMath` or `m:oMathPara` element itself as a math object, not only math descendants
- long equations with continuation lines use row-appropriate tab-stop logic: normal formula rows may use center + right tab stops, but short continuation-tail rows with the equation number should use a right tab stop only so the number does not stop at the center tab
- long equations are wrapped inside the equation paragraph rather than overflowing the page as one unbroken line
- explanation paragraphs do not expose raw source-like strings such as `Y_it`, `CR_it`, `Return_{im}`, `K_{it}`, `PR_{kt}`, `N_{ikt}`, `L_{jt}`, or collapsed residue such as `Resilienceit`, `Outcomeit`, `w1`, `w2`
- Greek letters, indexed coefficients, and disturbance terms used in prose are delivered with true subscript formatting

If any of these checks fail, the DOCX is not formal-delivery-safe and must not overwrite the main manuscript file.

### Renderer Artifact Triage

Some apparent formula defects are renderer-specific rather than true Word-content defects. For example, a headless LibreOffice render may show a stray leading symbol before a native equation even when WPS or Microsoft Word displays the formula correctly.

Do not automatically rewrite or simplify a formula solely because one renderer shows an isolated special symbol before a formula.

Triage rule:

- first check the DOCX structure for native `m:oMath` or `m:eqArr`
- render with the available automated renderer
- if the issue is an isolated leading or decorative artifact and the formula XML is otherwise structurally sound, ask the user to confirm in WPS/Microsoft Word before changing formula content
- record the artifact in the delivery log or QA notes
- only modify the formula in the next round if the user confirms that the artifact appears in the target editor or final submission renderer

This rule prevents unnecessary formula rewrites that may damage an otherwise correct native Word equation.

When Zotero live citation fields are required, this skill also provides the downstream recovery helper for citation-manager readiness before Word export:

- locate and open Zotero if closed
- open the target Zotero collection when a collection key or select URI is available
- verify Zotero local connector availability
- inspect a smoke-test DOCX for live Zotero citation markers when required

If the helper cannot recover Zotero, use Computer Use for one GUI attempt to select the target collection. If Zotero, Better BibTeX, or the approved MCP route still cannot create live fields, stop and report failure rather than producing a flattened Word file.

### Required Workflow

1. In Markdown-first mode, edit the UTF-8 source first. In Word-only refinement mode, treat the active DOCX pair as the working manuscripts and do not force a Markdown rebuild for localized polish.
2. Generate a temporary DOCX with the project-approved export chain in Markdown-first mode, or create temporary DOCX working copies from the active Word files in Word-only refinement mode.
3. If the project uses citation fields, verify that the temporary DOCX still contains them before post-processing.
4. Run DOCX post-processing to repair Word-structure issues such as figure sizing, table borders, paragraph indentation, and inline symbol rendering.
4a. For ordinary Chinese reports, run `scripts/clean_report_theme.py` before render QA so `Title`/`Heading` theme residue does not survive into another machine's Word/WPS rendering.
5. If the project uses citation fields, verify again that post-processing preserved them.
6. If rendering is available, run visual QA on page images/PDF before delivery.
7. If rendering is not available, explicitly disclose that structural fixes were applied but automatic page-image QA was not completed.
8. Only overwrite the final user-facing DOCX after the temporary DOCX has passed the structural checks and, when applicable, the citation-field checks.

### Delivery Mode Distinction

This skill must distinguish formal submission delivery from recovery or working-draft layout repair.

Formal submission mode:

- Use `--citation-policy strict`.
- Require live citation fields when the manuscript is citation-managed.
- Fail if `ADDIN ZOTERO_ITEM` or `CSL_CITATION` markers are absent.
- Do not overwrite the formal final Word files unless citation, garbling, formula, figure, table, and render checks pass.

Recovery or working-draft layout mode:

- Use `--citation-policy warn` only when the user explicitly needs a readable restored Word file before the citation-aware export chain is repaired.
- The script may complete typography, equation, figure, table, and pagination repair.
- The script must print or log that live citation fields are absent.
- The output must be described as a recovery draft or working draft, not a formal submission file.

Non-citation-managed mode:

- Use `--citation-policy off` only for documents that are not expected to contain Zotero, CSL, or equivalent live fields.
- Do not use `off` for empirical manuscripts with citekeys unless the user explicitly abandons live-citation delivery.

### Hard Rules

- `Markdown -> pandoc -> docx` is a draft-generation path, not a complete final-delivery path.
- After a manuscript enters Word-only refinement mode, do not trigger `Markdown -> pandoc -> docx` for localized polish unless the user explicitly asks for a full rebuild.
- For Chinese academic DOCX, separate content correctness from Word-structure correctness.
- For citation-managed manuscripts, separate citation-field correctness from visual formatting correctness and validate both.
- Do not assume the reference template will automatically fix three-line tables, figure sizing, inline symbols, or paragraph indentation.
- Always inspect table paragraphs for `first_line_indent`, `left_indent`, and `right_indent`; academic tables should normally reset all of them to zero unless the user explicitly requests otherwise.
- For academic table cells, do not rely on paragraph styles to cancel indentation. The finalizer must write explicit OOXML indentation attributes: `w:left="0"`, `w:right="0"`, `w:firstLine="0"`, and `w:firstLineChars="0"`, and remove hanging indentation attributes.
- When inline formulas or symbol explanations are prone to corruption in Word, prefer native Word runs with explicit subscript formatting over trusting automatic math conversion.
- If a symbol appears in both a display equation and its explanatory paragraph, the explanatory paragraph must be repaired too; fixing the display equation alone is insufficient.
- Do not treat a successful formula conversion in XML as enough. The workflow must also guard against visual failures caused by orphan numbering paragraphs, collapsed symbols, and overlong one-line equations.
- Before overwriting a user-facing DOCX, write a temporary output first so an open file or a failed pass does not destroy the working draft.
- When the source path uses Word's built-in styles or a generic `python-docx Document()` base template, assume hidden blue theme residue exists until `document.xml`, `styles.xml`, and `theme1.xml` have passed audit.
- If post-processing removes or flattens citation fields, treat that as a delivery failure, not a minor defect.
- In Markdown-first mode, do not directly hand-edit the previous "healthy" final DOCX and treat that as the new delivery baseline; re-export from Markdown, then re-run finalization.
- In Word-only refinement mode, targeted direct DOCX edits are allowed only through temporary working copies, citation-field-safe post-processing, and audits. This is the preferred path for late-stage formula, table, figure, citation, paragraph, cover-page, abstract, and small wording repairs.
- Chinese and English final DOCX outputs must be finalized with the same structural rule set unless the user explicitly requests divergence.
- Academic submission paragraphs must be left-aligned by default. Do not deliver body text, headings, or references as centered or justified unless the user explicitly requests that style. The default exceptions are pure figure/image paragraphs, figure captions, table captions, native equation blocks, equation-number paragraphs, and academic table-cell paragraphs.

### Mandatory Delivery Audit

Before a final DOCX is allowed to overwrite the main deliverable, the post-processing workflow must confirm all of the following:

- live citation fields still exist in `word/document.xml`
- `m:oMath` or `m:eqArr` objects exist when model equations are present
- equation numbers stay on the same visual line as their equations and are right-aligned
- no orphan equation-number-only paragraphs remain after finalization
- long equations use the single-equation-paragraph wrapping strategy instead of a page-overflowing one-line block
- equation layout tables are borderless, are not processed as three-line tables, and contain no fixed line-height residues
- explanation paragraphs no longer expose broken pseudo-formula forms such as `Y_it`, `CR_it`, or `z(...)` as raw degraded text
- inline explanatory variables such as `K_{it}`, `PR_{kt}`, `N_{ikt}`, `L_{jt}`, and similar patent/network notation are converted to real subscript runs rather than delivered as raw underscore text
- collapsed symbolic residues such as `Resilienceit`, `AIWit`, `Outcomeit`, `w1`, `w2`, `μi`, `λt`, and `εit` are converted to true subscript runs where required
- figures are inline rather than floating anchors
- figure captions are independent paragraphs and retain numbering
- three-line tables remain structurally intact
- body text, headings, and references are left-aligned unless explicitly overridden
- academic table-cell paragraphs have zero first-line, left, and right indentation and are centered by default unless a target journal profile explicitly requires another alignment
- academic table-cell paragraph XML explicitly contains zero indentation, including `firstLineChars=0`, so WPS/Word cannot reapply body-style first-line indentation
- academic tables are centered and fit the available page width by default, while equation-layout tables remain borderless and are not forced into three-line-table styling
- figure captions and table captions are independent centered paragraphs
- visible Chinese body text has no unnecessary spaces between Chinese characters, between numbers and Chinese measurement words, or before Chinese punctuation, while Zotero field metadata is not edited directly
- chapter-opening titles, abstract, and references use page-break-before where required
- temporary DOCX exports are cleaned up after delivery

For recovery or working-draft layout mode, run the same structural checks but report citation-field absence as a blocking issue for formal submission rather than as a layout failure.

### Default Figure Standards

- Prefer inline figures rather than floating anchors unless the user explicitly needs wrap-around.
- Side-by-side figures must be scaled to a safe page width so no image exceeds its cell width or page margins.
- Captions must remain immediately adjacent to the figure block.
- If a figure risks clipping in Word, reduce figure width or adjust the containing layout; do not force it to fill the page.
- For Chinese academic work, verified PNG is preferred over SVG when Word rendering is unstable.
- Figure paragraphs must not inherit ordinary body-text first-line indentation or fixed body line spacing.
- Figure captions must be separate paragraphs, not embedded in the interpretive paragraph above the figure.
- Figure captions must be centered in submission finalization unless the user explicitly requests a different journal style.
- Table captions must also be centered in submission finalization unless the user explicitly requests a different journal style.
- Preferred caption formats:
  - Chinese: `图 1 标题`
  - English: `Figure 1. Title`

### Caption and Chinese Punctuation Cleanup

Use this gate for Chinese or bilingual formal DOCX finalization.

Caption rules:

- Figure captions and table captions must be independent centered paragraphs.
- Do not confuse narrative paragraphs such as `表 1 报告了...` with actual table captions.
- Chinese captions should preserve the caption number separator, such as `图 2a 标题`, while removing unnecessary inner spacing such as `4 位` when it should read `4位`.
- English captions should use the project or journal caption pattern, such as `Figure 1. Title` and `Table 1. Title`, and remain centered unless a journal style says otherwise.

Chinese text cleanup rules:

- Remove unnecessary spaces before Chinese punctuation.
- Convert visible English punctuation next to Chinese text to Chinese punctuation where safe.
- Remove unnecessary spaces between Chinese characters and between numbers and Chinese measurement words, such as `1% 水平`, `4 位`, and `2010 至 2024 年`.
- Do not edit Zotero citation-field metadata, bibliography field instructions, formulas, drawings, or other field-code content to fix apparent spaces or punctuation inside OOXML.

Audit rules:

- XML audit must confirm all real figure/table captions are centered.
- XML audit must confirm narrative paragraphs beginning with `表 1 报告...` or similar are not accidentally centered.
- Visible-text audit should check common Chinese spacing problems, but ignore Zotero field metadata where English abstracts may contain spaces before punctuation.

### Default Academic Table Standards

- Default three-line table widths:
  - top rule: `1.5 pt`
  - header separator rule: `0.5 pt`
  - bottom rule: `1.5 pt`
- By default, do not keep vertical borders and do not keep ordinary row gridlines unless the user explicitly requests auxiliary lines.
- Table header cells and table body cells should be centered by default for polished submission finalization unless the target template explicitly requires field-specific alignment.
- If a table includes long narrative or definition columns, the journal profile may require those columns to be left-aligned, but this must be an explicit table-level decision rather than inherited body indentation.
- Prefer compact academic geometry:
  - fixed column widths
  - small cell margins
  - zero paragraph indentation
  - explicit OOXML zero indentation, including `firstLineChars=0`
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

The finalizer must also recognize common Pandoc or LaTeX-flavored formula text generated from Markdown, including:

- `PR_{kt} = ...`
- `AIPatent_{it} = ...`
- `AIW_{it} = z(...) - z(...)`
- `Resilience_{it} = ...`
- `Channel_{it} = ...`
- `Outcome_{it} = ...`

These should be converted to native Word math objects or repaired inline symbol runs rather than delivered as raw source-code-like formula strings.

If the generated DOCX shows degraded forms like `(_i)` or missing Greek letters, treat that as a required post-processing fix, not a cosmetic issue.

For journal-submission finalization:

- core model equations should be converted to native Word math objects
- numbered display equations should use native equation objects plus a right-aligned tab stop and same-line equation number as the primary layout strategy
- multiline equations should remain one native Word math block rather than fragmented line-by-line text boxes
- equation numbers must stay on the same visual row as the equation block; do not deliver equation numbers as standalone paragraphs below or above the formula
- the visible target is: centered formula body, same-paragraph right-aligned equation number
- do not downgrade formulas into visible tables, images, text boxes, or plain text to "fix" alignment
- when the tab-stop strategy is visually unstable for a long formula, first use formula-specific fitting such as modest math-run font reduction, safe tab-stop adjustment, or controlled native wrapping
- a borderless equation-layout container is no longer the general default; it is allowed only as a documented recovery fallback after explicit user approval or unavoidable journal-template constraints
- any approved equation-layout fallback container must be excluded from ordinary three-line table formatting so table rules do not cross, compress, or visually cover formulas
- any approved equation-layout fallback container must use zero indentation, no visible borders, vertically centered content, and enough before/after spacing to avoid cramped formulas
- fixed line-height residues such as `w:lineRule="exact"`, `w:line`, and fixed row-height settings must be removed from fallback equation containers because they can clip multiline formulas in Word, WPS, or LibreOffice rendering
- inline pseudo-formulas in explanation text should be repaired to true subscript or superscript runs
- Greek letters, subscripts, and squared terms must not be delivered as bare underscore strings

## Formula Finalization Pipeline

Use this pipeline whenever a formal manuscript contains display equations, numbered models, measurement formulas, inline symbolic explanations, Greek letters, overbars, standard-deviation operators, summation expressions, or other source-like math notation that may degrade during Markdown-to-Word export.

Pipeline rule:

1. Identify formula-bearing content in the source manuscript before final delivery.
2. Build a formula inventory for display equations and inline symbol explanations.
3. Classify each item as:
   - display equation
   - inline explanatory symbol
   - plain text that should remain non-math
4. Render display equations through the OMML equation compiler.
5. Render inline symbolic explanations through the inline symbol renderer.
6. Run formula delivery audit before the final DOCX may overwrite the main deliverable.

Failure rule:

- If the finalizer cannot confidently classify or repair a formula-bearing item, do not silently leave it as raw pseudo-formula text in a formal delivery file.
- Either repair it through the pipeline or block formal delivery with a clear audit failure.

## Formula Inventory

For formal manuscript delivery, the finalization workflow must treat formulas as an explicit audited inventory rather than as ad hoc text fragments.

The inventory should distinguish at least:

- numbered empirical models
- measurement formulas
- network or weighting formulas
- multiline equations
- inline variable explanations
- inline Greek-letter or subscript notation

Inventory expectations:

- each numbered display equation should map to one delivery equation block
- each inline explanatory symbol pattern should map to one renderer rule or an approved fallback
- the delivery audit should be able to explain why a formula-like string remained plain text, if that ever occurs by design

When no explicit formula inventory file exists in the project, the finalizer must still behave as though an implicit inventory exists and audit all formula-like content it can detect.

## Inline Symbol Renderer

This module is responsible for symbols that belong in ordinary explanation paragraphs rather than in display-equation blocks.

Typical targets include:

- `Return_{im}`
- `AIDisclosure_{it}`
- `AIPatent_{it}`
- `PR_{kt}`
- `K_{it}`
- `N_{ikt}`
- `w_1`
- `w_2`
- Greek-letter coefficient terms such as `α_1`, `β_k`, `μ_i`, `λ_t`, `ε_{it}`
- squared terms such as `AIW²`

Renderer rules:

- convert base-plus-subscript forms to true Word subscript runs
- convert coefficient and squared-term notation to readable native Word runs
- avoid leaving raw `_`, `{}`, `\\sigma`, `\\overline`, or similar source-code-like residue in explanatory prose
- preserve citation fields and non-math content while repairing symbolic runs

The inline symbol renderer is required because many formula failures do not occur in display equations; they occur in the explanatory prose around them.

## OMML Equation Compiler

This module is responsible for display-equation conversion.

Primary responsibilities:

- convert formal formulas into native Word OMML objects
- preserve multiline equations as one equation block
- keep equation numbering on the same row as the equation
- choose a stable layout strategy for numbering

Default layout strategy:

- Use native Word OMML plus center/right tab stops as the default for all numbered display equations.
- The visible equation body should be centered; the visible number should be right-aligned on the same paragraph.
- Short single-line equations should remain on one visual row.
- Long single-line equations should first be fitted through equation-only font-size reduction, safe tab-stop tuning, or native wrapping while preserving the same numbered paragraph.
- Do not use floating text boxes, drawings, images, or ordinary visible tables for equation numbering.
- Do not use a borderless equation-layout table as the default automated solution; reserve it for explicit recovery fallback only.
- For long models that are displayed across multiple native formula rows, keep the formula visually coherent and place the number on the final continuation row. If the final continuation row is short and begins with additive tail terms such as `+ μ_i + λ_t + ε_it`, remove the center tab stop from that numbered continuation row and keep only the right tab stop for the number.

Compiler rule:

- The manuscript controller should request formula-safe delivery.
- The compiler must preserve the centered-native-formula plus right-number visual contract unless the user explicitly approves a documented fallback.
- Any fallback must be deterministic, auditable, and clearly reported as a deviation from the default formula contract.

## Formula Delivery Audit

Formal DOCX delivery must include a formula-specific audit, not just a generic visual pass.

The audit must check all of the following:

- every detected display equation that should be native math has `m:oMath` or `m:eqArr`
- equation numbers are present, ordered, and aligned on the same visual row as their equations
- multiline formulas are not fragmented into multiple unrelated equation blocks
- explanatory paragraphs no longer expose raw source-like strings such as `Stability_{it}`, `CR_{it}`, `\\overline{...}`, `\\sigma(...)`, or similar residue
- inline symbolic variables are rendered with real subscript or superscript formatting where required
- equation containers do not inherit ordinary three-line table rules, fixed row heights, or clipping-prone exact line spacing
- the rendered pages containing the longest formulas are visually inspected when rendering is available

Fail-closed rule:

- If raw formula residue survives in a formal delivery DOCX, the file fails formula audit and must not overwrite the main deliverable.

### Equation Layout Finalization Gate

Use this gate whenever a formal manuscript contains numbered empirical models, measurement formulas, or multiline equations.

Required structure:

- Use a native Word equation object plus a right-aligned tab stop and same-line equation number by default.
- Preserve one native Word math block for multiline equations.
- Keep formula numbers formatted as `（1）` for Chinese drafts and `(1)` or journal-required style for English drafts when the user requests English punctuation.
- The visible formula body must be centered and the visible equation number must be right-aligned on the same formula paragraph.
- If the default tab-stop layout becomes visually unstable for a long or multiline equation, first use equation-only fitting or controlled native wrapping; use a borderless fallback equation container only with explicit user approval or documented template necessity.

Required cleanup:

- Remove standalone equation-number paragraphs.
- Remove all visible borders from any fallback equation containers and cells.
- Prevent fallback equation containers from receiving three-line table rules.
- Remove fixed line spacing and fixed row heights inside fallback equation containers.
- Add enough formula spacing to avoid cramped or clipped equations.

Required audit:

- XML audit must confirm equation blocks contain `m:oMath` or `m:eqArr`.
- XML audit must confirm no standalone equation-number paragraph remains.
- XML audit must identify numbered short continuation rows and confirm they do not retain a center tab stop that can trap the equation number away from the right margin.
- XML audit must confirm numbered equations are not delivered as ordinary visible tables unless a documented fallback exception exists.
- XML audit must confirm fallback equation containers have no visible `w:tblBorders` or `w:tcBorders`.
- XML audit must confirm fallback equation containers have no fixed `w:lineRule="exact"` or `w:line` residue.
- Render QA must inspect at least the pages containing the longest formulas; XML checks alone are not sufficient because clipping, centered-body alignment, and same-row numbering are visual.

### Recommended Companion Scripts

For non-trivial thesis DOCX delivery, `scripts/postprocess_thesis_docx.py` remains available for legacy thesis-style repairs.

For bilingual journal-submission finalization, use:

`scripts/finalize_submission_docx.py`

Before citation-aware export in Zotero-based projects, use:

`scripts/zotero_preflight_recover.py`

Recommended interface:

```bash
python3 finalize_submission_docx.py \
  --input-docx temp_export.docx \
  --output-docx final.docx \
  --lang cn \
  --mode journal_submission \
  --citation-policy strict
```

For targeted repair of already-native numbered equations whose numbers drift away from the right margin, use:

```bash
python3 "$HOME/.codex/skills/chinese-word-pro/scripts/fix_equation_number_tabs.py" \
  --input-docx "<input.docx>" \
  --output-docx "<output.docx>"
```

This helper preserves native OMML equations and rewrites equation-number tab stops. It is especially useful for long empirical models where the final continuation line is short and the equation number otherwise stops at a center tab.

Recovery-draft interface when citation-aware export is temporarily broken:

```bash
python3 finalize_submission_docx.py \
  --input-docx temp_export.docx \
  --output-docx recovery_layout.docx \
  --lang cn \
  --mode journal_submission \
  --citation-policy warn
```

Do not call the `warn` output a formal submission file.

This companion script is the preferred execution point for:

- font and paragraph normalization
- equation-object repair
- pseudo-formula explanation repair
- figure-paragraph and figure-caption normalization
- three-line table repair
- chapter pagination
- garbling checks
- citation-field protection checks

Script gate expectations for `finalize_submission_docx.py`:

- it must implement the Formula Finalization Pipeline
- it must implicitly or explicitly maintain a formula inventory during finalization
- it must run the inline symbol renderer before delivery
- it must run the OMML equation compiler for display equations before delivery
- it must run the formula delivery audit and fail closed when raw formula residue remains
- it must not describe a file as formally finalized if formula audit fails, even when the rest of the typography is acceptable
- it must implement the Figure Gate by rejecting floating anchor images and normalizing figure/caption paragraphs
- it must implement the Paragraph and Table Geometry Gate by forcing academic table-cell paragraphs to centered alignment and explicit zero indentation at both Word API and OOXML levels
- it must fail closed when ordinary academic table cells still inherit body first-line indentation, including character-based Chinese indentation such as `firstLineChars`

Recommended Zotero preflight interface:

```bash
python3 "$HOME/.codex/skills/chinese-word-pro/scripts/zotero_preflight_recover.py" \
  --collection-key "<COLLECTION_KEY>" \
  --timeout 90 \
  --strict
```

Optional smoke-test audit:

```bash
python3 "$HOME/.codex/skills/chinese-word-pro/scripts/zotero_preflight_recover.py" \
  --collection-key "<COLLECTION_KEY>" \
  --smoke-docx "<SMOKE_TEST_DOCX>" \
  --timeout 90
```


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
