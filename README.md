# chinese-word-pro

`chinese-word-pro` is a reusable Codex skill for creating, repairing, formatting, and validating Chinese Word `.docx` files and for finalizing bilingual academic submission Word deliverables.

It is built for academic and formal Chinese documents where typography, encoding, tables, figures, and formula rendering matter, especially when the final deliverable must be safe for submission rather than just visually acceptable at first glance.

## What This Skill Does

This skill is the Word-structure and typography specialist.

It helps the user:

- produce Chinese Word files from UTF-8 content sources
- enforce Chinese and Latin font separation correctly
- normalize table layout and three-line tables
- repair formula and subscript display problems
- reduce figure clipping and layout breakage
- inspect XML for garbling
- protect citation fields during DOCX post-processing
- finalize paired Chinese and English submission-ready Word files with the same structural rules

## Best Use Cases

Use this skill when the task involves:

- Chinese thesis or dissertation output
- Chinese journal article Word formatting
- table-heavy academic Word files
- mixed Chinese and English body text
- formula, Greek-letter, or subscript corruption in Word
- final manuscript delivery that must preserve citation fields

## Design Position

This skill is not the manuscript controller.

It should usually be used as the downstream Word-finishing layer after content has already been prepared elsewhere.

In a paper workflow:

- `management-empirical-writer` controls manuscript logic and delivery gates
- `chinese-word-pro` controls Word finalization, typography, and damage prevention

## Core Non-negotiables

- Chinese content must begin from UTF-8 source files.
- `run.font.name` alone is not sufficient; East Asian fonts must be set explicitly.
- Final DOCX delivery must be checked for garbling in OOXML.
- Temporary exports must not directly overwrite the user-facing final DOCX.
- If citation fields are part of the workflow, post-processing must not flatten or remove them.

## Default Formatting Intent

Unless the user specifies otherwise, this skill aims for:

- Chinese body: `宋体`
- English body: `Times New Roman`
- stable heading hierarchy
- left-aligned academic paragraphs by default
- academic table geometry
- three-line tables where appropriate
- safer inline formula display

It is especially suited to Chinese academic writing where Word output must look disciplined rather than flashy.

## Recommended Workflow

1. Edit the UTF-8 content source first.
2. Generate a temporary DOCX with the project-approved export chain.
3. If the document uses citation fields, verify they are present before post-processing.
4. Run DOCX post-processing for tables, figures, paragraphs, and formulas.
5. Re-check citation fields after post-processing.
6. Check `word/document.xml` for garbling markers.
7. Render to images when rendering is available.
8. Only then replace the final user-facing `.docx`.

## Formal Delivery vs Recovery Drafts

This skill distinguishes formal submission output from recovery or working-draft layout output.

Formal submission output:

- uses the citation-aware export chain first
- runs `finalize_submission_docx.py` with `--citation-policy strict`
- fails if live citation fields are absent in citation-managed manuscripts
- may replace user-facing final `.docx` files only after all citation, formula, table, figure, garbling, and render checks pass

Recovery or working-draft layout output:

- is allowed only when a readable Word file is needed before citation-aware export has been repaired
- runs `finalize_submission_docx.py` with `--citation-policy warn`
- may repair formulas, tables, captions, figures, pagination, and typography
- must be reported as a recovery draft or working draft
- must not be represented as a formal submission file when live citation fields are absent

Non-citation-managed output:

- may use `--citation-policy off`
- should not be used for empirical manuscripts with citekeys unless the user explicitly abandons live-citation delivery

## Academic Submission Finalization

This skill now includes a formal "Academic Submission Finalization" role for empirical paper delivery.

That role covers:

- bilingual `.docx` finalization
- journal-style table normalization
- figure-caption normalization
- native Word formula repair
- chapter pagination
- citation-safe post-processing

Typical finalization responsibilities:

- convert core equations into native Word math objects
- repair inline pseudo-formulas such as `Y_it`, `CR_it`, or `z(...)`
- keep numbered formulas on the same visual line as their equation numbers
- force figures to remain inline rather than floating
- separate figure captions from interpretation paragraphs
- preserve caption numbering
- center figure captions and table captions as independent paragraphs
- force body text, headings, captions, references, and table-cell text to left alignment unless explicitly overridden
- enforce chapter-open page breaks
- verify that citation fields survive post-processing

## Formal Delivery Flow

```mermaid
flowchart TD
    A["Prepare UTF-8 content source"] --> B["Generate temporary DOCX"]
    B --> C["Check citation fields if applicable"]
    C --> D["Apply DOCX post-processing"]
    D --> E["Re-check citation fields"]
    E --> F["Check XML for garbling markers"]
    F --> G["Render and visually inspect when available"]
    G --> H["Replace final DOCX only after all checks pass"]
```

## Mandatory Delivery Audit

A final submission DOCX should not be considered passed unless all of the following are true:

- live citation fields are still present
- no garbling markers appear in OOXML
- body text, headings, captions, references, and table-cell paragraphs are left-aligned unless the user explicitly requested otherwise
- figure paragraphs are inline and captions remain independent paragraphs
- figure captions and table captions are centered independently from body text
- core formulas remain native Word objects where required
- equation numbers are right-aligned on the same row as their equations
- equation layout tables are borderless and not treated as ordinary three-line tables
- tables preserve academic three-line structure
- abstract, major chapters, and references start on new pages where the workflow requires
- temporary exports are cleaned up after delivery

For recovery drafts, run the same structural audit but record missing citation fields as a formal-delivery blocker rather than silently accepting the file.

## Citation-Field Protection

This repository now explicitly treats citation-field protection as a delivery requirement.

For citation-managed manuscripts:

- post-processing must preserve Word citation fields
- a visually correct file is still a failed deliverable if citation fields disappeared
- the final check should inspect `word/document.xml`

Typical markers include:

- `ADDIN ZOTERO_ITEM`
- `CSL_CITATION`

## Zotero Preflight and Recovery

For Zotero-based manuscripts, run the Zotero preflight helper before citation-aware Word export:

```bash
python3 "$HOME/.codex/skills/chinese-word-pro/scripts/zotero_preflight_recover.py" \
  --collection-key "<ZOTERO_COLLECTION_KEY>" \
  --timeout 90 \
  --strict
```

The helper locates Zotero, opens it if needed, waits for the local connector, and tries to open the target collection through Zotero's `zotero://select` URI scheme.

If the scripted path cannot make the collection usable, the workflow may use Computer Use for one GUI recovery attempt to focus Zotero and select the intended collection. If the connector, collection, Better BibTeX, or MCP route still fails, stop and report the error. Do not produce a formal Word file with flattened citations.

When new references were added after the last healthy Word export, require a small live-citation smoke test before full delivery. The helper can audit that smoke DOCX:

```bash
python3 "$HOME/.codex/skills/chinese-word-pro/scripts/zotero_preflight_recover.py" \
  --collection-key "<ZOTERO_COLLECTION_KEY>" \
  --smoke-docx "<SMOKE_TEST_DOCX>" \
  --timeout 90
```

## Garbling and Formula Safety

This skill is designed to catch failures such as:

- replacement characters like `�`
- suspicious runs like `????`
- degraded inline math such as missing Greek letters
- broken subscript forms such as `(_i)` in place of true subscript

When needed, it prefers explicit Word runs and subscript formatting over trusting automatic conversion blindly.

For submission finalization, it also prefers native Word equation objects over plain-text pseudo-formulas whenever the manuscript presents formal empirical models.

The finalizer should recognize both plain pseudo-formulas and Pandoc/LaTeX-style formula strings, including `AIPatent_{it}`, `AIW_{it}`, `Resilience_{it}`, `Channel_{it}`, `Outcome_{it}`, `PR_{kt}`, and `z(...)`. These forms should not survive in Word as source-code-like text when they are display equations or symbol explanations.

## Equation Layout Finalization

Numbered display equations require a stricter Word layout than ordinary paragraphs.

Default layout:

- Use a borderless two-cell table for each numbered equation.
- Put the centered native Word equation object in the left cell.
- Put the equation number in the right cell and right-align it.
- Keep multiline formulas as one native math block, not separate formula paragraphs.
- Keep formula numbering on the same visual row as the equation block.

Required protection:

- Do not let equation layout tables receive ordinary three-line table borders.
- Remove all visible table and cell borders from equation layout tables.
- Remove fixed line spacing and fixed row heights inside equation layout tables.
- Add enough before/after spacing so multiline equations do not look cramped.
- Check the rendered pages containing the longest formulas; XML structure alone cannot prove that no clipping occurred.

Failure examples:

- equation number appears below the equation
- equation number appears on the left or in an inconsistent position
- table rules cross through formulas
- multiline formula is clipped because of fixed line height
- formula is correct in XML but visually cramped or cut off in PDF/Word rendering

Command examples:

```bash
# Formal delivery: fail if citation fields are absent.
python3 "$HOME/.codex/skills/chinese-word-pro/scripts/finalize_submission_docx.py" \
  --input-docx temp_export.docx \
  --output-docx final.docx \
  --lang cn \
  --mode journal_submission \
  --citation-policy strict

# Recovery or working draft: repair layout, but warn if citation fields are absent.
python3 "$HOME/.codex/skills/chinese-word-pro/scripts/finalize_submission_docx.py" \
  --input-docx temp_export.docx \
  --output-docx recovery_layout.docx \
  --lang cn \
  --mode journal_submission \
  --citation-policy warn
```

## Caption And Chinese Text Cleanup

Formal Chinese or bilingual Word delivery should also clean small typographic issues that are easy to miss.

Caption rules:

- Figure captions and table captions should be independent centered paragraphs.
- Narrative paragraphs such as `表 1 报告了...` are not captions and should not be centered.
- Chinese caption numbering should remain readable, such as `图 2a 标题`.
- English captions should follow the project pattern, such as `Figure 1. Title` and `Table 1. Title`.

Chinese visible-text cleanup:

- Remove unnecessary spaces before Chinese punctuation.
- Remove unnecessary spaces between Chinese characters.
- Remove unnecessary spaces between numbers and Chinese measurement words, such as `1%水平`, `4位`, and `2010至2024年`, unless the journal style requires otherwise.
- Convert visible English punctuation next to Chinese text to Chinese punctuation where safe.

Safety rule:

- Do not edit Zotero field metadata, field instructions, equations, drawings, or other hidden OOXML content merely because it contains English spaces or punctuation.
- Always re-audit citation fields after cleanup.

## Common Failure Modes This Skill Is Meant to Prevent

- Chinese text damaged before DOCX generation
- mixed Chinese and English fonts rendering inconsistently
- tables with wrong indentation or broken three-line rules
- clipped figures or unsafe side-by-side layouts
- centered captions drifting back to body alignment or body paragraphs being mistaken for captions
- formula text that looks acceptable in Markdown but degrades in Word
- formula numbers that drift into standalone paragraphs or formula tables that inherit three-line table borders
- post-processing that accidentally destroys citation fields

## Repository Contents

- `SKILL.md`: main operating rules
- `references/`: Chinese Word and thesis-format references
- `scripts/`: build and post-process helpers, including submission finalization
- `assets/`: templates and sample inputs

## Relationship to Final Manuscript Delivery

This skill improves Word quality, but it should not by itself decide whether a paper is ready for formal submission delivery.

That final decision belongs to the manuscript workflow controller. In the recommended setup:

- `management-empirical-writer` determines whether formal delivery may proceed
- `chinese-word-pro` makes sure the Word file is structurally safe enough to deliver

## Companion Scripts

- `scripts/build_chinese_word.py`: UTF-8 source to initial Chinese Word generation
- `scripts/postprocess_thesis_docx.py`: legacy thesis-oriented DOCX cleanup
- `scripts/finalize_submission_docx.py`: bilingual journal-submission finalization for formulas, figures, tables, pagination, and citation-safe DOCX delivery
- `scripts/zotero_preflight_recover.py`: Zotero launch, collection selection, connector preflight, and optional live-citation smoke DOCX audit
