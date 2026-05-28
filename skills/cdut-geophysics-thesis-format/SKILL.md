---
name: cdut-geophysics-thesis-format
description: Format, audit, and repair Chengdu University of Technology (CDUT, 成都理工大学) Geophysics College undergraduate thesis or graduation design DOCX files against the supplied Geophysics College DOTX template as the primary source of truth. Use when working on 成都理工大学地球物理学院本科学士学位论文, 本科毕业设计, DOCX thesis formatting, DOTX-template-first repair, cover/front matter, abstracts, 目录, Word actual page TOC checks, static TOC page-number repair, body headings, captions, formulas, headers, page numbers, references, or format-only cleanup. Preserve thesis content unless explicitly asked otherwise.
---

# CDUT Geophysics Thesis Format

## Purpose

Use this skill to audit and repair CDUT Geophysics College undergraduate thesis/design `.docx` formatting. Preserve the author's wording, data, argument, figures, tables, and references unless the user explicitly asks for content editing.

Use these authorized assets as traceable source materials:

- `assets/地球物理学院本科学士学位论文（设计）学术标准及基本规范_附件一（论文格式模板）-2018.dotx`
- `assets/论文（设计）编写格式.docx`

Public distributions of this skill may omit these template assets. Ask the user to provide authorized local copies before exact template-first repair.

Conflict rule: follow the DOTX template for all visible layout and typography, including cover, front matter, page setup, styles, headings, captions, headers, footers, and references. Use the format guide only to fill details that the DOTX template does not specify. Report material conflicts instead of hiding them.

## Default Workflow

1. Copy the user's thesis before editing. Never overwrite the only copy.
2. Read `references/format-rules.md` before structural or layout edits. Read `references/template-derived-rules.md` when changing scripts or checking exact template-derived values.
3. Run the audit script first:

   ```powershell
   $skillDir = "path\to\cdut-geophysics-thesis-format"
   $templateDotx = Join-Path $skillDir "assets\地球物理学院本科学士学位论文（设计）学术标准及基本规范_附件一（论文格式模板）-2018.dotx"
   python (Join-Path $skillDir "scripts\audit_cdut_geophysics_thesis_docx.py") `
     "path\to\thesis.docx" `
     --template-dotx $templateDotx `
     --out "path\to\audit-report.md" `
     --json-out "path\to\audit-report.json"
   ```

4. Run the repair script only on a copy or let it create its default repaired copy:

   ```powershell
   $skillDir = "path\to\cdut-geophysics-thesis-format"
   $templateDotx = Join-Path $skillDir "assets\地球物理学院本科学士学位论文（设计）学术标准及基本规范_附件一（论文格式模板）-2018.dotx"
   python (Join-Path $skillDir "scripts\fix_cdut_geophysics_thesis_docx.py") `
     "path\to\thesis.docx" `
     --template-dotx $templateDotx
   ```

5. Use optional repair flags only when the user asked for that behavior:
   - `--finalize-red-to-black`: convert red template/instruction text to black for final submission.
   - `--fix-caption-text`: normalize compact caption labels such as `图 21` to `图 2-1`.
   - `--no-repair-front-matter`: skip deletion/repair of obvious template sample front matter.
   - `--no-strict-body-songti`: skip strict regular-body Songti small-four enforcement.
   - `--header-text "地球物理学院 2026届本科毕业生学士学位论文"`: override the inferred template running header text.
6. For any table-of-contents page-number task, use Microsoft Word actual page numbers as the authority:

   ```powershell
   $skillDir = "path\to\cdut-geophysics-thesis-format"
   python (Join-Path $skillDir "scripts\audit_word_toc_pages.py") `
     "path\to\thesis.docx" `
     --out "path\to\toc-audit.md" `
     --json-out "path\to\toc-audit.json"
   ```

   If the TOC is static and only page numbers should change, repair only the final numeric run in each TOC row:

   ```powershell
   $skillDir = "path\to\cdut-geophysics-thesis-format"
   python (Join-Path $skillDir "scripts\fix_static_toc_word_pages.py") `
     "path\to\thesis.docx" `
     --out "path\to\thesis_Word实际页码修复.docx" `
     --audit-out "path\to\toc-repair-audit.md" `
     --json-out "path\to\toc-repair-audit.json"
   ```

   If a stale TOC field instruction remains while the visible TOC rows are static, repair the visible static rows and report that a later Word field update may overwrite them.

7. Render the repaired `.docx` with the Documents skill renderer and inspect PNG pages before delivery. Rendering is visual QA only; do not use LibreOffice/PNG page order to decide TOC page numbers unless the user explicitly requests LibreOffice/PDF pagination.
8. Report changed formatting plus remaining manual Word tasks, especially TOC field updates, unresolved placeholders, bibliography details, or caption/figure placement that needs human review.

## Scope Boundaries

- Do not rewrite, polish, summarize, translate, or invent thesis content.
- Do not fabricate references, missing metadata, dates, figure titles, or table titles.
- Do not renumber captions by default. Flag numbering gaps or compact labels in the audit report.
- Do not silently delete thesis content. It is acceptable to delete obvious template sample front matter when it contains sample markers such as `地物院 教授`, `2018年5月`, or `地球物理学院本科学士学位论文格式规范`.
- Do not overwrite the input file. The repair script refuses to write to the same path and requires `--force` before replacing an existing output file.

## Formatting Priorities

- Use the DOTX template first for all sections, not only the cover.
- Keep the document on A4 paper and apply template-derived section margins.
- Keep regular body text in 宋体/SimSun 小四 (12 pt), 1.5 line spacing, unless the DOTX explicitly marks the paragraph as cover/front-matter special text.
- Apply chapter headings as 小二黑体居中; section headings as 小三黑体; subsection headings as 小四黑体.
- Keep figure captions below figures and table captions above tables, numbered by chapter, 五号黑体居中.
- Keep formula numbers in parentheses and right aligned.
- Use the DOTX running header pattern from the body section onward, normally `地球物理学院 ****届本科毕业生学士学位论文`.
- Keep page numbers bottom-centered in 小五.
- For TOC page numbers, Microsoft Word `AdjustedPageNumber` is authoritative for the visible footer page number. Static TOC repair must change only the final numeric `w:t` run in each TOC row and must preserve TOC text, tabs, leaders, indentation, and fonts.
- Format references according to the DOTX and its GB/T 7714-2005 guidance; audit uncertain or inconsistent entries instead of guessing.
- Start terminal standalone sections such as 致谢 and 参考文献 on new pages. Do not let these headings appear at the bottom of the preceding section.
- If preserving a long English abstract creates a single orphaned keyword line, make the smallest front-matter spacing adjustment needed to keep the abstract block readable; do not shorten or rewrite the abstract.

## Resources

- `references/format-rules.md`: concise template-first operating rules.
- `references/template-derived-rules.md`: extracted DOTX page/style/front-matter values used by scripts.
- `scripts/audit_cdut_geophysics_thesis_docx.py`: read-only OOXML audit for `.docx` files; outputs Markdown and optional JSON.
- `scripts/fix_cdut_geophysics_thesis_docx.py`: deterministic OOXML formatter that writes a repaired copy.
- `scripts/audit_word_toc_pages.py`: read-only Microsoft Word actual-page audit for static or field-based TOCs.
- `scripts/fix_static_toc_word_pages.py`: deterministic static TOC page-number repair that writes a new copy and verifies against Word actual pages.
- `assets/`: user-provided authorized Word format guide and geophysics DOTX template for traceability.
