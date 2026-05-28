---
name: swpu-thesis-template
description: Format and audit Southwest Petroleum University (SWPU, 西南石油大学) undergraduate thesis and graduation design DOCX files against the supplied Electrical Engineering and Information template. Use when working on 西南石油大学本科毕业设计（论文）, SWPU thesis formatting, 电信院 or 电气信息学院 thesis templates, DOCX layout cleanup, Chinese/English covers, abstracts, table of contents, chapter headings, odd/even headers, page numbers, table/figure/formula captions, references, appendices, or final thesis format consistency. This skill is format-only unless the user explicitly asks for writing or rewriting content.
---

# SWPU Thesis Template

## Purpose

Use this skill to format and audit Southwest Petroleum University undergraduate thesis or graduation-design DOCX files. Preserve the author's thesis content, data, references, formulas, and conclusions unless the user explicitly asks for content editing.

Use `assets/swpu-eei-undergraduate-thesis-template.docx` as the source of truth when a rule is ambiguous. The asset is the supplied Electrical Engineering and Information template. Public distributions of this skill may omit the template asset; ask the user to provide an authorized local copy when needed.

## Default Workflow

1. Copy the user's DOCX before editing. Never overwrite the only copy.
2. Run the audit script:

   ```powershell
   $skillDir = "path\to\swpu-thesis-template"
   python (Join-Path $skillDir "scripts\audit_swpu_thesis_docx.py") `
     "path\to\thesis.docx" `
     --out "path\to\audit-report.md"
   ```

3. Read `references/format-rules.md` before layout, font, margin, heading, page-number, table, figure, or formula edits.
4. Read `references/structure-rules.md` before front-matter, chapter order, references, appendix, header/footer, or TOC changes.
5. Apply minimal DOCX edits that match the user's request and the relevant audit findings.
6. Render the edited DOCX to PNG pages with the Documents skill renderer and inspect the pages before delivery.
7. Report changed formatting and any remaining manual Word tasks, such as updating TOC fields, cross-references, or page-number fields.

## Formatting Priorities

- Keep A4 portrait pages with the template's margins and 4-section front-matter/body separation unless the user provides a newer school requirement.
- Preserve the Chinese cover, English cover, Chinese abstract, English abstract, Contents,正文 chapters,谢辞,参考文献, and附录 order.
- Use Songti/SimSun-style Chinese text with Times New Roman for English text where the template requires it.
- Keep正文 small 4 Songti, first-line indent 2 Chinese characters, and exact 22 pt line spacing.
- Keep正文 odd/even running headers: odd pages use the thesis title; even pages use `西南石油大学本科毕业设计（论文）`.
- Keep table captions above tables, figure captions below figures, and formula numbers right aligned in parentheses.
- Prefer document styles, section properties, fields, and OOXML-level edits over visual approximations.

## Scope Boundaries

- Do not invent thesis claims, experiments, data, references, appendix code, or school requirements.
- Do not silently rewrite, polish, translate, summarize, or expand正文.
- Do not fabricate missing bibliography fields; mark them for user review instead.
- Treat visible red/placeholder/instruction text in the template as guidance to remove or replace in a final thesis, not as required thesis content.
- If the user's explicit request conflicts with the SWPU template, follow the user and mention the conflict.

## Resources

- `assets/swpu-eei-undergraduate-thesis-template.docx`: user-provided authorized SWPU Electrical Engineering and Information undergraduate thesis template.
- `references/format-rules.md`: page setup, style, typography, heading, table, figure, formula, and reference formatting rules.
- `references/structure-rules.md`: required document order, front matter, section/page-number behavior, and header/footer rules.
- `scripts/audit_swpu_thesis_docx.py`: read-only OOXML audit for `.docx` files.
