---
name: cdut-english-thesis-template
description: Format and audit Chengdu University of Technology (CDUT, 成都理工大学) English-major undergraduate thesis DOCX files against the supplied thesis and MLA citation templates. Use when working on 成都理工英语专业学士学位论文, CDUT English thesis formatting, DOCX layout cleanup, cover/front matter, English and Chinese abstracts, Contents, headings, headers/footers, page numbers, tables/figures, Notes, Works Cited, MLA seventh edition references, or parenthetical citations/夹注. This skill is format-focused unless the user explicitly asks for writing or rewriting content.
---

# CDUT English Thesis Template

## Purpose

Use this skill to format and audit Chengdu University of Technology English-major undergraduate thesis documents. Preserve the author's wording, argument, data, and citations unless the user explicitly asks for content editing.

Use these authorized template assets as the source of truth when a rule is ambiguous:

- `assets/thesis-format-with-references-and-citations.docx`
- `assets/references-and-parenthetical-citation.docx`

Public distributions of this skill may omit these template assets. Ask the user to provide authorized local copies when exact style transfer is needed. The original Word 97-2003 `.doc` files can also be placed in `assets/` for traceability.

## Default Workflow

1. Copy the user's document before editing. Never overwrite the only copy.
2. If the user provides `.doc`, convert a working copy to `.docx` first; keep the source `.doc` unchanged.
3. Run the audit script:

   ```powershell
   $skillDir = "path\to\cdut-english-thesis-template"
   python (Join-Path $skillDir "scripts\audit_cdut_english_thesis_docx.py") `
     "path\to\thesis.docx" `
     --out "path\to\audit-report.md"
   ```

4. Read `references/format-rules.md` before structural or layout edits.
5. Read `references/mla-citation-rules.md` before editing Works Cited or parenthetical citations.
6. Apply minimal DOCX edits that match the user's request and relevant audit findings.
7. Render the edited DOCX to PNG pages with the Documents skill renderer and inspect the pages before delivery.
8. Report changed formatting and any remaining manual Word tasks, such as updating TOC fields or page-number fields.

## Formatting Priorities

- Keep the thesis in English, with required Chinese front-matter labels and Chinese abstract text.
- Use Times New Roman for English text and page numbers; use Songti/SimSun for Chinese body text and Heiti/SimHei-style fonts for Chinese headings.
- Keep A4 paper, thesis front matter, abstract pages, Contents, body, Acknowledgements, Notes, and Works Cited in the template order.
- Preserve required section titles exactly where the template requires them: `Abstract`, `摘  要`, `Contents`, `Introduction`, `Conclusion`, `Acknowledgements`, `Notes`, and `Works Cited`.
- Keep `Works Cited` plural, unnumbered, left aligned, and on a new page.
- Treat blue or highlighted template instruction text as formatting guidance, not thesis content.
- Prefer document styles, section properties, and OOXML-level changes over manual visual approximations.

## Scope Boundaries

- Do not invent thesis claims, experiment data, source details, or defense content.
- Do not silently translate, polish, summarize, or rewrite正文; ask or follow the user's explicit instruction for content editing.
- Do not fabricate references. For missing bibliographic fields, mark them for the user instead of guessing.
- If the user's explicit formatting request conflicts with the CDUT template, follow the user and mention the conflict.

## Resources

- `references/format-rules.md`: condensed thesis layout and typography rules extracted from the main template.
- `references/mla-citation-rules.md`: MLA seventh edition Works Cited and parenthetical citation rules extracted from the citation template.
- `scripts/audit_cdut_english_thesis_docx.py`: read-only OOXML audit for `.docx` files.
- `assets/thesis-format-with-references-and-citations.docx`: user-provided authorized main thesis format template.
- `assets/references-and-parenthetical-citation.docx`: user-provided authorized reference/citation format template.
