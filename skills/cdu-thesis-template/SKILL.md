---
name: cdu-thesis-template
description: Format and audit Chengdu University (CDU) undergraduate thesis DOCX files against the science/engineering thesis template. Use when working on 成都大学本科毕业设计（论文）, CDU thesis documents, 理科模板, DOCX formatting, cover/front matter, abstracts, table of contents, headings, body text, headers/footers, page numbers, figure/table captions, formulas, references, appendices, or acknowledgement formatting. This skill is format-only unless the user explicitly asks for writing or rewriting content.
---

# CDU Thesis Template

## Purpose

Use this skill to make DOCX formatting changes for Chengdu University undergraduate science/engineering thesis documents. Preserve the author's thesis content and make only the requested formatting, structure, style, and layout changes unless the user explicitly asks for content editing.

Keep the authorized template available at `assets/chengdu-university-science-thesis-template.docx`. Use it as the style and layout source of truth when a formatting rule is ambiguous. Public distributions of this skill may omit the template asset; ask the user to provide it locally when needed.

## Default Workflow

1. Copy the user's DOCX before editing. Never overwrite the only copy.
2. Run the audit script to identify obvious layout and style mismatches:

   ```powershell
   $skillDir = "path\to\cdu-thesis-template"
   python (Join-Path $skillDir "scripts\audit_cduthesis_docx.py") `
     "path\to\thesis.docx" `
     --out "path\to\audit-report.md"
   ```

3. Read `references/format-rules.md` before applying non-trivial formatting changes.
4. Apply minimal DOCX edits that address the user's request and the relevant audit findings.
5. Render the edited DOCX to PNG pages with the Documents skill renderer and inspect the pages before delivery.
6. Report what formatting was changed and any remaining items that need Word field updates, such as TOC refresh or page-number recalculation.

## Formatting Priorities

- Preserve thesis wording, technical content, citations, and equations.
- Prefer document styles and OOXML-level formatting over manual visual approximations.
- Keep front matter,正文,参考文献,附录,致谢 in the order required by the template.
- Treat template instruction text boxes and placeholder notes as removable only when the user asks to clean a final thesis.
- Use Chinese fonts for Chinese text and Times New Roman for Latin letters, numbers, formulas, page numbers, and English abstract content.
- Keep headings and captions compatible with automatic TOC/navigation behavior.

## Resources

- `references/format-rules.md`: condensed CDU science thesis formatting rules extracted from the template.
- `assets/chengdu-university-science-thesis-template.docx`: user-provided authorized Word template asset.
- `scripts/audit_cduthesis_docx.py`: read-only OOXML audit for sections, margins, styles, headers, captions, and references.

## Editing Guidance

- For layout-sensitive DOCX work, also use the Documents skill. Its render-and-inspect workflow is required before claiming a final DOCX is ready.
- For style transfer, open the template asset and copy the relevant style definitions or direct OOXML settings into the working document. Do not rely only on visual similarity.
- If the target DOCX has field-based TOC or page numbers, update fields in Word after editing when automation cannot safely refresh them.
- If the user's request conflicts with the CDU template, follow the user's explicit request and mention the conflict in the final response.
