# CDUT Geophysics Undergraduate Thesis Format Rules

These rules summarize how to use the supplied CDUT Geophysics College DOTX template and the `论文（设计）编写格式.docx` guide. If this summary and the source assets disagree, follow the DOTX template first and record the conflict in the audit.

## Source Priority

- Use the geophysics DOTX template as the primary source of truth for all visible formatting: page setup, cover, title/front matter, declaration, abstracts, contents, body headings, body text, captions, formulas, headers, footers, and references.
- Use the format guide only where the DOTX template is silent or ambiguous.
- If the document is a graduation design, replace visible `论文` wording with `设计` only when the user confirms the document type or the source document clearly identifies itself as design.

## Page Setup

- Paper: A4 portrait.
- Template-derived section margins are preferred over the guide's generic values.
- Header distance: about 1.5 cm in the DOTX.
- Footer distance: about 1.75 cm in the DOTX.
- Regular body text: 宋体/SimSun, 小四 (12 pt), 1.5 line spacing.
- Page number: 小五 (9 pt), bottom center.

## Required Order

Use this order unless the supplied thesis already has an approved school-specific variation:

1. Cover/title front matter using actual thesis metadata.
2. 学士学位论文（设计）诚信承诺书.
3. 摘要.
4. Abstract.
5. 目录.
6. Body chapters.
7. 结论.
8. 致谢.
9. 参考文献.
10. Optional appendices or software source listings when required.

Delete only obvious template sample front matter, such as pages containing `地球物理学院本科学士学位论文格式规范`, `地物院 教授`, or `2018年5月`, after preserving the real thesis metadata.

Standalone terminal sections, especially `致谢` and `参考文献`, should start on new pages in template-first repair. Flag or fix cases where those headings are left at the foot of the previous section.

Preserve abstract text exactly. When an English keywords line is orphaned onto a mostly blank page, use minimal front-matter spacing repair before considering any manual Word adjustment; never shorten keywords automatically.

## Headings

- Chinese abstract heading `摘要`: 小二黑体, centered.
- English abstract heading `Abstract`: centered; keep body text in 12 pt.
- Contents heading `目录`: 小二黑体, centered.
- Chapter headings such as `绪论`, `第一章 ...`, `结论`, `致谢`, and `参考文献`: 小二黑体, centered, Word outline level 1.
- Section headings such as `1.1 选题依据与研究意义`: 小三黑体, Word outline level 2.
- Subsection headings such as `1.1.1 傅里叶变换的公式`: 小四黑体, Word outline level 3.
- Keep heading styles with outline levels so Word can generate/update the table of contents.

## Contents And Page Numbers

- Always determine whether the contents area is a live Word TOC field or a static TOC before changing page numbers.
- For page-number correctness, Microsoft Word `AdjustedPageNumber` is the authority because it matches the visible footer page number. Do not infer TOC page numbers from physical page count, ZIP order, LibreOffice-rendered page sequence, PDF page sequence, or PNG filenames.
- LibreOffice/PNG rendering remains required for visual QA, but it is not a valid source for final TOC page numbers when Word and LibreOffice paginate differently.
- If the TOC is static and the user asks to fix only page numbers, replace only the final numeric `w:t` run on each TOC row. Preserve the label text, tabs, dot leaders, indentation, style, font, bold state, and all body headings.
- If the TOC is a Word field without visible static rows, prefer reporting that Word should update the field. Do not use Word automation to update fields and save the whole document unless the user explicitly requests that higher-risk operation.
- If a stale TOC field instruction remains but the visible contents are static rows with standalone page-number runs, repair those visible rows as static TOC content and report that a later Word field update may overwrite the static page numbers.
- When matching TOC entries to headings, normalize whitespace and control characters. Also allow common thesis-template differences such as `第一章 绪论` in the TOC matching a body heading `绪论`, and `ERA5Land` matching `ERA5-Land`.

## Tables And Figures

- Number tables and figures by chapter with Arabic numerals, e.g. `表 2-2`, `图 3-1`.
- Figure captions go below figures; table captions go above tables.
- Captions must start with the sequence number, use 五号黑体, and be centered.
- Do not fabricate missing figure/table titles or silently renumber captions. Flag gaps or compact labels for human review unless the user authorizes renumbering.

## Formulas

- Number formulas by chapter, e.g. `(5-1)` or `(2-1-1)`.
- Put the formula number in parentheses at the right end of the line.
- Do not insert dotted leaders between the formula and number.
- Explain physical quantities at first appearance.

## Header And Footer

- Running header starts from the body text section.
- Template header pattern: `地球物理学院 ****届本科毕业生学士学位论文`, with the year filled when known.
- Header font: 宋体/SimSun, 小五, centered.
- Page numbers should be bottom-centered, 小五.
- Front matter may have different visible page-number behavior; do not force a single rule if the template sectioning clearly handles it differently.

## References

- Follow the DOTX note that references should use GB/T 7714-2005 guidance.
- Reference heading `参考文献`: 小二黑体, centered, included in the contents.
- Reference entries: 五号宋体, 1.5 line spacing, hanging indent where needed.
- Do not fabricate missing bibliographic fields. Mark uncertain references for human review.

## Finalization Checks

- Remove or replace cover/title placeholder text according to actual thesis data before final printing.
- Change final-submission red template text to black when the user authorizes finalization.
- Update TOC fields in Word after heading repairs, or generate a static TOC and clearly report that it is static.
- Run the Word actual-page TOC audit after any TOC page-number edit. The expected success condition is that every TOC row matches the Word `AdjustedPageNumber` of its target heading.
- Re-render pages after repairs and visually inspect cover, declaration, abstracts, TOC, body headings, captions, references, headers, and page numbers. Treat rendering as a visual check, not a replacement for Word actual-page TOC audit.
