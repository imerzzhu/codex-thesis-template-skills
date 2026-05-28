# CDU Science Thesis Formatting Rules

These rules are extracted from `assets/chengdu-university-science-thesis-template.docx`. Use the template asset as the final authority if a document has unusual direct formatting.

## Page Setup

- Paper: A4, 11906 x 16838 twips.
- Main body page margins: top 2268 twips (4 cm), bottom 1418 twips (2.5 cm), left 1418 twips (2.5 cm), right 1418 twips (2.5 cm).
- Header and footer distance: 851 twips (1.5 cm).
- The cover/front section in the provided template uses a smaller top margin; keep existing cover geometry unless the user asks to normalize it.
- Printing convention: cover, originality statement, authorization statement, Chinese/English abstracts are single-sided; TOC and正文 are double-sided.

## Header, Footer, and Page Numbers

- Cover has no visible header.
- Running header text: `成都大学本科毕业设计（论文）`.
- Header font: 宋体, 五号, centered, 1.25 line spacing, no paragraph spacing.
- Abstracts and TOC use roman page numbers such as I, II, III, IV.
- 正文 starts from page 1 with Arabic page numbers centered in the footer.
- Page-number font: Times New Roman, 小五.

## Front Matter

- Cover title:
  - Chinese title: no more than 25 Chinese characters.
  - English title: no more than 15 content words.
  - Chinese title font: 黑体, 三号, centered.
  - Latin letters and numbers: Times New Roman, 三号.
- Originality statement and authorization statement:
  - Title: 黑体, 小二, centered, 1.5 line spacing, 12 pt after.
  - Body: 宋体, 小三, 1.25 line spacing, no before/after spacing.
- Chinese abstract page:
  - Thesis title: 黑体, 小二, centered, 1.5 line spacing, 12 pt after.
  - Major/student metadata: 楷体, 小四, 1.25 line spacing, 12 pt after.
  - `摘要` label: 黑体, 小四, left aligned, no first-line indent.
  - Abstract body: 宋体, 小四, first-line indent 2 Chinese characters, 1.25 line spacing, no before/after spacing.
  - Abstract length target: about 300-600 Chinese characters, within one page.
  - `关键词：`: label in 黑体 小四; 3-5 keywords separated with Chinese semicolons.
- English abstract page:
  - Title: Times New Roman, bold, 小二, centered.
  - Metadata and body: Times New Roman, 小四.
  - `Abstract:` and `Key words:` labels are bold.
  - Keywords: 3-5 entries separated by English semicolons.

## TOC

- TOC title: `目 录`, 宋体, 小二, bold, centered, 1.5 line spacing, 12 pt after.
- TOC entries: same base font and line spacing as正文; show up to level 3 headings.
- TOC should be generated from real heading styles, not manually typed entries, unless the user explicitly asks for a static TOC.

## Body Text and Headings

- Body text: 宋体 for Chinese, Times New Roman for Latin/numbers, 小四, 1.25 line spacing, no before/after spacing, first-line indent 2 Chinese characters.
- Heading 1:
  - Style id in template: `1`, Word name `heading 1`.
  - Format: 黑体, 小三, centered, 1.5 line spacing, 0 pt before, 12 pt after.
  - Outline level: 0.
  - Each chapter starts on a new page.
  - Number format examples: `1 绪 论`, `2 正文格式说明`, `结 论`, `参考文献`, `附录一 附录名称`, `致 谢`.
- Heading 2:
  - Style id in template: `21`, Word name `heading 2`.
  - Format: 黑体, 四号, left aligned, no first-line indent, 1.5 line spacing, no before/after spacing.
  - Outline level: 1.
  - Number format example: `2.1 论文格式基本要求`.
- Heading 3:
  - Style id in template: `31`, Word name `heading 3`.
  - Format: 黑体, 小四, left aligned, no first-line indent, 1.5 line spacing, no before/after spacing.
  - Outline level: 2.
  - Number format example: `3.1.1 图的格式示例`.
- Fourth-level headings use `(1)`, `(2)`, `(3)`; 楷体, 小四, left aligned, 1.5 line spacing.
- Fifth-level headings use circled numbers such as `①`, `②`, `③`; 宋体, bold, 小四, left aligned, 1.5 line spacing.

## Figures, Tables, and Formulas

- Figures, tables, and formulas are numbered by chapter.
- Figure references: examples `图1.2`, `图3.1`.
- Table references: examples `表2.3`, `表3.1`.
- Formula references: examples `式（3.1）`, `见式（3.1）`.
- Figure placement:
  - Place figures after first mention in正文.
  - Keep figure and caption on the same page.
  - Use inline or top-and-bottom wrapping, not floating over text.
  - Leave one blank line above a figure when needed.
- Figure caption:
  - Below the figure.
  - Centered, 黑体, 五号, single line spacing, 0 before, about 1 line after.
  - Template style id: `affff3`, name `图名`.
- Table placement:
  - Center tables.
  - Keep table and caption together.
  - Leave one blank line above the table caption and below the table when needed.
  - Prefer three-line table style: no left/right borders, top/bottom thick lines, middle thin line.
- Table caption:
  - Above the table.
  - Centered, 黑体, 五号, single line spacing.
  - The template also uses the `图名` style for table captions.
- Table body:
  - 宋体, Times New Roman for Latin/numbers, 五号, centered, single line spacing.
  - Template table text style id: `affff9`, name `表`.
- Formulas:
  - Center formulas.
  - Place formula number at the right margin, such as `（3.1）`.
  - Use half-line spacing around formulas when needed.

## References

- Heading `参考文献` is required and uses Heading 1 formatting.
- Reference entries follow citation order in正文.
- Minimum count: at least 15 references; at least 5 journal articles; include a reasonable number of foreign-language journals.
- Citation markers in正文 use bracketed Arabic numbers as superscript, for example `[1]`, `[3,4]`, `[6-10]`.
- Reference body format: 宋体, 五号, left aligned, 1.25 line spacing, no before/after spacing.
- Template reference body style id: `affff7`, name `参考文献 正文`.
- Bibliography standard in the template: GB/T 7714-2005.

## Final Sections

- Appendix heading format follows Heading 1, for example `附录一 附录名称`.
- Appendix body follows正文 formatting.
- `致 谢` is required.
- Acknowledgement body follows正文 formatting and should stay relevant to thesis work.
