---
name: cdu-defense-ppt-template
description: Adapt and audit Chengdu University (CDU) undergraduate thesis defense PowerPoint decks against the supplied CDU defense PPT template. Use when working on 成都大学本科毕业论文答辩 PPT, CDU defense slides, thesis defense presentation templates, PPTX template-following, slide layout cleanup, 4:3 red-white CDU visual style, cover/contents/content/image/icon/structure/closing slides, school emblem placement, typography, theme color, footer, or presentation format consistency. This skill follows the template visual system and does not generate full defense content from a thesis unless explicitly requested.
---

# CDU Defense PPT Template

## Purpose

Use this skill to adapt thesis defense PPTX files to the Chengdu University undergraduate defense template. Preserve the presenter's thesis content and use the supplied template as the visual system for layout, color, school identity, typography, and slide rhythm.

Keep the authorized template available at `assets/cdu-defense-ppt-template.pptx`. Use it as the source of truth when a layout or style rule is ambiguous. Public distributions of this skill may omit the template asset; ask the user to provide it locally when needed.

## Default Workflow

1. Copy the user's PPTX before editing. Never overwrite the only copy.
2. Run the audit script to check template compatibility:

   ```powershell
   $skillDir = "path\to\cdu-defense-ppt-template"
   python (Join-Path $skillDir "scripts\audit_cdu_defense_pptx.py") `
     "path\to\defense.pptx" `
     --out "path\to\audit-report.md"
   ```

3. Read `references/style-guide.md` before substantial layout work.
4. Apply template-following changes: slide size, master/layout grammar, red-white palette, school emblem treatment, header/footer bars, and content layout families.
5. Render the edited PPTX to PNG pages and inspect a contact sheet before delivery.
6. Report template changes made and any remaining manual PowerPoint tasks, such as replacing placeholder images or refreshing embedded media.

## Template Rules

- Preserve 4:3 slide size (`9144000 x 6858000` EMU).
- Use CDU red `#9F2925` as the primary brand color with white space and black/dark gray body text.
- Keep the school emblem as a verified asset from the template. Do not redraw or approximate it.
- Use 黑体 as the primary Chinese presentation font. Keep body copy legible at defense-room distance.
- Use the template's page families: cover, contents, plain content, expanded content with right rail, image/content, icon paragraph, structure diagram, and closing.
- Keep the top red gradient/header band and bottom red footer/page marker treatment on content slides.
- Avoid converting the deck into generic rounded-card layouts or a modern theme unrelated to the supplied template.

## Scope Boundaries

- Do not invent thesis findings, experiment data, or defense script content.
- Do not fabricate CDU logos, badges, icons, or campus images. Reuse template assets or user-provided verified assets.
- If the user asks to generate content from a thesis later, use this skill only for the visual template layer and verify all content against the source thesis.

## Resources

- `assets/cdu-defense-ppt-template.pptx`: user-provided authorized 8-slide CDU defense template.
- `references/style-guide.md`: condensed visual, layout, and usage rules extracted from the template.
- `scripts/audit_cdu_defense_pptx.py`: read-only OOXML audit for PPTX template consistency.
