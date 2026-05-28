# CDU Defense PPT Template Style Guide

Extracted from `assets/cdu-defense-ppt-template.pptx`. Use the user-provided authorized PPTX asset as the final source of truth for exact positioning, crops, and master/layout behavior.

## Package Facts

- Template slides: 8.
- Slide size: 4:3 screen, `9144000 x 6858000` EMU.
- Slide masters: 1.
- Slide layouts: 3.
- Media assets: 9 total, including JPEG/JPG campus photos and PNG emblem/graphic assets.
- Theme colors are mostly Office defaults; the CDU brand red is used as explicit shape color.

## Visual Identity

- Primary color: CDU red `#9F2925`.
- Secondary colors: white background, black headline/body text, dark gray secondary body text, occasional bright red accents such as `#C00000`, `#CC0000`, and `#FF3300`.
- Main font: 黑体. The template also contains Arial and Microsoft YaHei UI in small amounts.
- Identity asset: 成都大学校徽 from the supplied template. Reuse it; do not redraw, trace, or approximate it.
- Overall style: formal university defense deck, red-white institutional visual system, clean content space, moderate decoration.

## Slide Families

1. **Cover**
   - Deep CDU red top and bottom wave fields.
   - Centered school emblem.
   - Main title in CDU red.
   - Small white footer text.
2. **Contents**
   - Header with small emblem at top left and subtle red gradient band.
   - Centered `目录 / Contents` title.
   - Numbered list blocks using CDU red squares and dotted red leader lines.
   - Bottom red footer strip with small school text at right.
3. **Plain Content**
   - Header/footer treatment retained.
   - Title in CDU red near top left.
   - Bulleted body in black/dark gray.
   - Body text target from template: 18-20 pt for normal content pages.
4. **Expanded Content / Right Rail**
   - Main content on the left.
   - Optional lower/right expansion rail for smaller notes, about 10-12 pt.
   - Keywords may use CDU red for emphasis.
5. **Image Content**
   - Text column paired with campus images.
   - Keep crops rectangular and aligned; avoid decorative over-cropping.
   - Replace images through PowerPoint image replacement rather than rebuilding loose image stacks when possible.
6. **Icon Paragraph**
   - Uses small square/icon markers before paragraph blocks.
   - Icons should be simple, dark, and functional; do not add unrelated decorative icons.
7. **Structure Diagram**
   - Uses red-bordered or red-headed boxes, grouped lanes, and compact labels.
   - Keep connectors and groupings semantically meaningful.
   - Use straight edges and disciplined alignment; avoid generic card grids.
8. **Closing**
   - Matches cover red wave fields.
   - Centered school emblem.
   - `谢谢观看` in CDU red.

## Template-Following Guidance

- Keep all slides 4:3 unless the user explicitly asks to convert aspect ratio.
- Use the template PPTX as the starting deck for new defense slides when possible.
- For adapting an existing deck, transfer content into matching template slide families rather than changing the template to match the old deck.
- Preserve the red header/footer grammar on content slides.
- Keep page numbers or small school footer text unobtrusive and aligned to the template footer region.
- Use one main idea per slide and keep text shorter than a thesis document page.
- For final defense decks, replace placeholder text such as `单击此处编辑标题样式` and `点击此处添加内容`.
- Do not remove the emblem from template pages unless a target layout has a legitimate reason and the user asks for it.

## Rendering QA

Render PPTX to PNG pages before delivery when possible:

```powershell
$renderDir = "path\to\render-output"
New-Item -ItemType Directory -Force -Path $renderDir | Out-Null
& soffice --headless --convert-to pdf --outdir $renderDir "path\to\deck.pptx"
& pdftoppm -png -r 120 (Join-Path $renderDir "deck.pdf") (Join-Path $renderDir "slide")
```

Inspect the contact sheet for:

- consistent red header/footer placement;
- no stretched logo or campus images;
- no text pasted into placeholder boxes with broken wrapping;
- no slide that visually drifts into a different template;
- readable body type and clear hierarchy at thumbnail size.
