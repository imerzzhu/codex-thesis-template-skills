#!/usr/bin/env python3
"""Create a template-first repaired copy of a CDUT Geophysics thesis DOCX.

The script edits OOXML formatting only. It does not rewrite thesis prose,
renumber captions by default, fabricate references, or overwrite the input.
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


NS = {
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

XML_NS = "http://www.w3.org/XML/1998/namespace"

for prefix, uri in (("w", NS["w"]), ("r", NS["r"])):
    ET.register_namespace(prefix, uri)
ET.register_namespace("", NS["rel"])


def qn(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def w_attr(name: str) -> str:
    return qn("w", name)


def attr(el: ET.Element | None, name: str) -> str | None:
    if el is None:
        return None
    return el.get(w_attr(name))


def read_xml(zf: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(zf.read(name))


def to_xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def text_of(el: ET.Element) -> str:
    chunks: list[str] = []
    for node in el.iter():
        if node.tag == qn("w", "t") and node.text:
            chunks.append(node.text)
        elif node.tag == qn("w", "tab"):
            chunks.append("\t")
        elif node.tag in {qn("w", "br"), qn("w", "cr")}:
            chunks.append("\n")
    return "".join(chunks)


def norm_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def ensure_child(parent: ET.Element, tag: str, insert_at: int | None = None) -> ET.Element:
    child = parent.find(tag)
    if child is None:
        child = ET.Element(tag)
        if insert_at is None:
            parent.append(child)
        else:
            parent.insert(insert_at, child)
    return child


def ensure_ppr(p: ET.Element) -> ET.Element:
    ppr = p.find("w:pPr", {"w": NS["w"]})
    if ppr is None:
        ppr = ET.Element(qn("w", "pPr"))
        p.insert(0, ppr)
    return ppr


def set_p_style(p: ET.Element, style_id: str | None) -> None:
    if not style_id:
        return
    ppr = ensure_ppr(p)
    pstyle = ensure_child(ppr, qn("w", "pStyle"), 0)
    pstyle.set(w_attr("val"), style_id)


def set_jc(ppr: ET.Element, value: str) -> None:
    jc = ensure_child(ppr, qn("w", "jc"))
    jc.set(w_attr("val"), value)


def set_spacing(ppr: ET.Element, before: str = "0", after: str = "0", line: str = "360") -> None:
    spacing = ensure_child(ppr, qn("w", "spacing"))
    spacing.attrib.clear()
    spacing.set(w_attr("before"), before)
    spacing.set(w_attr("after"), after)
    spacing.set(w_attr("line"), line)
    spacing.set(w_attr("lineRule"), "auto")


def set_indent(
    ppr: ET.Element,
    *,
    first_line: str | None = None,
    hanging: str | None = None,
    left: str | None = None,
    clear: bool = False,
) -> None:
    ind = ensure_child(ppr, qn("w", "ind"))
    if clear:
        ind.attrib.clear()
    for key, value in (("firstLine", first_line), ("hanging", hanging), ("left", left)):
        qkey = w_attr(key)
        if value is None:
            ind.attrib.pop(qkey, None)
        else:
            ind.set(qkey, value)


def set_run_format(
    rpr: ET.Element,
    east_asia: str,
    ascii_font: str,
    size_half_points: str,
    *,
    bold: bool | None = None,
    color: str | None = None,
) -> None:
    rfonts = ensure_child(rpr, qn("w", "rFonts"), 0)
    rfonts.set(w_attr("ascii"), ascii_font)
    rfonts.set(w_attr("hAnsi"), ascii_font)
    rfonts.set(w_attr("cs"), ascii_font)
    rfonts.set(w_attr("eastAsia"), east_asia)
    for tag in ("sz", "szCs"):
        size = ensure_child(rpr, qn("w", tag))
        size.set(w_attr("val"), size_half_points)
    if bold is not None:
        for tag in (qn("w", "b"), qn("w", "bCs")):
            existing = rpr.find(tag)
            if bold and existing is None:
                rpr.append(ET.Element(tag))
            elif not bold and existing is not None:
                rpr.remove(existing)
    if color is not None:
        color_el = ensure_child(rpr, qn("w", "color"))
        color_el.set(w_attr("val"), color)


def format_paragraph_runs(
    p: ET.Element,
    east_asia: str,
    ascii_font: str,
    size_half_points: str,
    *,
    bold: bool | None = None,
    color: str | None = None,
) -> None:
    for r in p.findall(".//w:r", {"w": NS["w"]}):
        rpr = r.find("w:rPr", {"w": NS["w"]})
        if rpr is None:
            rpr = ET.Element(qn("w", "rPr"))
            r.insert(0, rpr)
        set_run_format(rpr, east_asia, ascii_font, size_half_points, bold=bold, color=color)


def clear_text_content(p: ET.Element) -> None:
    for child in list(p):
        if child.tag != qn("w", "pPr"):
            p.remove(child)


def append_text_run(
    p: ET.Element,
    text: str,
    *,
    east_asia: str,
    ascii_font: str,
    size_half_points: str,
    bold: bool,
    color: str = "000000",
) -> None:
    if not text:
        return
    r = ET.SubElement(p, qn("w", "r"))
    rpr = ET.SubElement(r, qn("w", "rPr"))
    set_run_format(rpr, east_asia, ascii_font, size_half_points, bold=bold, color=color)
    t = ET.SubElement(r, qn("w", "t"))
    if text[:1].isspace() or text[-1:].isspace():
        t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text


def replace_with_segments(p: ET.Element, segments: list[tuple[str, str, str, str, bool]]) -> None:
    clear_text_content(p)
    for text, east_asia, ascii_font, size, bold in segments:
        append_text_run(p, text, east_asia=east_asia, ascii_font=ascii_font, size_half_points=size, bold=bold)


def style_name(style: ET.Element) -> str:
    name = style.find("w:name", {"w": NS["w"]})
    return (attr(name, "val") or "").lower()


def find_style_id(styles: ET.Element, names: Iterable[str]) -> str | None:
    wanted = {name.lower() for name in names}
    for style in styles.findall("w:style", {"w": NS["w"]}):
        sid = attr(style, "styleId")
        if style_name(style) in wanted and sid:
            return sid
    return None


def ensure_style(styles: ET.Element, style_id: str, name: str, based_on: str | None = None) -> ET.Element:
    for style in styles.findall("w:style", {"w": NS["w"]}):
        if attr(style, "styleId") == style_id:
            return style
    style = ET.Element(
        qn("w", "style"),
        {w_attr("type"): "paragraph", w_attr("customStyle"): "1", w_attr("styleId"): style_id},
    )
    name_el = ET.SubElement(style, qn("w", "name"))
    name_el.set(w_attr("val"), name)
    if based_on:
        based = ET.SubElement(style, qn("w", "basedOn"))
        based.set(w_attr("val"), based_on)
    styles.append(style)
    return style


def update_style(
    style: ET.Element,
    east_asia: str,
    ascii_font: str,
    size_half_points: str,
    *,
    bold: bool | None,
    align: str | None = None,
    before: str = "0",
    after: str = "0",
    line: str = "360",
    outline_level: str | None = None,
    first_line: str | None = None,
    hanging: str | None = None,
) -> None:
    ppr = ensure_child(style, qn("w", "pPr"))
    set_spacing(ppr, before=before, after=after, line=line)
    if align:
        set_jc(ppr, align)
    set_indent(ppr, first_line=first_line, hanging=hanging, clear=True)
    if outline_level is not None:
        outline = ensure_child(ppr, qn("w", "outlineLvl"))
        outline.set(w_attr("val"), outline_level)
    rpr = ensure_child(style, qn("w", "rPr"))
    set_run_format(rpr, east_asia, ascii_font, size_half_points, bold=bold)


def normalize_styles(styles: ET.Element) -> dict[str, str]:
    normal_id = find_style_id(styles, ["normal"]) or "CDUTNormal"
    normal = ensure_style(styles, normal_id, "Normal")
    update_style(normal, "宋体", "Times New Roman", "24", bold=False, align="both", line="360")

    heading1_id = find_style_id(styles, ["heading 1"]) or "CDUTHeading1"
    heading2_id = find_style_id(styles, ["heading 2"]) or "CDUTHeading2"
    heading3_id = find_style_id(styles, ["heading 3"]) or "CDUTHeading3"
    caption_id = find_style_id(styles, ["caption", "题注"]) or "CDUTCaption"
    reference_id = find_style_id(styles, ["cdut reference"]) or "CDUTReference"

    h1 = ensure_style(styles, heading1_id, "heading 1", based_on=normal_id)
    h2 = ensure_style(styles, heading2_id, "heading 2", based_on=normal_id)
    h3 = ensure_style(styles, heading3_id, "heading 3", based_on=normal_id)
    cap = ensure_style(styles, caption_id, "caption", based_on=normal_id)
    ref = ensure_style(styles, reference_id, "CDUT Reference", based_on=normal_id)

    update_style(h1, "黑体", "Times New Roman", "36", bold=True, align="center", before="340", after="330", line="578", outline_level="0")
    update_style(h2, "黑体", "Times New Roman", "30", bold=True, align="left", before="260", after="260", line="416", outline_level="1")
    update_style(h3, "黑体", "Times New Roman", "24", bold=True, align="left", before="260", after="260", line="416", outline_level="2")
    update_style(cap, "黑体", "Times New Roman", "20", bold=False, align="center", before="120", after="120", line="300")
    update_style(ref, "宋体", "Times New Roman", "20", bold=False, before="0", after="0", line="360", hanging="240")

    return {
        "normal": normal_id,
        "heading1": heading1_id,
        "heading2": heading2_id,
        "heading3": heading3_id,
        "caption": caption_id,
        "reference": reference_id,
    }


@dataclass
class TemplateRules:
    sections: list[dict[str, str]]
    cover_pprs: list[ET.Element | None]


def default_template_path() -> Path:
    return Path(__file__).resolve().parents[1] / "assets" / "地球物理学院本科学士学位论文（设计）学术标准及基本规范_附件一（论文格式模板）-2018.dotx"


def read_template_rules(template_path: Path | None) -> TemplateRules:
    path = template_path or default_template_path()
    if not path.exists() or not zipfile.is_zipfile(path):
        return TemplateRules(sections=[], cover_pprs=[])
    with zipfile.ZipFile(path) as zf:
        document = read_xml(zf, "word/document.xml")
        sections: list[dict[str, str]] = []
        for sect in document.findall(".//w:sectPr", {"w": NS["w"]}):
            pg = sect.find("w:pgSz", {"w": NS["w"]})
            mar = sect.find("w:pgMar", {"w": NS["w"]})
            if pg is None or mar is None:
                continue
            sections.append(
                {
                    "w": attr(pg, "w") or "11906",
                    "h": attr(pg, "h") or "16838",
                    "top": attr(mar, "top") or "1701",
                    "right": attr(mar, "right") or "1800",
                    "bottom": attr(mar, "bottom") or "1440",
                    "left": attr(mar, "left") or "1800",
                    "header": attr(mar, "header") or "851",
                    "footer": attr(mar, "footer") or "992",
                    "gutter": attr(mar, "gutter") or "0",
                }
            )
        cover_pprs: list[ET.Element | None] = []
        for p in document.findall(".//w:p", {"w": NS["w"]}):
            text = norm_space(text_of(p))
            if not text:
                continue
            cover_pprs.append(deepcopy(p.find("w:pPr", {"w": NS["w"]})))
            if len(cover_pprs) >= 13:
                break
    return TemplateRules(sections=sections, cover_pprs=cover_pprs)


def classify_paragraph(text: str) -> str:
    if re.match(r"^(摘要|Abstract|目\s*录|绪论|结\s*论|致\s*谢|参考文献)$", text, re.IGNORECASE):
        return "heading1"
    if re.match(r"^第[一二三四五六七八九十0-9]+\s*章", text):
        return "heading1"
    if re.match(r"^\d+\.\d+\.\d+\s*\S*", text):
        return "heading3"
    if re.match(r"^\d+\.\d+\s*\S*", text):
        return "heading2"
    if is_caption_candidate(text):
        return "caption"
    if re.match(r"^\(\d+(?:-\d+){1,2}\)$", text):
        return "formula"
    return "body"


def is_caption_candidate(text: str) -> bool:
    if len(text) > 120:
        return False
    return bool(re.match(r"^(图|表)\s*\d", text))


def normalize_caption_text(text: str) -> str | None:
    if re.match(r"^(图|表)\s*\d+\s*[-－–—]\s*\d+", text):
        return None
    match = re.match(r"^(图|表)\s*(\d)(\d+)(.*)$", text)
    if not match:
        return None
    kind, chapter, serial, rest = match.groups()
    return f"{kind} {chapter}-{serial}{rest}"


def replace_paragraph_text(p: ET.Element, new_text: str) -> None:
    ppr = deepcopy(p.find("w:pPr", {"w": NS["w"]}))
    clear_text_content(p)
    if ppr is not None:
        existing = p.find("w:pPr", {"w": NS["w"]})
        if existing is not None:
            p.remove(existing)
        p.insert(0, ppr)
    append_text_run(p, new_text, east_asia="宋体", ascii_font="Times New Roman", size_half_points="24", bold=False)


def first_body_index(paragraphs: list[ET.Element]) -> int:
    for i, p in enumerate(paragraphs):
        text = norm_space(text_of(p))
        if text == "绪论" or re.match(r"^第一\s*章|^第[一二三四五六七八九十0-9]+\s*章", text):
            return i
    for i, p in enumerate(paragraphs):
        if re.match(r"^1\s*[.．、]", norm_space(text_of(p))):
            return i
    return 0


def format_paragraphs(
    document: ET.Element,
    styles: dict[str, str],
    *,
    cover_paragraph_ids: set[int],
    fix_caption_text: bool,
    finalize_red_to_black: bool,
    strict_body_songti: bool,
) -> None:
    paragraphs = document.findall(".//w:p", {"w": NS["w"]})
    in_references = False
    color = "000000" if finalize_red_to_black else None

    for p in paragraphs:
        if id(p) in cover_paragraph_ids:
            continue
        text = norm_space(text_of(p))
        if not text:
            continue
        if fix_caption_text:
            replacement = normalize_caption_text(text)
            if replacement:
                replace_paragraph_text(p, replacement)
                text = replacement
        kind = classify_paragraph(text)
        if re.match(r"^参考文献$", text):
            in_references = True

        ppr = ensure_ppr(p)
        if kind == "heading1":
            set_p_style(p, styles["heading1"])
            set_jc(ppr, "center")
            set_spacing(ppr, before="340", after="330", line="578")
            set_indent(ppr, clear=True)
            format_paragraph_runs(p, "黑体", "Times New Roman", "36", bold=True, color=color)
        elif kind == "heading2":
            set_p_style(p, styles["heading2"])
            set_jc(ppr, "left")
            set_spacing(ppr, before="260", after="260", line="416")
            set_indent(ppr, clear=True)
            format_paragraph_runs(p, "黑体", "Times New Roman", "30", bold=True, color=color)
        elif kind == "heading3":
            set_p_style(p, styles["heading3"])
            set_jc(ppr, "left")
            set_spacing(ppr, before="260", after="260", line="416")
            set_indent(ppr, clear=True)
            format_paragraph_runs(p, "黑体", "Times New Roman", "24", bold=True, color=color)
        elif kind == "caption":
            set_p_style(p, styles["caption"])
            set_jc(ppr, "center")
            set_spacing(ppr, before="120", after="120", line="300")
            set_indent(ppr, clear=True)
            format_paragraph_runs(p, "黑体", "Times New Roman", "20", bold=False, color=color)
        elif kind == "formula":
            set_p_style(p, styles["normal"])
            set_jc(ppr, "right")
            set_spacing(ppr, before="0", after="0", line="360")
            set_indent(ppr, clear=True)
            format_paragraph_runs(p, "宋体", "Times New Roman", "24", bold=None, color=color)
        elif in_references:
            set_p_style(p, styles["reference"])
            set_spacing(ppr, before="0", after="0", line="360")
            set_indent(ppr, hanging="240", clear=True)
            format_paragraph_runs(p, "宋体", "Times New Roman", "20", bold=False, color=color)
        else:
            set_p_style(p, styles["normal"])
            set_jc(ppr, "both")
            set_spacing(ppr, before="0", after="0", line="360")
            set_indent(ppr, clear=True)
            if strict_body_songti:
                format_paragraph_runs(p, "宋体", "Times New Roman", "24", bold=None, color=color)
            elif finalize_red_to_black:
                format_paragraph_runs(p, "宋体", "Times New Roman", "24", bold=None, color="000000")


def paragraph_text(child: ET.Element) -> str:
    if child.tag != qn("w", "p"):
        return ""
    return norm_space(text_of(child))


def remove_obvious_sample_front_matter(document: ET.Element) -> int:
    body = document.find("w:body", {"w": NS["w"]})
    if body is None:
        return 0
    children = list(body)
    declaration_idx = None
    for i, child in enumerate(children):
        text = paragraph_text(child)
        if "诚信承诺书" in text or "本人郑重声明" in text:
            declaration_idx = i
            break
    if declaration_idx is None:
        return 0

    markers = [
        "学校代码：10616",
        "地球物理学院本科学士学位论文格式规范",
        "地物院",
        "地球物理学",
        "2018.05.30",
        "2018年5月",
        "成 理 工",
    ]
    marker_indices = [
        i
        for i, child in enumerate(children[:declaration_idx])
        if any(marker in paragraph_text(child) for marker in markers)
    ]
    if not marker_indices:
        return 0
    start = min(marker_indices)
    end = max(marker_indices)
    for child in children[start : end + 1]:
        body.remove(child)
    return end - start + 1


def first_cover_nonempty_paragraphs(document: ET.Element) -> list[ET.Element]:
    body = document.find("w:body", {"w": NS["w"]})
    if body is None:
        return []
    result: list[ET.Element] = []
    for child in list(body):
        if child.tag != qn("w", "p"):
            continue
        text = paragraph_text(child)
        if "诚信承诺书" in text or text == "摘要":
            break
        if text:
            result.append(child)
        if len(result) >= 13:
            break
    return result


def split_prefix(text: str, prefix: str) -> str:
    return text[len(prefix) :].strip() if text.startswith(prefix) else ""


def apply_cover_format(document: ET.Element, rules: TemplateRules) -> set[int]:
    cover = first_cover_nonempty_paragraphs(document)
    cover_ids = {id(p) for p in cover}
    label = ("", "宋体", "Times New Roman", "28", True)
    plain_label = ("", "宋体", "Times New Roman", "28", False)
    value = ("", "楷体", "楷体", "28", True)

    for idx, p in enumerate(cover[:13]):
        if idx < len(rules.cover_pprs) and rules.cover_pprs[idx] is not None:
            old_ppr = p.find("w:pPr", {"w": NS["w"]})
            if old_ppr is not None:
                p.remove(old_ppr)
            p.insert(0, deepcopy(rules.cover_pprs[idx]))
        text = paragraph_text(p)
        ppr = ensure_ppr(p)
        if idx == 2:
            set_jc(ppr, "center")
            replace_with_segments(p, [(text, "宋体", "Times New Roman", "84", True)])
        elif idx == 3:
            set_jc(ppr, "center")
            replace_with_segments(p, [(text, "宋体", "Times New Roman", "44", True)])
        elif idx == 4:
            prefix = "题名和副题名"
            suffix = split_prefix(text, prefix)
            replace_with_segments(p, [(prefix, label[1], label[2], label[3], label[4]), (" " + suffix, value[1], value[2], value[3], value[4])])
        elif idx == 5:
            prefix = "作 者 姓 名"
            suffix = split_prefix(text, prefix)
            replace_with_segments(p, [(prefix, label[1], label[2], label[3], label[4]), (" " + suffix, value[1], value[2], value[3], value[4])])
        elif idx == 6:
            prefix = "指导教师姓名及职称"
            suffix = split_prefix(text, prefix)
            replace_with_segments(p, [(prefix, label[1], label[2], label[3], label[4]), (" " + suffix, value[1], value[2], value[3], value[4])])
        elif idx == 7:
            match = re.match(r"^(申请学位级别)\s*(\S+)\s*(专业名称)\s*(.+)$", text)
            if match:
                a, b, c, d = match.groups()
                replace_with_segments(
                    p,
                    [
                        (a + " ", label[1], label[2], label[3], label[4]),
                        (b + " ", value[1], value[2], value[3], value[4]),
                        (c + " ", label[1], label[2], label[3], label[4]),
                        (d, value[1], value[2], value[3], value[4]),
                    ],
                )
            else:
                format_paragraph_runs(p, "宋体", "Times New Roman", "28", bold=True, color="000000")
        elif idx in {0, 1, 12}:
            replace_with_segments(p, [(text, plain_label[1], plain_label[2], plain_label[3], plain_label[4])])
        else:
            replace_with_segments(p, [(text, label[1], label[2], label[3], label[4])])
    return cover_ids


def compact_cover_blank_paragraphs(document: ET.Element) -> None:
    """Remove empty spacer paragraphs on the first cover when a long title wraps.

    The DOTX first cover uses large title typography. Real thesis titles can wrap
    to two lines, so non-content spacer paragraphs can push required metadata to
    page 2. Keep the section-break paragraph and all text-bearing paragraphs.
    """
    body = document.find("w:body", {"w": NS["w"]})
    if body is None:
        return
    for child in list(body):
        if child.tag != qn("w", "p"):
            continue
        ppr = child.find("w:pPr", {"w": NS["w"]})
        if ppr is not None and ppr.find("w:sectPr", {"w": NS["w"]}) is not None:
            break
        text = paragraph_text(child)
        if text:
            if "诚信承诺书" in text or text == "摘要":
                break
            continue
        body.remove(child)


def ensure_front_matter_page_breaks(document: ET.Element) -> None:
    body = document.find("w:body", {"w": NS["w"]})
    if body is not None:
        for child in list(body):
            if child.tag != qn("w", "p") or paragraph_text(child) not in {"摘要", "Abstract"}:
                continue
            current_children = list(body)
            try:
                current_index = current_children.index(child)
            except ValueError:
                continue
            for previous in reversed(current_children[:current_index]):
                if previous.tag != qn("w", "p"):
                    break
                ppr = previous.find("w:pPr", {"w": NS["w"]})
                if ppr is not None and ppr.find("w:sectPr", {"w": NS["w"]}) is not None:
                    break
                if paragraph_text(previous):
                    break
                body.remove(previous)
    for p in document.findall(".//w:p", {"w": NS["w"]}):
        text = paragraph_text(p)
        if text in {"摘要", "Abstract"}:
            ppr = ensure_ppr(p)
            page_break = ensure_child(ppr, qn("w", "pageBreakBefore"), 0)
            page_break.set(w_attr("val"), "1")


def set_page_break_before(p: ET.Element) -> None:
    ppr = ensure_ppr(p)
    page_break = ensure_child(ppr, qn("w", "pageBreakBefore"), 0)
    page_break.set(w_attr("val"), "1")


def ensure_terminal_section_page_breaks(document: ET.Element) -> None:
    """Keep standalone terminal sections from starting at the foot of a page."""
    for p in document.findall(".//w:p", {"w": NS["w"]}):
        compact = re.sub(r"\s+", "", paragraph_text(p))
        if compact in {"致谢", "参考文献"}:
            set_page_break_before(p)


def remove_empty_paragraphs_before_body(document: ET.Element) -> None:
    body = document.find("w:body", {"w": NS["w"]})
    body_idx = body_heading_child_index(document)
    if body is None or body_idx is None:
        return
    for child in reversed(list(body)[:body_idx]):
        if child.tag != qn("w", "p"):
            break
        ppr = child.find("w:pPr", {"w": NS["w"]})
        if ppr is not None and ppr.find("w:sectPr", {"w": NS["w"]}) is not None:
            break
        if paragraph_text(child):
            break
        body.remove(child)


def tighten_english_abstract_heading(document: ET.Element) -> None:
    """Avoid a one-line English-keywords orphan while preserving thesis text."""
    for p in document.findall(".//w:p", {"w": NS["w"]}):
        if paragraph_text(p).strip() != "Abstract":
            continue
        ppr = ensure_ppr(p)
        set_spacing(ppr, before="240", after="120", line="360")
        set_jc(ppr, "center")
        break


def set_section_page_setup(sect: ET.Element, setup: dict[str, str] | None = None) -> None:
    setup = setup or {
        "w": "11906",
        "h": "16838",
        "top": "1701",
        "right": "1800",
        "bottom": "1440",
        "left": "1800",
        "header": "851",
        "footer": "992",
        "gutter": "0",
    }
    pg_sz = ensure_child(sect, qn("w", "pgSz"))
    pg_sz.set(w_attr("w"), setup["w"])
    pg_sz.set(w_attr("h"), setup["h"])
    pg_mar = ensure_child(sect, qn("w", "pgMar"))
    for key in ("top", "right", "bottom", "left", "header", "footer", "gutter"):
        pg_mar.set(w_attr(key), setup[key])


def body_heading_child_index(document: ET.Element) -> int | None:
    body = document.find("w:body", {"w": NS["w"]})
    if body is None:
        return None
    for idx, child in enumerate(list(body)):
        if child.tag != qn("w", "p"):
            continue
        text = paragraph_text(child)
        if text == "绪论" or re.match(r"^第一\s*章|^第[一二三四五六七八九十0-9]+\s*章", text):
            return idx
    return None


def ensure_body_section_break(document: ET.Element) -> None:
    body = document.find("w:body", {"w": NS["w"]})
    body_idx = body_heading_child_index(document)
    if body is None or body_idx is None or body_idx <= 0:
        return
    previous_paragraph = None
    for child in reversed(list(body)[:body_idx]):
        if child.tag == qn("w", "p"):
            previous_paragraph = child
            break
    if previous_paragraph is None:
        return
    ppr = ensure_ppr(previous_paragraph)
    sect = ppr.find("w:sectPr", {"w": NS["w"]})
    if sect is None:
        sect = ET.SubElement(ppr, qn("w", "sectPr"))
    for child in list(sect):
        if child.tag in {qn("w", "headerReference"), qn("w", "footerReference"), qn("w", "pgNumType")}:
            sect.remove(child)


def setup_for_section(index: int, count: int, rules: TemplateRules) -> dict[str, str] | None:
    if not rules.sections:
        return None
    if count == 1:
        return rules.sections[-1]
    if index == 0:
        return rules.sections[0]
    if index == count - 1:
        return rules.sections[-1]
    return rules.sections[min(index, len(rules.sections) - 2)]


def normalize_sections(document: ET.Element, header_rid: str, footer_rid: str, rules: TemplateRules) -> None:
    ensure_body_section_break(document)
    sections = document.findall(".//w:sectPr", {"w": NS["w"]})
    if not sections:
        body = document.find("w:body", {"w": NS["w"]})
        if body is None:
            return
        sect = ET.SubElement(body, qn("w", "sectPr"))
        sections = [sect]
    for index, sect in enumerate(sections):
        set_section_page_setup(sect, setup_for_section(index, len(sections), rules))

    body_section = sections[-1]
    for child in list(body_section):
        if child.tag in {qn("w", "headerReference"), qn("w", "footerReference"), qn("w", "pgNumType")}:
            body_section.remove(child)
    header_ref = ET.Element(qn("w", "headerReference"), {w_attr("type"): "default", qn("r", "id"): header_rid})
    footer_ref = ET.Element(qn("w", "footerReference"), {w_attr("type"): "default", qn("r", "id"): footer_rid})
    page_start = ET.Element(qn("w", "pgNumType"), {w_attr("start"): "1"})
    body_section.insert(0, footer_ref)
    body_section.insert(0, header_ref)
    body_section.insert(2, page_start)


def next_part_name(existing: set[str], prefix: str) -> str:
    nums: list[int] = []
    pattern = re.compile(rf"word/{prefix}(\d+)\.xml$")
    for name in existing:
        match = pattern.match(name)
        if match:
            nums.append(int(match.group(1)))
    number = max(nums, default=0) + 1
    candidate = f"word/{prefix}{number}.xml"
    while candidate in existing:
        number += 1
        candidate = f"word/{prefix}{number}.xml"
    return candidate


def next_rid(rels: ET.Element) -> str:
    nums: list[int] = []
    for rel in rels.findall(f"{{{NS['rel']}}}Relationship"):
        rid = rel.get("Id", "")
        match = re.match(r"rId(\d+)$", rid)
        if match:
            nums.append(int(match.group(1)))
    return f"rId{max(nums, default=0) + 1}"


def add_relationship(rels: ET.Element, rid: str, rel_type: str, target: str) -> None:
    rel = ET.Element(f"{{{NS['rel']}}}Relationship")
    rel.set("Id", rid)
    rel.set("Type", rel_type)
    rel.set("Target", target)
    rels.append(rel)


def ensure_content_override_text(xml_text: str, part_name: str, content_type: str) -> str:
    part = f"/{part_name}"
    pattern = re.compile(rf'(<Override\b[^>]*PartName="{re.escape(part)}"[^>]*/>)')
    replacement = f'<Override PartName="{part}" ContentType="{content_type}"/>'
    if pattern.search(xml_text):
        return pattern.sub(replacement, xml_text, count=1)
    close_match = re.search(r"</(?:[A-Za-z0-9_]+:)?Types\s*>", xml_text)
    if not close_match:
        raise RuntimeError("Cannot find closing Types element in [Content_Types].xml")
    return xml_text[: close_match.start()] + replacement + xml_text[close_match.start() :]


def infer_year(document: ET.Element) -> str:
    text = text_of(document)
    match = re.search(r"(20\d{2})\s*年\s*\d{1,2}\s*月", text)
    if match:
        return match.group(1)
    match = re.search(r"\b(20\d{2})\b", text)
    if match:
        return match.group(1)
    return "****"


def make_header_xml(header_text: str) -> bytes:
    root = ET.Element(qn("w", "hdr"))
    p = ET.SubElement(root, qn("w", "p"))
    ppr = ensure_ppr(p)
    set_jc(ppr, "center")
    r = ET.SubElement(p, qn("w", "r"))
    rpr = ET.SubElement(r, qn("w", "rPr"))
    set_run_format(rpr, "宋体", "Times New Roman", "18", bold=False)
    t = ET.SubElement(r, qn("w", "t"))
    t.text = header_text
    return to_xml_bytes(root)


def make_footer_xml() -> bytes:
    root = ET.Element(qn("w", "ftr"))
    p = ET.SubElement(root, qn("w", "p"))
    ppr = ensure_ppr(p)
    set_jc(ppr, "center")
    for tag, value in (("begin", None), ("instr", " PAGE "), ("end", None)):
        r = ET.SubElement(p, qn("w", "r"))
        rpr = ET.SubElement(r, qn("w", "rPr"))
        set_run_format(rpr, "宋体", "Times New Roman", "18", bold=False)
        if tag == "instr":
            instr = ET.SubElement(r, qn("w", "instrText"))
            instr.set(f"{{{XML_NS}}}space", "preserve")
            instr.text = value
        else:
            fld = ET.SubElement(r, qn("w", "fldChar"))
            fld.set(w_attr("fldCharType"), tag)
    return to_xml_bytes(root)


def repair_docx(
    input_path: Path,
    output_path: Path,
    *,
    force: bool,
    template_dotx: Path | None,
    template_first: bool,
    finalize_red_to_black: bool,
    fix_caption_text: bool,
    repair_front_matter: bool,
    strict_body_songti: bool,
    header_text: str | None,
) -> None:
    if input_path.resolve() == output_path.resolve():
        raise SystemExit("Refusing to write repaired output over the input file.")
    if output_path.exists() and not force:
        raise SystemExit(f"Output already exists: {output_path}. Use --force to replace it.")
    if input_path.suffix.lower() != ".docx" or not zipfile.is_zipfile(input_path):
        raise SystemExit("Input must be a valid .docx file.")

    rules = read_template_rules(template_dotx if template_first else None)
    with zipfile.ZipFile(input_path, "r") as zin:
        existing = set(zin.namelist())
        document = read_xml(zin, "word/document.xml")
        styles = read_xml(zin, "word/styles.xml")
        rels = read_xml(zin, "word/_rels/document.xml.rels")
        content_types_text = zin.read("[Content_Types].xml").decode("utf-8")

        if repair_front_matter:
            remove_obvious_sample_front_matter(document)
        cover_ids = apply_cover_format(document, rules)
        compact_cover_blank_paragraphs(document)
        style_ids = normalize_styles(styles)
        format_paragraphs(
            document,
            style_ids,
            cover_paragraph_ids=cover_ids,
            fix_caption_text=fix_caption_text,
            finalize_red_to_black=finalize_red_to_black,
            strict_body_songti=strict_body_songti,
        )
        ensure_front_matter_page_breaks(document)
        ensure_terminal_section_page_breaks(document)
        remove_empty_paragraphs_before_body(document)
        tighten_english_abstract_heading(document)

        header_part = next_part_name(existing, "header")
        existing.add(header_part)
        footer_part = next_part_name(existing, "footer")
        header_rid = next_rid(rels)
        add_relationship(rels, header_rid, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", header_part.replace("word/", ""))
        footer_rid = next_rid(rels)
        add_relationship(rels, footer_rid, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", footer_part.replace("word/", ""))
        normalized_header = header_text or f"地球物理学院 {infer_year(document)}届本科毕业生学士学位论文"
        normalize_sections(document, header_rid, footer_rid, rules)

        content_types_text = ensure_content_override_text(
            content_types_text,
            header_part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        )
        content_types_text = ensure_content_override_text(
            content_types_text,
            footer_part,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        )

        replacements = {
            "word/document.xml": to_xml_bytes(document),
            "word/styles.xml": to_xml_bytes(styles),
            "word/_rels/document.xml.rels": to_xml_bytes(rels),
            "[Content_Types].xml": content_types_text.encode("utf-8"),
            header_part: make_header_xml(normalized_header),
            footer_part: make_footer_xml(),
        }

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in replacements:
                    zout.writestr(item, replacements.pop(item.filename))
                else:
                    zout.writestr(item, zin.read(item.filename))
            for name, data in replacements.items():
                zout.writestr(name, data)


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_格式修复{input_path.suffix}")


def parse_args(argv: Iterable[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path, help="Path to the input .docx file.")
    parser.add_argument("--out", type=Path, help="Output .docx path. Defaults to <name>_格式修复.docx.")
    parser.add_argument("--force", action="store_true", help="Replace the output path if it already exists.")
    parser.add_argument("--template-dotx", type=Path, default=default_template_path(), help="DOTX template path. Defaults to the bundled geophysics template.")
    parser.add_argument("--template-first", action="store_true", default=True, help="Use DOTX-derived rules as the highest priority.")
    parser.add_argument("--finalize-red-to-black", action="store_true", help="Convert directly formatted red text to black.")
    parser.add_argument("--fix-caption-text", action="store_true", help="Normalize compact caption labels such as 图 21 to 图 2-1.")
    parser.add_argument("--repair-front-matter", dest="repair_front_matter", action="store_true", default=True, help="Remove obvious template sample front matter and repair the real cover.")
    parser.add_argument("--no-repair-front-matter", dest="repair_front_matter", action="store_false", help="Skip sample-front-matter cleanup.")
    parser.add_argument("--strict-body-songti", dest="strict_body_songti", action="store_true", default=True, help="Force regular body paragraphs to 宋体小四.")
    parser.add_argument("--no-strict-body-songti", dest="strict_body_songti", action="store_false", help="Do not force regular body paragraphs to 宋体小四.")
    parser.add_argument("--header-text", help="Override the inferred body header text.")
    return parser.parse_args(list(argv))


def main(argv: Iterable[str] = sys.argv[1:]) -> int:
    args = parse_args(argv)
    input_path = args.docx
    output_path = args.out or default_output_path(input_path)
    repair_docx(
        input_path,
        output_path,
        force=args.force,
        template_dotx=args.template_dotx,
        template_first=args.template_first,
        finalize_red_to_black=args.finalize_red_to_black,
        fix_caption_text=args.fix_caption_text,
        repair_front_matter=args.repair_front_matter,
        strict_body_songti=args.strict_body_songti,
        header_text=args.header_text,
    )
    print(f"Wrote repaired copy: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
