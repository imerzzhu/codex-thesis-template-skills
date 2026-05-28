#!/usr/bin/env python3
"""Read-only OOXML audit for CDUT Geophysics College thesis DOCX formatting."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def qn(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def attr(el: ET.Element | None, name: str) -> str | None:
    if el is None:
        return None
    return el.get(qn("w", name))


def twips_to_cm(value: str | None) -> float | None:
    if not value:
        return None
    try:
        return round(int(value) / 1440 * 2.54, 2)
    except ValueError:
        return None


def half_points_to_points(value: str | None) -> float | None:
    if not value:
        return None
    try:
        return int(value) / 2
    except ValueError:
        return None


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


def md_escape(value: str) -> str:
    return value.replace("|", "\\|").replace("\n", " ")


def default_template_path() -> Path:
    return Path(__file__).resolve().parents[1] / "assets" / "地球物理学院本科学士学位论文（设计）学术标准及基本规范_附件一（论文格式模板）-2018.dotx"


@dataclass
class Issue:
    severity: str
    code: str
    message: str
    location: str = ""


@dataclass
class Paragraph:
    index: int
    text: str
    style_id: str | None
    style_name: str | None
    align: str | None
    sizes: list[float]
    fonts: list[str]
    colors: list[str]
    bold: bool
    has_drawing: bool
    has_page_break_before: bool
    body_child_index: int | None = None


class DocxAudit:
    def __init__(self, path: Path, template_dotx: Path | None = None, template_first: bool = True) -> None:
        self.path = path
        self.template_dotx = template_dotx if template_first else None
        self.issues: list[Issue] = []
        self.paragraphs: list[Paragraph] = []
        self.sections: list[dict] = []
        self.template_sections: list[dict[str, int]] = []
        self.headers: list[str] = []
        self.footers: list[str] = []
        self.has_page_field = False
        self.has_toc_field = False
        self.style_names: dict[str, str] = {}
        self.style_ids_by_name: dict[str, str] = {}
        self.style_sizes: dict[str, float] = {}
        self.style_fonts: dict[str, list[str]] = {}

    def add(self, severity: str, code: str, message: str, location: str = "") -> None:
        self.issues.append(Issue(severity, code, message, location))

    def read_xml(self, zf: zipfile.ZipFile, name: str) -> ET.Element | None:
        try:
            return ET.fromstring(zf.read(name))
        except KeyError:
            return None
        except ET.ParseError as exc:
            self.add("blocking", "xml-parse-error", f"Cannot parse {name}: {exc}")
            return None

    def run(self) -> dict:
        if self.path.suffix.lower() != ".docx":
            self.add("blocking", "input-not-docx", "Audit input must be a .docx file.")
            return self.report()
        if not zipfile.is_zipfile(self.path):
            self.add("blocking", "invalid-docx", "File is not a valid DOCX zip package.")
            return self.report()

        self.parse_template_sections()
        with zipfile.ZipFile(self.path) as zf:
            document = self.read_xml(zf, "word/document.xml")
            if document is None:
                self.add("blocking", "missing-document-xml", "word/document.xml is missing.")
                return self.report()
            self.parse_styles(zf)
            self.parse_paragraphs(document)
            self.parse_sections(document)
            self.parse_headers_and_footers(zf, document)

        self.check_page_setup()
        self.check_styles()
        self.check_cover()
        self.check_required_sections()
        self.check_toc()
        self.check_headers_and_footers()
        self.check_body_songti()
        self.check_captions()
        self.check_references()
        self.check_placeholders()
        return self.report()

    def parse_template_sections(self) -> None:
        path = self.template_dotx or default_template_path()
        if not path.exists() or not zipfile.is_zipfile(path):
            return
        with zipfile.ZipFile(path) as zf:
            document = self.read_xml(zf, "word/document.xml")
            if document is None:
                return
            for sect in document.findall(".//w:sectPr", NS):
                pg_sz = sect.find("w:pgSz", NS)
                pg_mar = sect.find("w:pgMar", NS)
                if pg_sz is None or pg_mar is None:
                    continue
                self.template_sections.append(
                    {
                        "width_twips": int(attr(pg_sz, "w") or 0),
                        "height_twips": int(attr(pg_sz, "h") or 0),
                        "top": int(attr(pg_mar, "top") or 0),
                        "right": int(attr(pg_mar, "right") or 0),
                        "bottom": int(attr(pg_mar, "bottom") or 0),
                        "left": int(attr(pg_mar, "left") or 0),
                        "header": int(attr(pg_mar, "header") or 0),
                        "footer": int(attr(pg_mar, "footer") or 0),
                    }
                )

    def parse_styles(self, zf: zipfile.ZipFile) -> None:
        styles = self.read_xml(zf, "word/styles.xml")
        if styles is None:
            self.add("warning", "missing-styles", "word/styles.xml is missing; style checks will be limited.")
            return
        for style in styles.findall("w:style", NS):
            style_id = attr(style, "styleId")
            name_el = style.find("w:name", NS)
            name = attr(name_el, "val")
            if style_id and name:
                self.style_names[style_id] = name
                self.style_ids_by_name[name.lower()] = style_id
            rpr = style.find("w:rPr", NS)
            if rpr is None or not style_id:
                continue
            size = half_points_to_points(attr(rpr.find("w:sz", NS), "val"))
            if size:
                self.style_sizes[style_id] = size
            fonts: list[str] = []
            rfonts = rpr.find("w:rFonts", NS)
            if rfonts is not None:
                for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                    value = attr(rfonts, key)
                    if value and value not in fonts:
                        fonts.append(value)
            if fonts:
                self.style_fonts[style_id] = fonts

    def parse_paragraphs(self, document: ET.Element) -> None:
        body = document.find("w:body", NS)
        index = 0
        if body is None:
            return
        for child_index, child in enumerate(list(body)):
            if child.tag != qn("w", "p"):
                continue
            text = norm_space(text_of(child))
            if text:
                index += 1
            ppr = child.find("w:pPr", NS)
            style_id = attr(ppr.find("w:pStyle", NS), "val") if ppr is not None else None
            align = attr(ppr.find("w:jc", NS), "val") if ppr is not None else None
            has_page_break_before = ppr is not None and ppr.find("w:pageBreakBefore", NS) is not None
            sizes: list[float] = []
            fonts: list[str] = []
            colors: list[str] = []
            bold = False
            for r in child.findall(".//w:r", NS):
                rpr = r.find("w:rPr", NS)
                if rpr is None:
                    continue
                size = half_points_to_points(attr(rpr.find("w:sz", NS), "val"))
                if size and size not in sizes:
                    sizes.append(size)
                rfonts = rpr.find("w:rFonts", NS)
                if rfonts is not None:
                    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                        value = attr(rfonts, key)
                        if value and value not in fonts:
                            fonts.append(value)
                color = attr(rpr.find("w:color", NS), "val")
                if color and color not in colors:
                    colors.append(color)
                if rpr.find("w:b", NS) is not None:
                    bold = True
            if text:
                self.paragraphs.append(
                    Paragraph(
                        index=index,
                        text=text,
                        style_id=style_id,
                        style_name=self.style_names.get(style_id or ""),
                        align=align,
                        sizes=sizes,
                        fonts=fonts,
                        colors=colors,
                        bold=bold,
                        has_drawing=child.find(".//w:drawing", NS) is not None or child.find(".//w:pict", NS) is not None,
                        has_page_break_before=has_page_break_before,
                        body_child_index=child_index,
                    )
                )
        self.has_toc_field = any(node.tag == qn("w", "instrText") and node.text and "TOC" in node.text.upper() for node in document.iter())

    def parse_sections(self, document: ET.Element) -> None:
        for index, sect in enumerate(document.findall(".//w:sectPr", NS), start=1):
            pg_sz = sect.find("w:pgSz", NS)
            pg_mar = sect.find("w:pgMar", NS)
            self.sections.append(
                {
                    "index": index,
                    "width_twips": int(attr(pg_sz, "w") or 0) if pg_sz is not None else 0,
                    "height_twips": int(attr(pg_sz, "h") or 0) if pg_sz is not None else 0,
                    "top_twips": int(attr(pg_mar, "top") or 0) if pg_mar is not None else 0,
                    "right_twips": int(attr(pg_mar, "right") or 0) if pg_mar is not None else 0,
                    "bottom_twips": int(attr(pg_mar, "bottom") or 0) if pg_mar is not None else 0,
                    "left_twips": int(attr(pg_mar, "left") or 0) if pg_mar is not None else 0,
                    "header_twips": int(attr(pg_mar, "header") or 0) if pg_mar is not None else 0,
                    "footer_twips": int(attr(pg_mar, "footer") or 0) if pg_mar is not None else 0,
                    "top_cm": twips_to_cm(attr(pg_mar, "top")) if pg_mar is not None else None,
                    "bottom_cm": twips_to_cm(attr(pg_mar, "bottom")) if pg_mar is not None else None,
                    "left_cm": twips_to_cm(attr(pg_mar, "left")) if pg_mar is not None else None,
                    "right_cm": twips_to_cm(attr(pg_mar, "right")) if pg_mar is not None else None,
                    "header_cm": twips_to_cm(attr(pg_mar, "header")) if pg_mar is not None else None,
                    "footer_cm": twips_to_cm(attr(pg_mar, "footer")) if pg_mar is not None else None,
                }
            )

    def parse_headers_and_footers(self, zf: zipfile.ZipFile, document: ET.Element) -> None:
        rels = self.read_xml(zf, "word/_rels/document.xml.rels")
        if rels is None:
            return
        rel_map: dict[str, str] = {}
        for rel in rels.findall("rel:Relationship", NS):
            rid = rel.get("Id")
            target = rel.get("Target")
            if rid and target:
                rel_map[rid] = target.lstrip("/")
        for ref_tag, bucket in (("headerReference", self.headers), ("footerReference", self.footers)):
            for ref in document.findall(f".//w:{ref_tag}", NS):
                rid = ref.get(qn("r", "id"))
                target = rel_map.get(rid or "")
                if not target:
                    continue
                part = target if target.startswith("word/") else f"word/{target}"
                xml = self.read_xml(zf, part)
                if xml is not None:
                    value = norm_space(text_of(xml))
                    if value:
                        bucket.append(value)
                    if any(node.tag == qn("w", "instrText") and node.text and "PAGE" in node.text.upper() for node in xml.iter()):
                        self.has_page_field = True

    def template_setup_for(self, index: int, count: int) -> dict[str, int] | None:
        if not self.template_sections:
            return None
        if count == 1:
            return self.template_sections[-1]
        if index == 0:
            return self.template_sections[0]
        if index == count - 1:
            return self.template_sections[-1]
        return self.template_sections[min(index, len(self.template_sections) - 2)]

    def check_page_setup(self) -> None:
        if not self.sections:
            self.add("warning", "no-sections", "No section properties found; page setup cannot be verified.")
            return
        for idx, section in enumerate(self.sections):
            loc = f"section {section['index']}"
            width = section["width_twips"]
            height = section["height_twips"]
            if abs(width - 11906) > 120 or abs(height - 16838) > 120:
                self.add("blocking", "page-size", f"Expected A4 portrait page size near 11906 x 16838 twips, found {width} x {height}.", loc)
            expected = self.template_setup_for(idx, len(self.sections))
            if expected:
                for key in ("top", "right", "bottom", "left", "header", "footer"):
                    found = section[f"{key}_twips"]
                    if expected.get(key) and abs(found - expected[key]) > 120:
                        self.add("warning", "template-margin", f"{key} is {found} twips; template expects near {expected[key]}.", loc)

    def check_styles(self) -> None:
        for required in ("normal", "heading 1", "heading 2", "heading 3"):
            if required not in self.style_ids_by_name:
                self.add("warning", "missing-style", f"Expected Word style not found: {required}.")
        normal_id = self.style_ids_by_name.get("normal")
        if normal_id:
            size = self.style_sizes.get(normal_id)
            if size and abs(size - 12) > 0.5:
                self.add("warning", "normal-size", f"Normal style is {size} pt; expected 小四 / 12 pt.")
            fonts = " ".join(self.style_fonts.get(normal_id, []))
            if fonts and not re.search(r"宋体|SimSun", fonts, re.IGNORECASE):
                self.add("warning", "normal-font", f"Normal style fonts do not show 宋体/SimSun: {fonts}.")

    def check_cover(self) -> None:
        first = self.paragraphs[:15]
        if len(first) < 13:
            self.add("warning", "cover-short", "Could not find the expected first cover metadata block.")
            return
        school = next((p for p in first if "成 都 理 工 大 学" in p.text), None)
        if school and not any(abs(size - 42) <= 1 for size in school.sizes):
            self.add("warning", "cover-school-size", "Cover school name should follow the DOTX 42 pt format.", f"paragraph {school.index}")
        degree = next((p for p in first if "学 位 论 文" in p.text), None)
        if degree and not any(abs(size - 22) <= 1 for size in degree.sizes):
            self.add("warning", "cover-degree-size", "Cover `学 位 论 文` should follow the DOTX 22 pt format.", f"paragraph {degree.index}")
        if any("地物院" in p.text or "2018年5月" in p.text for p in self.paragraphs[:40]):
            self.add("blocking", "sample-cover-leftover", "Template sample cover/front-matter text remains near the document front.")

    def first_para_index(self, pattern: str) -> int | None:
        rx = re.compile(pattern, re.IGNORECASE)
        for i, p in enumerate(self.paragraphs):
            if rx.search(p.text):
                return i
        return None

    def check_required_sections(self) -> None:
        required = [
            ("诚信承诺书", r"诚信承诺书"),
            ("摘要", r"^摘\s*要$"),
            ("Abstract", r"^Abstract$"),
            ("body chapter", r"^(绪论|第[一二三四五六七八九十0-9]+\s*章)"),
            ("结论", r"^结\s*论$|^第[一二三四五六七八九十0-9]+\s*章\s*结论"),
            ("致谢", r"^致\s*谢$"),
            ("参考文献", r"^参考文献$"),
        ]
        found: list[tuple[str, int | None]] = []
        for label, pattern in required:
            idx = self.first_para_index(pattern)
            found.append((label, idx))
            if idx is None:
                self.add("warning", "missing-required-section", f"Expected section title not found: {label}.")
        ordered = [(label, idx) for label, idx in found if idx is not None]
        for (left_label, left_idx), (right_label, right_idx) in zip(ordered, ordered[1:]):
            if left_idx is not None and right_idx is not None and left_idx > right_idx:
                self.add("warning", "section-order", f"`{left_label}` appears after `{right_label}`; expected template order is likely wrong.")
        for section_name in {"致谢", "参考文献"}:
            compact_name = re.sub(r"\s+", "", section_name)
            para = next((p for p in self.paragraphs if re.sub(r"\s+", "", p.text) == compact_name), None)
            if para is not None and not para.has_page_break_before:
                self.add(
                    "warning",
                    "terminal-section-pagebreak",
                    f"`{section_name}` should start on a new page in the DOTX-first layout.",
                    f"paragraph {para.index}",
                )

    def check_toc(self) -> None:
        toc_idx = self.first_para_index(r"^目\s*录$")
        if toc_idx is None:
            self.add("warning", "missing-toc", "目录 heading was not detected.")
            return
        if not self.has_toc_field:
            self.add("warning", "static-toc", "目录 exists, but no Word TOC field was detected. Update or regenerate the TOC in Word after heading repairs.")

    def check_headers_and_footers(self) -> None:
        header_text = " ".join(self.headers)
        if not self.headers:
            self.add("warning", "missing-header", "No running header part was detected. Body pages should use the DOTX thesis header.")
        elif "地球物理学院" not in header_text or "本科毕业生学士学位论文" not in header_text:
            self.add("warning", "header-text", "Running header does not match the DOTX geophysics college header wording.")
        if not self.footers:
            self.add("warning", "missing-footer", "No footer part was detected. Page numbers should be bottom-centered.")
        elif not self.has_page_field:
            self.add("warning", "missing-page-field", "Footer exists, but no PAGE field was detected.")

    def heading_kind(self, text: str) -> str:
        if re.match(r"^(摘要|Abstract|目录|绪论|结\s*论|致\s*谢|参考文献)$", text) or re.match(r"^第[一二三四五六七八九十0-9]+\s*章", text):
            return "heading"
        if re.match(r"^\d+\.\d+", text):
            return "heading"
        return "body"

    def body_range(self) -> tuple[int | None, int | None]:
        start = self.first_para_index(r"^(绪论|第[一二三四五六七八九十0-9]+\s*章)")
        end = self.first_para_index(r"^参考文献$")
        return start, end

    def check_body_songti(self) -> None:
        start, end = self.body_range()
        if start is None:
            return
        checked = 0
        bad: list[Paragraph] = []
        for p in self.paragraphs[start : end or len(self.paragraphs)]:
            text = p.text
            if self.heading_kind(text) == "heading" or is_caption_candidate(text) or re.match(r"^\(\d+(?:-\d+){1,2}\)$", text):
                continue
            if len(text) < 8:
                continue
            checked += 1
            fonts = " ".join(p.fonts + self.style_fonts.get(p.style_id or "", []))
            sizes = p.sizes or ([self.style_sizes[p.style_id]] if p.style_id in self.style_sizes else [])
            if (fonts and not re.search(r"宋体|SimSun", fonts, re.IGNORECASE)) or (sizes and not any(abs(size - 12) <= 0.5 for size in sizes)):
                bad.append(p)
                if len(bad) >= 5:
                    break
        if bad:
            sample = "; ".join(f"p{p.index}:{p.text[:24]}" for p in bad)
            self.add("warning", "body-songti-small4", f"Regular body paragraphs should be 宋体小四; sample deviations: {sample}")
        elif checked == 0:
            self.add("info", "body-songti-small4", "No regular body paragraphs were sampled for 宋体小四 checking.")

    def check_captions(self) -> None:
        caption_re = re.compile(r"^(图|表)\s*([0-9]+)\s*[-－–—]\s*([0-9]+)")
        compact_re = re.compile(r"^(图|表)\s*([0-9]{2,})(?!\s*[-－–—0-9])")
        seen: set[tuple[str, str, str]] = set()
        by_chapter: dict[tuple[str, str], list[int]] = {}
        for p in self.paragraphs:
            if not is_caption_candidate(p.text):
                continue
            compact = compact_re.match(p.text)
            if compact:
                self.add("warning", "compact-caption-number", f"Caption looks compact or missing a hyphen: `{p.text[:60]}`.", f"paragraph {p.index}")
            match = caption_re.match(p.text)
            if not match:
                continue
            kind, chapter, serial = match.groups()
            key = (kind, chapter, serial)
            if key in seen:
                self.add("warning", "duplicate-caption-number", f"Repeated caption number {kind}{chapter}-{serial}.", f"paragraph {p.index}")
            seen.add(key)
            by_chapter.setdefault((kind, chapter), []).append(int(serial))
            if p.align not in {None, "center"}:
                self.add("warning", "caption-alignment", f"{kind} caption should be centered.", f"paragraph {p.index}")
            if p.sizes and not any(abs(size - 10) <= 0.5 for size in p.sizes):
                self.add("info", "caption-size", f"{kind} caption font size is not visibly 五号 / 10 pt.", f"paragraph {p.index}")
        for (kind, chapter), values in by_chapter.items():
            unique = sorted(set(values))
            if len(unique) > 1:
                expected = list(range(unique[0], unique[-1] + 1))
                if unique != expected:
                    self.add("warning", "caption-gap", f"{kind} captions in chapter {chapter} are not continuous: {unique}.")

    def check_references(self) -> None:
        start = self.first_para_index(r"^参考文献$")
        if start is None:
            return
        refs = [p for p in self.paragraphs[start + 1 :] if p.text]
        if not refs:
            self.add("warning", "empty-references", "参考文献 section is present but contains no entries.")
            return
        numbered = [p for p in refs if re.match(r"^\[\d+\]", p.text)]
        if numbered:
            self.add("info", "numbered-references", "Numbered reference entries were detected. Confirm whether this thesis should use numbered GB/T style.")
        for p in refs[:20]:
            if p.sizes and not any(abs(size - 10) <= 0.5 for size in p.sizes):
                self.add("info", "reference-size", "Reference entry does not appear to use 五号 / 10 pt.", f"paragraph {p.index}")
                break

    def check_placeholders(self) -> None:
        placeholder_re = re.compile(r"(××|□|只做格式样例|格式有关问题说明|百度学术搜索界面截图|文献检索截图|地物院 教授|2018\.05\.30)")
        for p in self.paragraphs:
            if placeholder_re.search(p.text):
                self.add("warning", "placeholder-text", f"Template placeholder or instruction text may remain: `{p.text[:80]}`.", f"paragraph {p.index}")
        red_paras = [p for p in self.paragraphs if any(c.upper() in {"FF0000", "C00000", "FF3333"} for c in p.colors)]
        if red_paras:
            sample = "; ".join(p.text[:40] for p in red_paras[:5])
            self.add("warning", "red-text", f"Red text remains; final submissions normally require resolving or changing it to black. Sample: {sample}")

    def report(self) -> dict:
        counts: dict[str, int] = {}
        for issue in self.issues:
            counts[issue.severity] = counts.get(issue.severity, 0) + 1
        major_headings = [
            {"paragraph": p.index, "text": p.text}
            for p in self.paragraphs
            if re.search(r"^(摘要|Abstract|目\s*录|绪论|第[一二三四五六七八九十0-9]+\s*章|结\s*论|致\s*谢|参考文献)$", p.text)
        ]
        return {
            "file": str(self.path),
            "template_dotx": str(self.template_dotx or default_template_path()),
            "issue_counts": counts,
            "issues": [asdict(issue) for issue in self.issues],
            "sections": self.sections,
            "headers": self.headers,
            "footers": self.footers,
            "has_toc_field": self.has_toc_field,
            "has_page_field": self.has_page_field,
            "major_headings": major_headings,
        }


def is_caption_candidate(text: str) -> bool:
    if len(text) > 120:
        return False
    return bool(re.match(r"^(图|表)\s*\d", text))


def render_markdown(report: dict) -> str:
    lines: list[str] = []
    lines.append("# CDUT Geophysics Thesis DOCX Audit")
    lines.append("")
    lines.append(f"- File: `{report['file']}`")
    lines.append(f"- Template: `{report.get('template_dotx', '')}`")
    counts = report.get("issue_counts", {})
    lines.append(f"- Issues: blocking={counts.get('blocking', 0)}, warning={counts.get('warning', 0)}, info={counts.get('info', 0)}")
    lines.append(f"- TOC field detected: `{report.get('has_toc_field')}`")
    lines.append(f"- PAGE field detected: `{report.get('has_page_field')}`")
    lines.append("")
    lines.append("## Issues")
    issues = report.get("issues", [])
    if not issues:
        lines.append("")
        lines.append("No issues detected by the structural audit. Still render the document and inspect pages visually.")
    else:
        lines.append("")
        lines.append("| Severity | Code | Location | Message |")
        lines.append("| --- | --- | --- | --- |")
        for issue in issues:
            lines.append(f"| {md_escape(issue['severity'])} | {md_escape(issue['code'])} | {md_escape(issue.get('location', ''))} | {md_escape(issue['message'])} |")
    lines.append("")
    lines.append("## Sections")
    lines.append("")
    if report.get("sections"):
        lines.append("| Section | Page twips | Margins cm top/right/bottom/left | Header/Footer cm |")
        lines.append("| --- | --- | --- | --- |")
        for section in report["sections"]:
            margins = f"{section.get('top_cm')}/{section.get('right_cm')}/{section.get('bottom_cm')}/{section.get('left_cm')}"
            hf = f"{section.get('header_cm')}/{section.get('footer_cm')}"
            lines.append(f"| {section['index']} | {section['width_twips']} x {section['height_twips']} | {margins} | {hf} |")
    else:
        lines.append("No section properties were found.")
    lines.append("")
    lines.append("## Detected Major Headings")
    lines.append("")
    if report.get("major_headings"):
        for heading in report["major_headings"]:
            lines.append(f"- Paragraph {heading['paragraph']}: {heading['text']}")
    else:
        lines.append("No major headings detected.")
    lines.append("")
    lines.append("## Reminder")
    lines.append("")
    lines.append("This audit is structural. Always render the repaired DOCX and inspect page images before delivery.")
    return "\n".join(lines) + "\n"


def parse_args(argv: Iterable[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path, help="Path to the thesis .docx file.")
    parser.add_argument("--template-dotx", type=Path, default=default_template_path(), help="DOTX template path for template-first checks.")
    parser.add_argument("--template-first", action="store_true", default=True, help="Use DOTX-derived page/style checks as primary expectations.")
    parser.add_argument("--out", type=Path, help="Write Markdown audit report to this path. Defaults to stdout.")
    parser.add_argument("--json-out", type=Path, help="Write JSON audit report to this path.")
    return parser.parse_args(list(argv))


def main(argv: Iterable[str] = sys.argv[1:]) -> int:
    args = parse_args(argv)
    report = DocxAudit(args.docx, template_dotx=args.template_dotx, template_first=args.template_first).run()
    markdown = render_markdown(report)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(markdown, encoding="utf-8")
    else:
        sys.stdout.write(markdown)
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return 1 if report.get("issue_counts", {}).get("blocking", 0) else 0


if __name__ == "__main__":
    raise SystemExit(main())
