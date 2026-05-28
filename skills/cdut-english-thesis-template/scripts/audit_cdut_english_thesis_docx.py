#!/usr/bin/env python3
"""Read-only audit for CDUT English-major thesis DOCX formatting."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from dataclasses import dataclass, asdict
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


@dataclass
class Issue:
    severity: str
    code: str
    message: str
    location: str = ""


@dataclass
class Paragraph:
    text: str
    style: str | None
    align: str | None
    sizes: list[float]
    fonts: list[str]
    bold: bool


class DocxAudit:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.issues: list[Issue] = []
        self.paragraphs: list[Paragraph] = []
        self.sections: list[dict] = []
        self.headers: list[str] = []
        self.footers: list[str] = []
        self.style_ids: set[str] = set()
        self.style_fonts: set[str] = set()
        self.style_sizes: dict[str, float] = {}

    def add(self, severity: str, code: str, message: str, location: str = "") -> None:
        self.issues.append(Issue(severity, code, message, location))

    def read_xml(self, zf: zipfile.ZipFile, name: str) -> ET.Element | None:
        try:
            return ET.fromstring(zf.read(name))
        except KeyError:
            return None

    def run(self) -> dict:
        if self.path.suffix.lower() != ".docx":
            self.add("blocking", "input-not-docx", "Audit input must be a .docx file. Convert .doc working copies before auditing.")
            return self.report()
        if not zipfile.is_zipfile(self.path):
            self.add("blocking", "invalid-docx", "File is not a valid DOCX zip package.")
            return self.report()

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
        self.check_required_sections()
        self.check_headers()
        self.check_styles_and_fonts()
        self.check_captions()
        self.check_works_cited()
        self.check_parenthetical_citations()
        return self.report()

    def parse_styles(self, zf: zipfile.ZipFile) -> None:
        styles = self.read_xml(zf, "word/styles.xml")
        if styles is None:
            self.add("warning", "missing-styles", "word/styles.xml is missing; style checks will be limited.")
            return
        for style in styles.findall("w:style", NS):
            style_id = attr(style, "styleId")
            if style_id:
                self.style_ids.add(style_id)
            rpr = style.find("w:rPr", NS)
            if rpr is None:
                continue
            rfonts = rpr.find("w:rFonts", NS)
            if rfonts is not None:
                for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                    value = attr(rfonts, key)
                    if value:
                        self.style_fonts.add(value)
            size = half_points_to_points(attr(rpr.find("w:sz", NS), "val"))
            if style_id and size:
                self.style_sizes[style_id] = size

    def parse_paragraphs(self, document: ET.Element) -> None:
        for p in document.findall(".//w:p", NS):
            text = norm_space(text_of(p))
            ppr = p.find("w:pPr", NS)
            style = attr(ppr.find("w:pStyle", NS), "val") if ppr is not None else None
            align = attr(ppr.find("w:jc", NS), "val") if ppr is not None else None
            sizes: list[float] = []
            fonts: list[str] = []
            bold = False
            for r in p.findall("w:r", NS):
                rpr = r.find("w:rPr", NS)
                if rpr is None:
                    continue
                size = half_points_to_points(attr(rpr.find("w:sz", NS), "val"))
                if size:
                    sizes.append(size)
                rfonts = rpr.find("w:rFonts", NS)
                if rfonts is not None:
                    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                        value = attr(rfonts, key)
                        if value and value not in fonts:
                            fonts.append(value)
                if rpr.find("w:b", NS) is not None:
                    bold = True
            self.paragraphs.append(Paragraph(text, style, align, sizes, fonts, bold))

    def parse_sections(self, document: ET.Element) -> None:
        for index, sect in enumerate(document.findall(".//w:sectPr", NS), start=1):
            pg_sz = sect.find("w:pgSz", NS)
            pg_mar = sect.find("w:pgMar", NS)
            self.sections.append(
                {
                    "index": index,
                    "width_twips": int(attr(pg_sz, "w") or 0) if pg_sz is not None else 0,
                    "height_twips": int(attr(pg_sz, "h") or 0) if pg_sz is not None else 0,
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

    def check_page_setup(self) -> None:
        if not self.sections:
            self.add("warning", "no-sections", "No section properties found; page setup cannot be verified.")
            return
        for section in self.sections:
            loc = f"section {section['index']}"
            width = section["width_twips"]
            height = section["height_twips"]
            if abs(width - 11906) > 120 or abs(height - 16838) > 120:
                self.add("blocking", "page-size", f"Expected A4 portrait page size near 11906 x 16838 twips, found {width} x {height}.", loc)
            for side in ("top_cm", "bottom_cm", "left_cm", "right_cm"):
                value = section.get(side)
                if value is not None and not (2.2 <= value <= 3.3):
                    self.add("warning", "margin-range", f"{side.replace('_cm', '')} margin {value} cm is outside the template's expected 2.4-3.17 cm range.", loc)
            header = section.get("header_cm")
            if header is not None and not (1.3 <= header <= 1.8):
                self.add("warning", "header-distance", f"Header distance {header} cm differs from the template's about 1.5 cm setting.", loc)

    def check_required_sections(self) -> None:
        all_text = "\n".join(p.text for p in self.paragraphs if p.text)
        required = [
            ("Abstract", r"\bAbstract\b"),
            ("Chinese abstract heading", r"摘\s*要"),
            ("Contents", r"\bContents\b"),
            ("Introduction", r"\bIntroduction\b"),
            ("Conclusion", r"\bConclusion\b"),
            ("Acknowledgements", r"\bAcknowledgements?\b"),
            ("Works Cited", r"\bWorks\s+Cited\b"),
        ]
        for label, pattern in required:
            if not re.search(pattern, all_text, flags=re.IGNORECASE):
                self.add("warning", "missing-required-section", f"Expected section title not found: {label}.")
        if re.search(r"\bWork\s+Cited\b", all_text, flags=re.IGNORECASE) and not re.search(r"\bWorks\s+Cited\b", all_text, flags=re.IGNORECASE):
            self.add("blocking", "works-cited-plural", "Use the exact plural heading `Works Cited`, not `Work Cited`.")

    def check_headers(self) -> None:
        header_text = " ".join(self.headers)
        if header_text and ("成都理工大学" not in header_text or "学士学位论文" not in header_text):
            self.add("warning", "header-text", "Running header should contain `成都理工大学20xx届学士学位论文`.")
        elif not header_text:
            self.add("warning", "missing-header", "No header text found. The template starts the running header from the English abstract page.")

    def check_styles_and_fonts(self) -> None:
        fonts = set(self.style_fonts)
        for p in self.paragraphs:
            fonts.update(p.fonts)
        font_blob = " ".join(sorted(fonts)).lower()
        if "times new roman" not in font_blob:
            self.add("warning", "font-times-new-roman", "Times New Roman was not detected in styles or direct formatting.")
        if not any(font_marker in font_blob for font_marker in ("simsun", "song", "宋")):
            self.add("warning", "font-chinese-songti", "Songti/SimSun-style Chinese font was not detected in styles or direct formatting.")
        for style in ("Heading1", "Heading2", "Heading3", "Heading4"):
            if style not in self.style_ids:
                self.add("info", "heading-style", f"{style} style was not detected; TOC/navigation behavior may need manual checking.")

    def check_captions(self) -> None:
        table_caps = [p for p in self.paragraphs if re.match(r"^表\s*\d+\s*-\s*\d+", p.text)]
        figure_caps = [p for p in self.paragraphs if re.match(r"^图\s*\d+\s*-\s*\d+|^附图", p.text)]
        for caption_type, caps in (("table", table_caps), ("figure", figure_caps)):
            for p in caps:
                if p.align not in ("center", "both"):
                    self.add("warning", f"{caption_type}-caption-align", f"{caption_type.capitalize()} caption should be centered: `{p.text[:80]}`.")
        if table_caps or figure_caps:
            self.add("info", "captions-detected", f"Detected {len(table_caps)} table captions and {len(figure_caps)} figure captions.")

    def check_works_cited(self) -> None:
        index = next((i for i, p in enumerate(self.paragraphs) if re.fullmatch(r"Works\s+Cited", p.text, flags=re.IGNORECASE)), None)
        if index is None:
            return
        heading = self.paragraphs[index]
        if heading.align not in (None, "left", "start"):
            self.add("warning", "works-cited-heading-align", "`Works Cited` should be left aligned.", "Works Cited")
        entries = [
            p.text
            for p in self.paragraphs[index + 1 :]
            if p.text and not re.match(r"^(补充|更新|Additional|Appendix)\b", p.text, flags=re.IGNORECASE)
        ]
        if not entries:
            self.add("warning", "works-cited-empty", "`Works Cited` heading was found but no reference entries were detected.")
            return
        first_cjk = next((i for i, e in enumerate(entries) if re.search(r"[\u4e00-\u9fff]", e)), None)
        later_latin = False
        if first_cjk is not None:
            later_latin = any(re.match(r"^[A-Z][A-Za-z' -]+[,\.]", e) for e in entries[first_cjk + 1 :])
        if later_latin:
            self.add("warning", "works-cited-order", "English references should appear before Chinese references.")

    def check_parenthetical_citations(self) -> None:
        all_text = " ".join(p.text for p in self.paragraphs)
        english = re.findall(r"\([A-Z][A-Za-z' -]{1,40}\s+\d+(?:-\d+)?\)", all_text)
        chinese = re.findall(r"\([\u4e00-\u9fff]{1,12}(?:、[\u4e00-\u9fff]{1,12}|等)?\s+\d+(?:-\d+)?\)", all_text)
        if english or chinese:
            self.add("info", "parenthetical-citations", f"Detected {len(english)} English MLA-style and {len(chinese)} Chinese parenthetical citation examples.")

    def report(self) -> dict:
        counts = {"blocking": 0, "warning": 0, "info": 0}
        for issue in self.issues:
            counts[issue.severity] = counts.get(issue.severity, 0) + 1
        return {
            "file": str(self.path),
            "paragraphs": len(self.paragraphs),
            "nonempty_paragraphs": sum(1 for p in self.paragraphs if p.text),
            "sections": self.sections,
            "headers": self.headers,
            "footers": self.footers,
            "issue_counts": counts,
            "issues": [asdict(issue) for issue in self.issues],
        }


def render_markdown(report: dict) -> str:
    lines = [
        "# CDUT English Thesis DOCX Audit",
        "",
        f"- File: `{report['file']}`",
        f"- Paragraphs: {report['paragraphs']} ({report['nonempty_paragraphs']} nonempty)",
        f"- Sections: {len(report['sections'])}",
        f"- Blocking issues: {report['issue_counts'].get('blocking', 0)}",
        f"- Warnings: {report['issue_counts'].get('warning', 0)}",
        f"- Info: {report['issue_counts'].get('info', 0)}",
        "",
        "## Sections",
        "",
    ]
    if report["sections"]:
        lines.append("| # | Size twips | Margins cm (T/B/L/R) | Header/Footer cm |")
        lines.append("|---|---:|---|---|")
        for s in report["sections"]:
            margins = f"{s['top_cm']}/{s['bottom_cm']}/{s['left_cm']}/{s['right_cm']}"
            header_footer = f"{s['header_cm']}/{s['footer_cm']}"
            lines.append(f"| {s['index']} | {s['width_twips']} x {s['height_twips']} | {margins} | {header_footer} |")
    else:
        lines.append("No section properties detected.")
    lines.extend(["", "## Issues", ""])
    if report["issues"]:
        for issue in report["issues"]:
            loc = f" ({issue['location']})" if issue.get("location") else ""
            lines.append(f"- **{issue['severity']}** `{issue['code']}`{loc}: {issue['message']}")
    else:
        lines.append("No issues detected.")
    lines.append("")
    return "\n".join(lines)


def write_output(report: dict, out: Path | None, fmt: str) -> None:
    if fmt == "json":
        content = json.dumps(report, ensure_ascii=False, indent=2)
    else:
        content = render_markdown(report)
    if out:
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(content, encoding="utf-8")
    else:
        print(content)


def parse_args(argv: Iterable[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Audit a DOCX against CDUT English thesis template expectations.")
    parser.add_argument("docx", type=Path, help="Path to a .docx thesis file.")
    parser.add_argument("--out", type=Path, help="Write report to this path.")
    parser.add_argument("--format", choices=("markdown", "json"), help="Report format. Defaults to JSON for .json outputs, otherwise Markdown.")
    parser.add_argument("--fail-on-blocking", action="store_true", help="Exit with status 1 when blocking issues are found.")
    return parser.parse_args(list(argv))


def main(argv: Iterable[str] = sys.argv[1:]) -> int:
    args = parse_args(argv)
    fmt = args.format or ("json" if args.out and args.out.suffix.lower() == ".json" else "markdown")
    report = DocxAudit(args.docx).run()
    write_output(report, args.out, fmt)
    if args.fail_on_blocking and report["issue_counts"].get("blocking", 0):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
