#!/usr/bin/env python3
"""Read-only audit for SWPU undergraduate thesis DOCX formatting."""

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


def attr(el: ET.Element | None, name: str, prefix: str = "w") -> str | None:
    if el is None:
        return None
    return el.get(qn(prefix, name))


def twips_to_cm(value: str | None) -> float | None:
    if not value:
        return None
    try:
        return round(int(value) / 1440 * 2.54, 2)
    except ValueError:
        return None


def half_points(value: str | None) -> float | None:
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
        elif node.tag == qn("w", "instrText") and node.text:
            chunks.append(node.text)
        elif node.tag == qn("w", "tab"):
            chunks.append("\t")
        elif node.tag in {qn("w", "br"), qn("w", "cr")}:
            chunks.append("\n")
    return "".join(chunks)


def visible_text_of(el: ET.Element) -> str:
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
    first_line_twips: int | None
    line_twips: int | None


class SwpuAudit:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.issues: list[Issue] = []
        self.paragraphs: list[Paragraph] = []
        self.sections: list[dict] = []
        self.headers: list[dict] = []
        self.footers: list[dict] = []
        self.style_ids: set[str] = set()
        self.style_names: dict[str, str] = {}
        self.style_sizes: dict[str, float] = {}
        self.field_counts: dict[str, int] = {}
        self.table_count = 0

    def add(self, severity: str, code: str, message: str, location: str = "") -> None:
        self.issues.append(Issue(severity, code, message, location))

    def read_xml(self, zf: zipfile.ZipFile, name: str) -> ET.Element | None:
        try:
            return ET.fromstring(zf.read(name))
        except KeyError:
            return None

    def run(self) -> dict:
        if self.path.suffix.lower() != ".docx":
            self.add("blocking", "input-not-docx", "Audit input must be a .docx file.")
            return self.report()
        if not zipfile.is_zipfile(self.path):
            self.add("blocking", "invalid-docx", "File is not a valid DOCX package.")
            return self.report()

        with zipfile.ZipFile(self.path) as zf:
            document = self.read_xml(zf, "word/document.xml")
            if document is None:
                self.add("blocking", "missing-document-xml", "word/document.xml is missing.")
                return self.report()
            self.parse_styles(zf)
            self.parse_document(zf, document)
            self.parse_headers_footers(zf, document)

        self.check_page_setup()
        self.check_required_structure()
        self.check_headers_footers()
        self.check_styles()
        self.check_captions_formulas()
        self.check_references()
        self.check_placeholders()
        return self.report()

    def parse_styles(self, zf: zipfile.ZipFile) -> None:
        styles = self.read_xml(zf, "word/styles.xml")
        if styles is None:
            self.add("warning", "missing-styles", "word/styles.xml is missing; style checks will be limited.")
            return
        for style in styles.findall("w:style", NS):
            sid = attr(style, "styleId")
            if not sid:
                continue
            self.style_ids.add(sid)
            name = attr(style.find("w:name", NS), "val")
            if name:
                self.style_names[sid] = name
            rpr = style.find("w:rPr", NS)
            size = half_points(attr(rpr.find("w:sz", NS), "val")) if rpr is not None else None
            if size:
                self.style_sizes[sid] = size

    def parse_document(self, zf: zipfile.ZipFile, document: ET.Element) -> None:
        document_xml = ET.tostring(document, encoding="unicode")
        self.field_counts = {
            "TOC": document_xml.count("TOC"),
            "PAGE": document_xml.count("PAGE"),
            "NUMPAGES": document_xml.count("NUMPAGES"),
        }
        self.table_count = len(document.findall(".//w:tbl", NS))

        for p in document.findall(".//w:p", NS):
            text = norm_space(visible_text_of(p))
            if not text:
                continue
            ppr = p.find("w:pPr", NS)
            style = attr(ppr.find("w:pStyle", NS), "val") if ppr is not None else None
            align = attr(ppr.find("w:jc", NS), "val") if ppr is not None else None
            ind = ppr.find("w:ind", NS) if ppr is not None else None
            spacing = ppr.find("w:spacing", NS) if ppr is not None else None
            sizes: list[float] = []
            fonts: list[str] = []
            bold = False
            for r in p.findall("w:r", NS):
                rpr = r.find("w:rPr", NS)
                if rpr is None:
                    continue
                size = half_points(attr(rpr.find("w:sz", NS), "val"))
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
            self.paragraphs.append(
                Paragraph(
                    text=text,
                    style=style,
                    align=align,
                    sizes=sorted(set(sizes)),
                    fonts=fonts,
                    bold=bold,
                    first_line_twips=int(attr(ind, "firstLine") or 0) if ind is not None and attr(ind, "firstLine") else None,
                    line_twips=int(attr(spacing, "line") or 0) if spacing is not None and attr(spacing, "line") else None,
                )
            )

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

    def parse_headers_footers(self, zf: zipfile.ZipFile, document: ET.Element) -> None:
        rels = self.read_xml(zf, "word/_rels/document.xml.rels")
        if rels is None:
            return
        rel_map: dict[str, str] = {}
        for rel in rels.findall("rel:Relationship", NS):
            rid = rel.get("Id")
            target = rel.get("Target")
            if rid and target:
                rel_map[rid] = target.lstrip("/")

        for section_index, sect in enumerate(document.findall(".//w:sectPr", NS), start=1):
            for ref_tag, bucket in (("headerReference", self.headers), ("footerReference", self.footers)):
                for ref in sect.findall(f"w:{ref_tag}", NS):
                    rid = attr(ref, "id", "r")
                    target = rel_map.get(rid or "")
                    if not target:
                        continue
                    part = target if target.startswith("word/") else f"word/{target}"
                    xml = self.read_xml(zf, part)
                    text = norm_space(visible_text_of(xml)) if xml is not None else ""
                    bucket.append(
                        {
                            "section": section_index,
                            "type": attr(ref, "type") or "default",
                            "target": part,
                            "text": text,
                        }
                    )

    def check_page_setup(self) -> None:
        if not self.sections:
            self.add("warning", "no-sections", "No section properties found; page setup cannot be verified.")
            return
        if len(self.sections) != 4:
            self.add("info", "section-count", f"Template uses 4 sections; this document has {len(self.sections)}.")
        for s in self.sections:
            loc = f"section {s['index']}"
            if abs(s["width_twips"] - 11906) > 120 or abs(s["height_twips"] - 16838) > 120:
                self.add("blocking", "page-size", f"Expected A4 portrait page size near 11906 x 16838 twips, found {s['width_twips']} x {s['height_twips']}.", loc)
            expected = {"top_cm": 2.5, "bottom_cm": 2.5, "left_cm": 3.0, "right_cm": 2.5, "header_cm": 1.5, "footer_cm": 1.5}
            for key, target in expected.items():
                value = s.get(key)
                if value is None:
                    continue
                tolerance = 0.18 if key in {"top_cm", "bottom_cm", "left_cm", "right_cm"} else 0.12
                if abs(value - target) > tolerance:
                    self.add("warning", "page-geometry", f"{key.replace('_cm', '')} is {value} cm; template expects about {target} cm.", loc)

    def check_required_structure(self) -> None:
        all_text = "\n".join(p.text for p in self.paragraphs)
        required = [
            ("Chinese cover title", r"本科毕业设计（论文）"),
            ("English cover university", r"Southwest Petroleum University"),
            ("English cover thesis label", r"Graduation Thesis"),
            ("Chinese abstract", r"摘\s*要"),
            ("English abstract", r"\bAbstract\b"),
            ("Contents", r"目\s*录"),
            ("Introduction chapter", r"1\s+绪论"),
            ("Conclusion chapter", r"7\s+结论与展望"),
            ("Acknowledgements", r"谢\s*辞"),
            ("References", r"参考文献"),
            ("Appendix", r"附\s*录"),
        ]
        for label, pattern in required:
            if not re.search(pattern, all_text, flags=re.IGNORECASE):
                self.add("warning", "missing-required-structure", f"Expected section or title not found: {label}.")
        if self.field_counts.get("TOC", 0) == 0:
            self.add("warning", "toc-field", "No TOC field was detected; Contents may not update automatically.")

    def check_headers_footers(self) -> None:
        header_text = " ".join(h["text"] for h in self.headers)
        if "西南石油大学本科毕业设计（论文）" not in header_text:
            self.add("warning", "even-header", "正文 even-page header should contain `西南石油大学本科毕业设计（论文）`.")
        nonempty_headers = [h for h in self.headers if h["text"]]
        title_headers = [h for h in nonempty_headers if h["text"] != "西南石油大学本科毕业设计（论文）"]
        if not title_headers:
            self.add("warning", "odd-header", "正文 odd-page header should contain the thesis title.")
        elif any(h["text"] in {"你的题目", "题目"} for h in title_headers):
            self.add("info", "placeholder-header", "Odd-page header still contains a template placeholder such as `你的题目`; replace it with the final thesis title.")
        if self.field_counts.get("PAGE", 0) == 0:
            self.add("warning", "page-fields", "No PAGE fields detected; page numbering may be manual or missing.")

    def check_styles(self) -> None:
        expected_styles = {
            "2": "正文样式",
            "aa": "一级标题样式",
            "ac": "一级标题居中样式",
            "ad": "摘要样式",
            "abstract": "abstract样式",
            "ae": "二级标题样式",
            "af": "三级标题样式",
            "TOC1": "toc 1",
            "TOC2": "toc 2",
            "TOC3": "toc 3",
        }
        for sid, name in expected_styles.items():
            if sid not in self.style_ids:
                self.add("warning", "missing-style", f"Expected template style `{sid}` ({name}) was not detected.")

        body_size = self.style_sizes.get("2")
        if body_size and abs(body_size - 12.0) > 0.1:
            self.add("warning", "body-style-size", f"正文样式 should be small 4 / 12 pt; found {body_size} pt.")
        for sid, expected_size in {"aa": 16.0, "ac": 16.0, "ad": 16.0, "abstract": 16.0, "ae": 15.0, "af": 14.0}.items():
            size = self.style_sizes.get(sid)
            if size and abs(size - expected_size) > 0.1:
                self.add("warning", "heading-style-size", f"Style `{sid}` should be {expected_size} pt; found {size} pt.")

        body_paras = [
            p for p in self.paragraphs
            if p.style == "2"
            and p.align not in {"center", "right"}
            and not re.match(r"^[表图]\s*\d+\.\d+|^（\d+\.\d+）$|^\(\d+\.\d+\)$", p.text)
        ]
        if body_paras:
            non_22 = [p.text[:40] for p in body_paras if p.line_twips and p.line_twips != 440]
            if non_22:
                self.add("warning", "body-line-spacing", "Some 正文样式 paragraphs do not use exact 22 pt line spacing.", non_22[0])
        else:
            self.add("info", "body-style-usage", "No paragraphs using template body style `2` were detected.")

    def check_captions_formulas(self) -> None:
        table_caps = [p for p in self.paragraphs if re.match(r"^表\s*\d+\.\d+", p.text)]
        figure_caps = [p for p in self.paragraphs if re.match(r"^图\s*\d+\.\d+", p.text)]
        formula_nums = [p for p in self.paragraphs if re.fullmatch(r"（\d+\.\d+）|\(\d+\.\d+\)", p.text)]
        for p in table_caps:
            if p.align != "center":
                self.add("warning", "table-caption-align", "Table captions should be centered above tables.", p.text[:80])
        for p in figure_caps:
            if p.align != "center":
                self.add("warning", "figure-caption-align", "Figure captions should be centered below figures.", p.text[:80])
        for p in formula_nums:
            if p.align != "right":
                self.add("warning", "formula-number-align", "Formula numbers should be right aligned.", p.text)
        self.add("info", "detected-objects", f"Detected {self.table_count} tables, {len(table_caps)} table captions, {len(figure_caps)} figure captions, and {len(formula_nums)} formula-number paragraphs.")

    def check_references(self) -> None:
        entries = [p.text for p in self.paragraphs if re.match(r"^\[\d+\]", p.text)]
        if entries and len(entries) < 20:
            self.add("info", "reference-count", f"Detected {len(entries)} numbered reference examples/entries; final undergraduate thesis should have more than 20 references including foreign-language references.")
        if not entries:
            self.add("warning", "reference-format", "No numbered references like `[1]` were detected.")

    def check_placeholders(self) -> None:
        placeholder_patterns = [
            r"你的题目",
            r"你的姓名",
            r"选择一项",
            r"关键词1",
            r"Keywords1",
            r"XXXXXXXXXX",
            r"第2章标题",
        ]
        all_text = "\n".join(p.text for p in self.paragraphs)
        hits = [pat for pat in placeholder_patterns if re.search(pat, all_text)]
        if hits:
            self.add("info", "template-placeholders", "Template placeholder/instruction text remains; replace or remove it in a final thesis: " + ", ".join(hits))

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
            "field_counts": self.field_counts,
            "table_count": self.table_count,
            "issue_counts": counts,
            "issues": [asdict(issue) for issue in self.issues],
        }


def render_markdown(report: dict) -> str:
    lines = [
        "# SWPU Thesis DOCX Audit",
        "",
        f"- File: `{report['file']}`",
        f"- Paragraphs: {report['paragraphs']} ({report['nonempty_paragraphs']} nonempty)",
        f"- Sections: {len(report['sections'])}",
        f"- Tables: {report['table_count']}",
        f"- Field counts: TOC={report['field_counts'].get('TOC', 0)}, PAGE={report['field_counts'].get('PAGE', 0)}",
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

    lines.extend(["", "## Headers", ""])
    if report["headers"]:
        for h in report["headers"]:
            lines.append(f"- Section {h['section']} {h['type']} `{h['target']}`: `{h['text']}`")
    else:
        lines.append("No header references detected.")

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
    parser = argparse.ArgumentParser(description="Audit a DOCX against the SWPU undergraduate thesis template.")
    parser.add_argument("docx", type=Path, help="Path to a .docx thesis file.")
    parser.add_argument("--out", type=Path, help="Write report to this path.")
    parser.add_argument("--format", choices=("markdown", "json"), help="Report format. Defaults to JSON for .json outputs, otherwise Markdown.")
    parser.add_argument("--fail-on-blocking", action="store_true", help="Exit with status 1 when blocking issues are found.")
    return parser.parse_args(list(argv))


def main(argv: Iterable[str] = sys.argv[1:]) -> int:
    args = parse_args(argv)
    fmt = args.format or ("json" if args.out and args.out.suffix.lower() == ".json" else "markdown")
    report = SwpuAudit(args.docx).run()
    write_output(report, args.out, fmt)
    if args.fail_on_blocking and report["issue_counts"].get("blocking", 0):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
