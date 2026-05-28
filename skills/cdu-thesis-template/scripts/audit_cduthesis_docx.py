#!/usr/bin/env python3
"""Read-only audit for Chengdu University science thesis DOCX formatting."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"


def qn(name: str) -> str:
    return W + name


def read_xml(docx: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(docx.read(name))
    except KeyError:
        return None


def w_attr(el: ET.Element | None, name: str) -> str | None:
    return el.get(qn(name)) if el is not None else None


def text_of(el: ET.Element | None) -> str:
    if el is None:
        return ""
    return "".join(t.text or "" for t in el.iter(qn("t"))).strip()


def int_attr(el: ET.Element | None, name: str) -> int | None:
    value = w_attr(el, name)
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def half_points(el: ET.Element | None) -> float | None:
    value = int_attr(el, "val")
    return value / 2 if value is not None else None


def style_id(paragraph: ET.Element) -> str | None:
    ppr = paragraph.find(qn("pPr"))
    if ppr is None:
        return None
    return w_attr(ppr.find(qn("pStyle")), "val")


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


@dataclass
class FindingSet:
    issues: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    def issue(self, message: str) -> None:
        self.issues.append(message)

    def warning(self, message: str) -> None:
        self.warnings.append(message)

    def note(self, message: str) -> None:
        self.notes.append(message)


def near(value: int | None, expected: int, tolerance: int = 4) -> bool:
    return value is not None and abs(value - expected) <= tolerance


def describe_twips(value: int | None) -> str:
    if value is None:
        return "missing"
    return f"{value} twips ({value / 567:.2f} cm)"


def collect_styles(styles_root: ET.Element | None) -> dict[str, dict[str, Any]]:
    styles: dict[str, dict[str, Any]] = {}
    if styles_root is None:
        return styles
    for style in styles_root.findall(qn("style")):
        sid = w_attr(style, "styleId")
        if not sid:
            continue
        ppr = style.find(qn("pPr"))
        rpr = style.find(qn("rPr"))
        spacing = ppr.find(qn("spacing")) if ppr is not None else None
        jc = ppr.find(qn("jc")) if ppr is not None else None
        outline = ppr.find(qn("outlineLvl")) if ppr is not None else None
        fonts = rpr.find(qn("rFonts")) if rpr is not None else None
        styles[sid] = {
            "type": w_attr(style, "type"),
            "name": w_attr(style.find(qn("name")), "val"),
            "align": w_attr(jc, "val"),
            "line": int_attr(spacing, "line"),
            "after": int_attr(spacing, "after"),
            "outline": int_attr(outline, "val"),
            "size_pt": half_points(rpr.find(qn("sz")) if rpr is not None else None),
            "east_asia_font": w_attr(fonts, "eastAsia"),
            "ascii_font": w_attr(fonts, "ascii"),
        }
    return styles


def check_sections(document: ET.Element | None, findings: FindingSet) -> None:
    if document is None:
        findings.issue("Missing word/document.xml.")
        return
    sections = document.findall(".//" + qn("sectPr"))
    if not sections:
        findings.issue("No Word sections found; page setup cannot be verified.")
        return

    expected_page = {"w": 11906, "h": 16838}
    expected_common_margins = {
        "bottom": 1418,
        "left": 1418,
        "right": 1418,
        "header": 851,
        "footer": 851,
    }
    for index, section in enumerate(sections, start=1):
        pg_size = section.find(qn("pgSz"))
        for key, expected in expected_page.items():
            actual = int_attr(pg_size, key)
            if not near(actual, expected):
                findings.issue(
                    f"Section {index} page {key} is {describe_twips(actual)}, expected {expected} twips for A4."
                )

        margins = section.find(qn("pgMar"))
        top = int_attr(margins, "top")
        allowed_top = (1985, 2268) if index == 1 else (2268,)
        if top is None or all(not near(top, value) for value in allowed_top):
            expected_text = " or ".join(str(v) for v in allowed_top)
            findings.issue(
                f"Section {index} top margin is {describe_twips(top)}, expected {expected_text} twips."
            )

        for key, expected in expected_common_margins.items():
            actual = int_attr(margins, key)
            if not near(actual, expected):
                findings.issue(
                    f"Section {index} {key} margin is {describe_twips(actual)}, expected {expected} twips."
                )


def check_styles(styles: dict[str, dict[str, Any]], findings: FindingSet) -> None:
    required = {
        "1": "Heading 1 / 一级标题",
        "21": "Heading 2 / 二级标题",
        "31": "Heading 3 / 三级标题",
        "affff3": "figure/table caption / 图名",
        "affff7": "reference body / 参考文献 正文",
        "affff9": "table text / 表",
    }
    for sid, label in required.items():
        if sid not in styles:
            findings.issue(f"Missing template style `{sid}` ({label}).")

    checks = [
        ("1", "align", "center", "Heading 1 should be centered."),
        ("1", "line", 360, "Heading 1 should use 1.5 line spacing."),
        ("1", "after", 240, "Heading 1 should have 12 pt after spacing."),
        ("1", "outline", 0, "Heading 1 should use outline level 0."),
        ("21", "line", 360, "Heading 2 should use 1.5 line spacing."),
        ("21", "outline", 1, "Heading 2 should use outline level 1."),
        ("31", "line", 360, "Heading 3 should use 1.5 line spacing."),
        ("31", "outline", 2, "Heading 3 should use outline level 2."),
        ("affff3", "align", "center", "Caption style should be centered."),
        ("affff3", "line", 240, "Caption style should use single line spacing."),
        ("affff9", "align", "center", "Table text style should be centered."),
        ("affff9", "line", 240, "Table text style should use single line spacing."),
    ]
    for sid, key, expected, message in checks:
        if sid not in styles:
            continue
        actual = styles[sid].get(key)
        if actual != expected:
            findings.issue(f"{message} Found `{actual}` in style `{sid}`.")

    for sid, expected_size, label in [
        ("1", 15.0, "Heading 1"),
        ("21", 14.0, "Heading 2"),
        ("affff3", 10.5, "Caption"),
        ("affff7", 10.5, "Reference body"),
    ]:
        actual = styles.get(sid, {}).get("size_pt")
        if actual is not None and abs(actual - expected_size) > 0.1:
            findings.issue(f"{label} style `{sid}` size is {actual} pt, expected {expected_size} pt.")


def check_headers_footers(docx: zipfile.ZipFile, findings: FindingSet) -> None:
    rels = read_xml(docx, "word/_rels/document.xml.rels")
    if rels is None:
        findings.warning("Missing document relationships; headers and footers cannot be checked.")
        return

    header_texts: list[str] = []
    footer_texts: list[str] = []
    for rel in rels:
        rel_type = rel.get("Type", "").rsplit("/", 1)[-1]
        if rel_type not in {"header", "footer"}:
            continue
        target = rel.get("Target", "")
        part_name = target.lstrip("/")
        if not part_name.startswith("word/"):
            part_name = "word/" + part_name
        root = read_xml(docx, part_name)
        content = re.sub(r"\s+", " ", text_of(root))
        if rel_type == "header":
            header_texts.append(content)
        else:
            footer_texts.append(content)

    expected_header = "成都大学本科毕业设计（论文）"
    if expected_header not in header_texts:
        findings.issue("Expected running header text `成都大学本科毕业设计（论文）` was not found.")
    if not footer_texts:
        findings.warning("No footer parts found; page numbers may be missing.")
    elif not any(text for text in footer_texts):
        findings.warning("Footer parts exist but all visible footer text is empty; page numbers may rely on fields or be missing.")


def collect_paragraphs(document: ET.Element | None) -> list[tuple[str | None, str]]:
    if document is None:
        return []
    paragraphs: list[tuple[str | None, str]] = []
    for paragraph in document.findall(".//" + qn("p")):
        content = re.sub(r"\s+", " ", text_of(paragraph))
        if content:
            paragraphs.append((style_id(paragraph), content))
    return paragraphs


def check_paragraph_structure(paragraphs: list[tuple[str | None, str]], findings: FindingSet) -> None:
    heading_counts = {"1": 0, "21": 0, "31": 0}
    for sid, _ in paragraphs:
        if sid in heading_counts:
            heading_counts[sid] += 1
    if heading_counts["1"] == 0:
        findings.issue("No Heading 1 paragraphs found.")
    if heading_counts["21"] == 0:
        findings.warning("No Heading 2 paragraphs found.")
    if heading_counts["31"] == 0:
        findings.warning("No Heading 3 paragraphs found.")

    caption_styles = {"affff3", "affff8", "affffb", "afffff3"}
    caption_pattern = re.compile(r"^(图|表)\s*(?:[0-9A-Z]+(?:\.[0-9A-Z]+)?|[一二三四五六七八九十]+)\s+")
    for sid, content in paragraphs:
        if caption_pattern.match(content) and sid not in caption_styles:
            findings.issue(f"Caption-like paragraph should use a caption style, found style `{sid}`: {content[:80]}")

    formula_number = re.compile(r"^（\d+(?:\.\d+)+）$")
    for sid, content in paragraphs:
        if formula_number.match(content) and sid not in {"affffe", None}:
            findings.warning(f"Formula-number paragraph has unexpected style `{sid}`: {content}")

    if not any(content == "参考文献" and sid in {"1", "affffa", "affffd", "12"} for sid, content in paragraphs):
        findings.issue("Required `参考文献` heading was not found with a recognized heading/reference style.")

    ref_start = None
    for index, (_, content) in enumerate(paragraphs):
        if content == "参考文献":
            ref_start = index
            break
    if ref_start is not None:
        ref_entries = [content for _, content in paragraphs[ref_start + 1 :] if re.match(r"^\[\d+\]", content)]
        if len(ref_entries) < 15:
            findings.warning(f"Only {len(ref_entries)} bracket-numbered reference entries found after `参考文献`; CDU template expects at least 15.")

    if not any(content.replace(" ", "") in {"致谢", "致謝"} and sid == "1" for sid, content in paragraphs):
        findings.warning("Required `致 谢` section was not found as Heading 1.")


def check_tables(document: ET.Element | None, findings: FindingSet) -> None:
    if document is None:
        return
    tables = document.findall(".//" + qn("tbl"))
    if not tables:
        findings.warning("No tables found.")
        return
    missing_grid = 0
    for table in tables:
        if table.find(qn("tblGrid")) is None:
            missing_grid += 1
    if missing_grid:
        findings.warning(f"{missing_grid} table(s) are missing explicit tblGrid definitions.")


def audit(path: Path) -> tuple[FindingSet, dict[str, Any]]:
    findings = FindingSet()
    metadata: dict[str, Any] = {"path": str(path), "paragraph_count": 0, "table_count": 0, "section_count": 0}

    if not path.exists():
        findings.issue(f"File does not exist: {path}")
        return findings, metadata
    if path.suffix.lower() != ".docx":
        findings.issue("Input file must be a .docx document.")
        return findings, metadata

    try:
        with zipfile.ZipFile(path) as docx:
            document = read_xml(docx, "word/document.xml")
            styles_root = read_xml(docx, "word/styles.xml")
            styles = collect_styles(styles_root)
            paragraphs = collect_paragraphs(document)
            metadata["paragraph_count"] = len(paragraphs)
            metadata["section_count"] = len(document.findall(".//" + qn("sectPr"))) if document is not None else 0
            metadata["table_count"] = len(document.findall(".//" + qn("tbl"))) if document is not None else 0

            check_sections(document, findings)
            check_styles(styles, findings)
            check_headers_footers(docx, findings)
            check_paragraph_structure(paragraphs, findings)
            check_tables(document, findings)
    except zipfile.BadZipFile:
        findings.issue("Input is not a valid DOCX/ZIP package.")

    return findings, metadata


def markdown_report(path: Path, findings: FindingSet, metadata: dict[str, Any]) -> str:
    lines = [
        "# CDU Thesis DOCX Audit",
        "",
        f"- File: `{path}`",
        f"- Paragraphs: {metadata.get('paragraph_count', 0)}",
        f"- Sections: {metadata.get('section_count', 0)}",
        f"- Tables: {metadata.get('table_count', 0)}",
        f"- Issues: {len(findings.issues)}",
        f"- Warnings: {len(findings.warnings)}",
        "",
    ]

    def add_section(title: str, values: list[str], empty: str) -> None:
        lines.append(f"## {title}")
        lines.append("")
        if values:
            lines.extend(f"- {value}" for value in values)
        else:
            lines.append(empty)
        lines.append("")

    add_section("Issues", findings.issues, "No blocking issues found.")
    add_section("Warnings", findings.warnings, "No warnings.")
    add_section("Notes", findings.notes, "No additional notes.")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Audit a DOCX against CDU science thesis template rules.")
    parser.add_argument("docx", type=Path, help="DOCX file to audit")
    parser.add_argument("--out", type=Path, help="Write Markdown report to this path")
    parser.add_argument("--json", dest="json_out", type=Path, help="Write machine-readable JSON report")
    parser.add_argument("--fail-on-issues", action="store_true", help="Exit with code 1 if blocking issues are found")
    args = parser.parse_args(argv)

    findings, metadata = audit(args.docx)
    report = markdown_report(args.docx, findings, metadata)
    print(report)

    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(report, encoding="utf-8")
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(
            json.dumps(
                {
                    "metadata": metadata,
                    "issues": findings.issues,
                    "warnings": findings.warnings,
                    "notes": findings.notes,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

    if args.fail_on_issues and findings.issues:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
