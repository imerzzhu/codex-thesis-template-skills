#!/usr/bin/env python3
"""Read-only audit for Chengdu University defense PPTX template consistency."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

EXPECTED_SIZE = {"cx": 9144000, "cy": 6858000}
PRIMARY_RED = "9F2925"


def q(ns: str, tag: str) -> str:
    return "{" + NS[ns] + "}" + tag


def read_xml(pptx: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(pptx.read(name))
    except KeyError:
        return None


def text_of(root: ET.Element | None) -> str:
    if root is None:
        return ""
    return "".join(t.text or "" for t in root.iter(q("a", "t"))).strip()


def sorted_numbered_parts(names: list[str], pattern: str) -> list[str]:
    rx = re.compile(pattern)
    return sorted([name for name in names if rx.match(name)], key=lambda name: int(re.search(r"(\d+)", name).group(1)))


@dataclass
class Findings:
    issues: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    def issue(self, msg: str) -> None:
        self.issues.append(msg)

    def warning(self, msg: str) -> None:
        self.warnings.append(msg)

    def note(self, msg: str) -> None:
        self.notes.append(msg)


def collect_package_facts(path: Path) -> tuple[Findings, dict[str, Any]]:
    findings = Findings()
    facts: dict[str, Any] = {
        "path": str(path),
        "slide_count": 0,
        "layout_count": 0,
        "master_count": 0,
        "media_count": 0,
        "media_ext_counts": {},
        "slide_size": {},
        "explicit_colors": {},
        "fonts": {},
        "slide_text": [],
    }

    if not path.exists():
        findings.issue(f"File does not exist: {path}")
        return findings, facts
    if path.suffix.lower() != ".pptx":
        findings.issue("Input file must be a .pptx file.")
        return findings, facts

    try:
        with zipfile.ZipFile(path) as deck:
            names = deck.namelist()
            slides = sorted_numbered_parts(names, r"ppt/slides/slide\d+\.xml$")
            layouts = sorted_numbered_parts(names, r"ppt/slideLayouts/slideLayout\d+\.xml$")
            masters = sorted_numbered_parts(names, r"ppt/slideMasters/slideMaster\d+\.xml$")
            media = [name for name in names if name.startswith("ppt/media/") and not name.endswith("/")]

            facts["slide_count"] = len(slides)
            facts["layout_count"] = len(layouts)
            facts["master_count"] = len(masters)
            facts["media_count"] = len(media)
            facts["media_ext_counts"] = dict(Counter(Path(name).suffix.lower() for name in media))

            pres = read_xml(deck, "ppt/presentation.xml")
            if pres is None:
                findings.issue("Missing ppt/presentation.xml.")
            else:
                sld_size = pres.find(q("p", "sldSz"))
                if sld_size is None:
                    findings.issue("Missing slide size in presentation.xml.")
                else:
                    size = {
                        "cx": int(sld_size.get("cx", "0")),
                        "cy": int(sld_size.get("cy", "0")),
                        "type": sld_size.get("type", ""),
                    }
                    facts["slide_size"] = size
                    if size["cx"] != EXPECTED_SIZE["cx"] or size["cy"] != EXPECTED_SIZE["cy"]:
                        findings.issue(
                            f"Slide size is {size['cx']} x {size['cy']} EMU, expected 4:3 {EXPECTED_SIZE['cx']} x {EXPECTED_SIZE['cy']}."
                        )

            if len(masters) < 1:
                findings.issue("No slide master found.")
            if len(layouts) < 3:
                findings.warning(f"Only {len(layouts)} slide layout(s) found; the source template has 3.")
            if len(slides) < 1:
                findings.issue("No slides found.")
            if len(media) < 1:
                findings.warning("No media assets found; the CDU emblem and campus image assets may be missing.")

            empty_media = [name for name in media if len(deck.read(name)) == 0]
            if empty_media:
                findings.issue(f"Empty media files found: {', '.join(empty_media[:10])}")

            fonts: Counter[str] = Counter()
            colors: Counter[str] = Counter()
            slide_text: list[dict[str, str]] = []
            for part in slides + layouts + masters:
                root = read_xml(deck, part)
                if root is None:
                    continue
                for latin in root.iter(q("a", "latin")):
                    typeface = latin.get("typeface")
                    if typeface:
                        fonts[typeface] += 1
                for east_asia in root.iter(q("a", "ea")):
                    typeface = east_asia.get("typeface")
                    if typeface:
                        fonts[typeface] += 1
                for color in root.iter(q("a", "srgbClr")):
                    value = color.get("val")
                    if value:
                        colors[value.upper()] += 1
                if part.startswith("ppt/slides/"):
                    index = int(re.search(r"slide(\d+)\.xml", part).group(1))
                    slide_text.append({"slide": index, "text": re.sub(r"\s+", " ", text_of(root))})

            facts["fonts"] = dict(fonts.most_common(20))
            facts["explicit_colors"] = dict(colors.most_common(20))
            facts["slide_text"] = slide_text

            if colors.get(PRIMARY_RED, 0) < 10:
                findings.issue(
                    f"Primary CDU red #{PRIMARY_RED} appears only {colors.get(PRIMARY_RED, 0)} time(s); template-following decks should preserve it prominently."
                )
            if not any(font in fonts for font in ["黑体", "SimHei"]):
                findings.warning("Primary Chinese font 黑体/SimHei was not detected.")
            if facts["slide_count"] == 8:
                expected_markers = ["目录", "Contents", "谢谢观看"]
                all_text = " ".join(item["text"] for item in slide_text)
                for marker in expected_markers:
                    if marker not in all_text:
                        findings.warning(f"Expected template marker `{marker}` was not found in the 8-slide source-style deck.")
            elif facts["slide_count"] < 8:
                findings.warning(f"Deck has {facts['slide_count']} slide(s); the source template has 8 page families.")

    except zipfile.BadZipFile:
        findings.issue("Input is not a valid PPTX/ZIP package.")

    return findings, facts


def markdown_report(path: Path, findings: Findings, facts: dict[str, Any]) -> str:
    lines = [
        "# CDU Defense PPTX Audit",
        "",
        f"- File: `{path}`",
        f"- Slides: {facts.get('slide_count', 0)}",
        f"- Masters: {facts.get('master_count', 0)}",
        f"- Layouts: {facts.get('layout_count', 0)}",
        f"- Media: {facts.get('media_count', 0)} {facts.get('media_ext_counts', {})}",
        f"- Slide size: {facts.get('slide_size', {})}",
        f"- Issues: {len(findings.issues)}",
        f"- Warnings: {len(findings.warnings)}",
        "",
        "## Issues",
        "",
    ]
    lines.extend([f"- {issue}" for issue in findings.issues] or ["No blocking issues found."])
    lines.extend(["", "## Warnings", ""])
    lines.extend([f"- {warning}" for warning in findings.warnings] or ["No warnings."])
    lines.extend(["", "## Top Fonts", ""])
    lines.extend([f"- `{font}`: {count}" for font, count in facts.get("fonts", {}).items()] or ["No fonts detected."])
    lines.extend(["", "## Top Explicit Colors", ""])
    lines.extend([f"- `#{color}`: {count}" for color, count in facts.get("explicit_colors", {}).items()] or ["No explicit RGB colors detected."])
    lines.extend(["", "## Slide Text Preview", ""])
    for item in facts.get("slide_text", []):
        preview = item["text"][:220]
        lines.append(f"- Slide {item['slide']}: {preview}")
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Audit a PPTX against the CDU defense PPT template.")
    parser.add_argument("pptx", type=Path, help="PPTX file to audit")
    parser.add_argument("--out", type=Path, help="Write Markdown report to this path")
    parser.add_argument("--json", dest="json_out", type=Path, help="Write JSON report to this path")
    parser.add_argument("--fail-on-issues", action="store_true", help="Exit 1 when blocking issues are found")
    args = parser.parse_args(argv)

    findings, facts = collect_package_facts(args.pptx)
    report = markdown_report(args.pptx, findings, facts)
    print(report)

    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(report, encoding="utf-8")
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(
            json.dumps({"facts": facts, "issues": findings.issues, "warnings": findings.warnings, "notes": findings.notes}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    if args.fail_on_issues and findings.issues:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
