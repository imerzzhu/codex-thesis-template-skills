#!/usr/bin/env python3
"""Repair static TOC page-number runs using Microsoft Word actual pages.

The script writes a new DOCX. It changes only the final numeric w:t run in each
static TOC row and then re-audits the result with Microsoft Word actual page
numbers. It does not update live Word TOC fields.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from audit_word_toc_pages import NS, audit_docx, parse_toc_entry_text, qn, render_markdown, sha256_file, text_of


PRESERVE_PREFIXES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
}

for _prefix, _uri in PRESERVE_PREFIXES.items():
    ET.register_namespace(_prefix, _uri)


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_Word实际页码修复{input_path.suffix}")


def zip_part_hashes(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as zf:
        return {info.filename: hashlib.sha256(zf.read(info.filename)).hexdigest() for info in zf.infolist()}


def structure_stats(path: Path) -> dict[str, int]:
    with zipfile.ZipFile(path) as zf:
        document = ET.fromstring(zf.read("word/document.xml"))
    return {
        "paragraphs": len(document.findall(".//w:body/w:p", NS)),
        "tables": len(document.findall(".//w:tbl", NS)),
        "drawings": len(document.findall(".//w:drawing", NS)),
    }


def ensure_ignorable_namespace_declarations(xml_bytes: bytes) -> bytes:
    text = xml_bytes.decode("utf-8")
    match = re.search(r'mc:Ignorable="([^"]+)"', text)
    if not match:
        return xml_bytes
    additions: list[str] = []
    for prefix in match.group(1).split():
        uri = PRESERVE_PREFIXES.get(prefix)
        if uri and f"xmlns:{prefix}=" not in text:
            additions.append(f'xmlns:{prefix}="{uri}"')
    if additions:
        text = text.replace("<w:document ", "<w:document " + " ".join(additions) + " ", 1)
    return text.encode("utf-8")


def replace_static_toc_pages(input_path: Path, output_path: Path, entries: list[dict]) -> list[dict]:
    with zipfile.ZipFile(input_path, "r") as zin:
        document = ET.fromstring(zin.read("word/document.xml"))
        body = document.find("w:body", NS)
        if body is None:
            raise RuntimeError("word/document.xml has no w:body.")
        body_children = list(body)
        changes: list[dict] = []
        for entry in entries:
            actual_page = entry.get("actual_page")
            if not actual_page:
                raise RuntimeError(f"Cannot repair unmatched TOC entry: {entry.get('label')}")
            para_index = int(entry["paragraph_index"])
            if para_index >= len(body_children) or body_children[para_index].tag != qn("w", "p"):
                raise RuntimeError(f"TOC paragraph index is not a paragraph: {para_index}")
            para = body_children[para_index]
            text = re.sub(r"\s+", " ", text_of(para)).strip()
            parsed = parse_toc_entry_text(text)
            if not parsed:
                raise RuntimeError(f"TOC row no longer has a trailing page number: paragraph {para_index}")
            label, old_page = parsed
            if label != entry.get("label") or old_page != entry.get("toc_page"):
                raise RuntimeError(
                    "TOC row changed before repair at paragraph "
                    f"{para_index}: expected {entry.get('label')!r}/{entry.get('toc_page')!r}, "
                    f"found {label!r}/{old_page!r}"
                )
            text_runs = para.findall(".//w:t", NS)
            numeric_runs = [run for run in text_runs if re.fullmatch(r"\d+", run.text or "")]
            if not numeric_runs:
                raise RuntimeError(f"No standalone numeric w:t run in TOC row: {label}")
            target = numeric_runs[-1]
            target.text = str(actual_page)
            changes.append(
                {
                    "paragraph_index": para_index,
                    "label": label,
                    "old_page": old_page,
                    "new_page": str(actual_page),
                    "changed": old_page != str(actual_page),
                }
            )
        new_document_xml = ensure_ignorable_namespace_declarations(
            ET.tostring(document, encoding="utf-8", xml_declaration=True)
        )
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = new_document_xml if info.filename == "word/document.xml" else zin.read(info.filename)
                zout.writestr(info, data)
    return changes


def render_repair_markdown(report: dict) -> str:
    lines = [
        "# Static TOC Word actual-page repair audit",
        "",
        f"- Input: `{report['input']}`",
        f"- Output: `{report.get('output')}`",
        f"- Input SHA256: `{report.get('input_sha256')}`",
        f"- Output SHA256: `{report.get('output_sha256')}`",
        f"- Word physical pages after repair: `{report.get('after', {}).get('physical_pages')}`",
        f"- Word last adjusted page after repair: `{report.get('after', {}).get('last_adjusted_page')}`",
        f"- Final TOC match result: `{report.get('after', {}).get('matched_count')}/{report.get('after', {}).get('static_toc_entries_count')}`",
        f"- Final mismatches: `{report.get('after', {}).get('mismatch_count')}`",
        f"- Final unmatched headings: `{report.get('after', {}).get('unmatched_count')}`",
        f"- Changed ZIP parts: `{', '.join(report.get('changed_zip_parts', []))}`",
        "",
        "## Updated Page Numbers",
        "",
        "| # | TOC label | old | new |",
        "|---:|---|---:|---:|",
    ]
    for idx, change in enumerate(report.get("changes", []), 1):
        label = str(change["label"]).replace("|", "\\|")
        lines.append(f"| {idx} | {label} | {change['old_page']} | {change['new_page']} |")
    lines.extend(
        [
            "",
            "## Notes",
            "",
            "- Only static TOC page-number runs were changed.",
            "- Microsoft Word AdjustedPageNumber is the page-number authority for this report.",
            "- LibreOffice/PNG rendering remains visual QA only and is not used to decide TOC page numbers.",
        ]
    )
    return "\n".join(lines) + "\n"


def write_field_only_report(report: dict, audit_out: Path | None, json_out: Path | None) -> None:
    text = render_markdown(report)
    if audit_out:
        audit_out.write_text(text, encoding="utf-8")
    if json_out:
        json_out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out", type=Path, help="Output DOCX path. Defaults to *_Word实际页码修复.docx.")
    parser.add_argument("--audit-out", type=Path, help="Markdown repair audit path.")
    parser.add_argument("--json-out", type=Path, help="JSON repair audit path.")
    parser.add_argument("--force", action="store_true", help="Allow replacing an existing output path, but never the input path.")
    parser.add_argument("--word-timeout", type=int, default=180, help="Word COM scan timeout in seconds.")
    args = parser.parse_args(argv)

    input_path = args.docx.resolve()
    output_path = (args.out or default_output_path(input_path)).resolve()
    audit_out = args.audit_out or output_path.with_name(f"{output_path.stem}_audit.md")
    json_out = args.json_out or output_path.with_name(f"{output_path.stem}_audit.json")

    if input_path == output_path:
        print("ERROR: output path must not be the same as input path.", file=sys.stderr)
        return 1
    if output_path.exists() and not args.force:
        print(f"ERROR: output already exists: {output_path}", file=sys.stderr)
        return 1

    before = audit_docx(input_path, word_timeout=args.word_timeout)
    if not before.get("word_available"):
        write_field_only_report(before, audit_out, json_out)
        print(f"ERROR: Word actual-page audit failed: {before.get('word_error')}", file=sys.stderr)
        return 1
    if before.get("has_toc_field") and not before.get("static_toc_detected"):
        write_field_only_report(before, audit_out, json_out)
        print("ERROR: live Word TOC field detected. Update fields in Word only when explicitly requested.", file=sys.stderr)
        return 2
    if not before.get("static_toc_detected"):
        write_field_only_report(before, audit_out, json_out)
        print("ERROR: no static TOC entries detected.", file=sys.stderr)
        return 1
    if before.get("unmatched_count"):
        write_field_only_report(before, audit_out, json_out)
        print("ERROR: at least one TOC row could not be matched to a heading.", file=sys.stderr)
        return 1

    entries = before["toc_entries"]
    changes = replace_static_toc_pages(input_path, output_path, entries)
    after = audit_docx(output_path, word_timeout=args.word_timeout)
    input_hashes = zip_part_hashes(input_path)
    output_hashes = zip_part_hashes(output_path)
    changed_parts = sorted(name for name in set(input_hashes) | set(output_hashes) if input_hashes.get(name) != output_hashes.get(name))

    report = {
        "input": str(input_path),
        "output": str(output_path),
        "input_sha256": sha256_file(input_path),
        "output_sha256": sha256_file(output_path),
        "before": before,
        "after": after,
        "changes": changes,
        "changed_entries": len([change for change in changes if change["changed"]]),
        "unchanged_entries": len([change for change in changes if not change["changed"]]),
        "changed_zip_parts": changed_parts,
        "input_structure": structure_stats(input_path),
        "output_structure": structure_stats(output_path),
    }
    audit_out.write_text(render_repair_markdown(report), encoding="utf-8")
    json_out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    if not after.get("word_available"):
        print(f"ERROR: Word actual-page audit failed after repair. See {audit_out}", file=sys.stderr)
        return 2
    if after.get("mismatch_count") or after.get("unmatched_count"):
        print(f"ERROR: repair output still has TOC mismatches. See {audit_out}", file=sys.stderr)
        return 2
    print(f"Wrote {output_path}")
    print(f"Wrote {audit_out}")
    print(f"Wrote {json_out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
