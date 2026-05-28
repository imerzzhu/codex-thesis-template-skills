#!/usr/bin/env python3
"""Audit CDUT thesis TOC page numbers against Microsoft Word actual pages.

This script is intentionally read-only. It parses the DOCX package to find a
static TOC and uses Microsoft Word COM, through PowerShell, to read the actual
adjusted page number shown by Word for each target heading.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


def qn(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().upper()


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


def clean_text(text: str) -> str:
    text = re.sub(r"[\x00-\x1f\x7f]", "", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_match_text(text: str) -> str:
    text = re.sub(r"[\x00-\x1f\x7f]", "", text)
    text = text.replace("－", "-").replace("–", "-").replace("—", "-")
    text = text.replace("-", "")
    text = re.sub(r"\s+", "", text)
    return text.lower()


def toc_label_candidates(label: str) -> list[str]:
    candidates = [normalize_match_text(label)]
    match = re.match(r"^第[一二三四五六七八九十百0-9]+\s*章\s*(.+)$", label)
    if match:
        candidates.append(normalize_match_text(match.group(1)))
    return list(dict.fromkeys(candidates))


def parse_toc_entry_text(text: str) -> tuple[str, str] | None:
    text = text.strip()
    match = re.match(r"^(?P<label>.*?)(?P<page>\d+)\s*$", text)
    if not match:
        return None
    label = match.group("label")
    label = re.sub(r"[\t .·•…]+$", "", label).strip()
    page = match.group("page")
    if not label:
        return None
    return label, page


def parse_docx_package(path: Path) -> dict:
    if path.suffix.lower() != ".docx":
        raise ValueError("Input must be a .docx file.")
    if not zipfile.is_zipfile(path):
        raise ValueError("Input is not a valid DOCX zip package.")
    with zipfile.ZipFile(path) as zf:
        document = ET.fromstring(zf.read("word/document.xml"))

    has_toc_field = False
    for node in document.iter():
        if node.tag == qn("w", "instrText") and node.text and "TOC" in node.text.upper():
            has_toc_field = True
            break
        if node.tag == qn("w", "fldSimple"):
            instr = node.get(qn("w", "instr"), "")
            if "TOC" in instr.upper():
                has_toc_field = True
                break

    body = document.find("w:body", NS)
    paragraphs: list[dict] = []
    if body is not None:
        for idx, child in enumerate(list(body)):
            if child.tag == qn("w", "p"):
                paragraphs.append({"paragraph_index": idx, "text": clean_text(text_of(child))})

    toc_heading_index: int | None = None
    for para in paragraphs:
        if normalize_match_text(para["text"]) == normalize_match_text("目录"):
            toc_heading_index = int(para["paragraph_index"])
            break

    static_entries: list[dict] = []
    if toc_heading_index is not None:
        started = False
        for para in paragraphs:
            idx = int(para["paragraph_index"])
            if idx <= toc_heading_index:
                continue
            text = para["text"]
            parsed = parse_toc_entry_text(text)
            if parsed:
                label, page = parsed
                static_entries.append(
                    {
                        "paragraph_index": idx,
                        "label": label,
                        "toc_page": page,
                        "raw_text": text,
                    }
                )
                started = True
                continue
            if started and text:
                break

    return {
        "has_toc_field": has_toc_field,
        "toc_heading_index": toc_heading_index,
        "paragraph_count": len(paragraphs),
        "static_toc_entries": static_entries,
    }


WORD_SCAN_PS1 = r"""
param(
  [Parameter(Mandatory=$true)][string]$DocPath,
  [Parameter(Mandatory=$true)][string]$JsonPath
)
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($DocPath, $false, $true, $false)
  $doc.Repaginate()
  $physicalPages = $doc.ComputeStatistics(2)
  $end = [Math]::Max(0, $doc.Content.End - 1)
  $lastRange = $doc.Range($end, $end)
  $lastAdjusted = $lastRange.Information(1)
  $items = @()
  for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
    $range = $doc.Paragraphs.Item($i).Range
    $text = $range.Text
    $text = $text.Trim([char]13, [char]7, [char]11, [char]12, ' ', "`t")
    $items += [pscustomobject]@{
      paragraph_index = $i - 1
      text = $text
      adjusted_page = [string]$range.Information(1)
      physical_page = [string]$range.Information(3)
    }
  }
  $result = [ordered]@{
    word_available = $true
    physical_pages = [int]$physicalPages
    last_adjusted_page = [string]$lastAdjusted
    word_paragraph_count = [int]$doc.Paragraphs.Count
    paragraphs = $items
  }
  $result | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $JsonPath -Encoding UTF8
}
finally {
  if ($doc -ne $null) { $doc.Close($false) | Out-Null }
  if ($word -ne $null) { $word.Quit() | Out-Null }
  if ($doc -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null }
  if ($word -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
"""


def run_word_scan(path: Path, timeout: int) -> dict:
    if os.name != "nt":
        raise RuntimeError("Microsoft Word COM audit requires Windows.")
    with tempfile.TemporaryDirectory(prefix="cdut_word_toc_") as tmp:
        tmp_path = Path(tmp)
        ps1 = tmp_path / "scan_word_pages.ps1"
        json_path = tmp_path / "word_pages.json"
        ps1.write_text(WORD_SCAN_PS1, encoding="utf-8")
        cmd = [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(ps1),
            "-DocPath",
            str(path),
            "-JsonPath",
            str(json_path),
        ]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        if proc.returncode != 0:
            stderr = (proc.stderr or proc.stdout or "").strip()
            raise RuntimeError(f"Microsoft Word COM scan failed: {stderr}")
        return json.loads(json_path.read_text(encoding="utf-8-sig"))


def match_entries_to_word_pages(entries: list[dict], word_scan: dict, start_after_index: int | None) -> list[dict]:
    word_paragraphs = word_scan.get("paragraphs", [])
    matched: list[dict] = []
    for entry in entries:
        candidates = set(toc_label_candidates(entry["label"]))
        actual = None
        heading_text = None
        heading_index = None
        for para in word_paragraphs:
            pidx = int(para.get("paragraph_index", -1))
            if start_after_index is not None and pidx <= start_after_index:
                continue
            if normalize_match_text(str(para.get("text", ""))) in candidates:
                actual = str(para.get("adjusted_page"))
                heading_text = str(para.get("text", ""))
                heading_index = pidx
                break
        row = dict(entry)
        row.update(
            {
                "actual_page": actual,
                "matched_heading_text": heading_text,
                "matched_heading_index": heading_index,
                "ok": actual == entry["toc_page"] if actual is not None else False,
            }
        )
        matched.append(row)
    return matched


def audit_docx(path: Path, *, word_timeout: int = 180) -> dict:
    package = parse_docx_package(path)
    word_error = None
    word_scan: dict | None = None
    try:
        word_scan = run_word_scan(path, word_timeout)
    except Exception as exc:  # noqa: BLE001 - report, do not guess page numbers
        word_error = str(exc)

    entries = package["static_toc_entries"]
    toc_end_index = entries[-1]["paragraph_index"] if entries else package["toc_heading_index"]
    matched_entries = match_entries_to_word_pages(entries, word_scan or {}, toc_end_index) if word_scan else []
    mismatches = [entry for entry in matched_entries if entry.get("matched_heading_index") is not None and not entry.get("ok")]
    unmatched = [entry for entry in matched_entries if entry.get("matched_heading_index") is None]

    return {
        "input": str(path),
        "input_sha256": sha256_file(path),
        "word_available": bool(word_scan),
        "word_error": word_error,
        "physical_pages": word_scan.get("physical_pages") if word_scan else None,
        "last_adjusted_page": word_scan.get("last_adjusted_page") if word_scan else None,
        "word_paragraph_count": word_scan.get("word_paragraph_count") if word_scan else None,
        "has_toc_field": package["has_toc_field"],
        "static_toc_detected": bool(entries),
        "static_toc_entries_count": len(entries),
        "toc_heading_index": package["toc_heading_index"],
        "toc_entries": matched_entries,
        "matched_count": len([entry for entry in matched_entries if entry.get("ok")]),
        "mismatch_count": len(mismatches),
        "unmatched_count": len(unmatched),
        "mismatches": mismatches,
        "unmatched": unmatched,
        "notes": [
            "Microsoft Word AdjustedPageNumber is the authority for visible footer page numbers.",
            "LibreOffice/PDF/PNG page order is visual QA only unless explicitly requested.",
        ],
    }


def md_escape(text: object) -> str:
    return str(text if text is not None else "").replace("|", "\\|").replace("\n", " ")


def render_markdown(report: dict) -> str:
    lines = [
        "# Word actual-page TOC audit",
        "",
        f"- Input: `{report['input']}`",
        f"- Input SHA256: `{report['input_sha256']}`",
        f"- Word available: `{report['word_available']}`",
    ]
    if report.get("word_error"):
        lines.append(f"- Word error: `{md_escape(report['word_error'])}`")
    lines.extend(
        [
            f"- Word physical pages: `{report.get('physical_pages')}`",
            f"- Word last adjusted page: `{report.get('last_adjusted_page')}`",
            f"- TOC field detected: `{report['has_toc_field']}`",
            f"- Static TOC entries: `{report['static_toc_entries_count']}`",
            f"- Matched entries: `{report['matched_count']}`",
            f"- Mismatches: `{report['mismatch_count']}`",
            f"- Unmatched headings: `{report['unmatched_count']}`",
            "",
            "## TOC Entries",
            "",
            "| # | TOC label | TOC page | Word page | Status | Matched heading |",
            "|---:|---|---:|---:|---|---|",
        ]
    )
    for idx, entry in enumerate(report.get("toc_entries", []), 1):
        if entry.get("matched_heading_index") is None:
            status = "unmatched"
        elif entry.get("ok"):
            status = "ok"
        else:
            status = "mismatch"
        lines.append(
            "| {idx} | {label} | {toc_page} | {actual_page} | {status} | {heading} |".format(
                idx=idx,
                label=md_escape(entry.get("label")),
                toc_page=md_escape(entry.get("toc_page")),
                actual_page=md_escape(entry.get("actual_page")),
                status=status,
                heading=md_escape(entry.get("matched_heading_text")),
            )
        )
    lines.extend(
        [
            "",
            "## Notes",
            "",
            "- Use this report to repair static TOC page numbers. Do not use LibreOffice render page numbers for this decision.",
            "- If a live Word TOC field exists, prefer updating the field in Word only when the user explicitly asks for that higher-risk save operation.",
        ]
    )
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out", type=Path, help="Write Markdown audit report.")
    parser.add_argument("--json-out", type=Path, help="Write JSON audit report.")
    parser.add_argument("--word-timeout", type=int, default=180, help="Word COM scan timeout in seconds.")
    args = parser.parse_args(argv)

    try:
        report = audit_docx(args.docx, word_timeout=args.word_timeout)
    except Exception as exc:  # noqa: BLE001 - CLI should return clear failure
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    markdown = render_markdown(report)
    if args.out:
        args.out.write_text(markdown, encoding="utf-8")
    else:
        print(markdown)
    if args.json_out:
        args.json_out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    if not report.get("word_available"):
        return 1
    return 0 if report.get("mismatch_count") == 0 and report.get("unmatched_count") == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
