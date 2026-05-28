# Skill Index

This repository contains five Codex skills for thesis and thesis-defense formatting.

## Thesis DOCX Skills

| Skill | Core Inputs | Core Outputs |
| --- | --- | --- |
| `cdu-thesis-template` | CDU undergraduate thesis DOCX and authorized science template asset | Format audit, layout guidance, template-compatible DOCX edits |
| `cdut-english-thesis-template` | CDUT English-major thesis DOCX and authorized format/citation template assets | Format audit, MLA citation-format guidance, front-matter/body layout cleanup |
| `cdut-geophysics-thesis-format` | CDUT Geophysics thesis DOCX and authorized DOTX/template guide assets | OOXML audit, deterministic repair copy, Word actual-page TOC audit/repair |
| `swpu-thesis-template` | SWPU thesis or graduation-design DOCX and authorized template asset | Format audit, structure checks, page/header/footer guidance |

## Defense PPT Skill

| Skill | Core Inputs | Core Outputs |
| --- | --- | --- |
| `cdu-defense-ppt-template` | CDU defense PPTX and authorized defense template asset | Template-style audit and slide-layout adaptation guidance |

## Notes

- The skills are format-first. They preserve thesis content unless the user explicitly asks for writing or rewriting.
- Scripted audits operate on OOXML packages and are designed to read or write copies, not overwrite the only user file.
- Template assets are intentionally excluded from Git history; users must provide authorized copies locally.
