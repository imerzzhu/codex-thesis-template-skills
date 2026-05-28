# Codex Thesis Template Skills

面向高校论文、毕业设计和答辩材料格式化的 Codex Skills 集合。仓库提供可复用的格式规则、审查脚本和工作流说明，帮助 Codex 按指定学校模板审查或整理 DOCX/PPTX 文件。

## English Summary

Codex Thesis Template Skills is a public Codex Skills package for thesis and defense-presentation formatting. It includes reusable skills for auditing and repairing Word thesis documents, checking table-of-contents page numbers, and adapting defense PowerPoint decks to a supplied template style.

## Skill List

### Thesis DOCX

| Skill | Use |
| --- | --- |
| `cdu-thesis-template` | Format and audit Chengdu University undergraduate science/engineering thesis DOCX files. |
| `cdut-english-thesis-template` | Format and audit Chengdu University of Technology English-major thesis DOCX files, including MLA-style Works Cited guidance. |
| `cdut-geophysics-thesis-format` | Audit and repair CDUT Geophysics College thesis/design DOCX files, including Word actual-page TOC checks. |
| `swpu-thesis-template` | Format and audit Southwest Petroleum University undergraduate thesis or graduation-design DOCX files. |

### Defense PPT

| Skill | Use |
| --- | --- |
| `cdu-defense-ppt-template` | Adapt and audit Chengdu University thesis defense PPTX files against the supplied red-white template style. |

## Template Assets

This repository does not redistribute school template files such as `.doc`, `.docx`, `.dotx`, or `.pptx`. Those files may be subject to school or third-party licensing terms.

To use a template-backed skill, place your authorized copy of the relevant template file under that skill's `assets/` directory with the expected filename. See [docs/template-assets.md](docs/template-assets.md).

## Install

Windows PowerShell example:

```powershell
$repo = "D:\Development\Projects\codex-thesis-template-skills"
$codexSkills = "path\to\your\Codex\skills"

New-Item -ItemType Directory -Force -Path $codexSkills | Out-Null
Copy-Item -Recurse -Force "$repo\skills\cdu-thesis-template" $codexSkills
Copy-Item -Recurse -Force "$repo\skills\cdut-english-thesis-template" $codexSkills
Copy-Item -Recurse -Force "$repo\skills\cdut-geophysics-thesis-format" $codexSkills
Copy-Item -Recurse -Force "$repo\skills\swpu-thesis-template" $codexSkills
Copy-Item -Recurse -Force "$repo\skills\cdu-defense-ppt-template" $codexSkills
```

If you keep large template files or generated reports locally, prefer a D-drive work/cache directory such as `D:\Cache\Dev\thesis-template-skills`.

## Usage Examples

```text
Use $cdu-thesis-template to audit this thesis DOCX against the Chengdu University science thesis template without changing thesis content.
```

```text
Use $cdut-geophysics-thesis-format to repair this CDUT geophysics thesis DOCX against the DOTX template and verify static TOC page numbers with Microsoft Word actual pages.
```

```text
Use $cdu-defense-ppt-template to adapt this thesis defense PPTX to the Chengdu University defense template visual system.
```

## Project Structure

```text
.
|-- skills/
|   |-- cdu-thesis-template/
|   |-- cdut-english-thesis-template/
|   |-- cdut-geophysics-thesis-format/
|   |-- swpu-thesis-template/
|   `-- cdu-defense-ppt-template/
|-- docs/
|   |-- open-source-policy.md
|   |-- skill-index.md
|   `-- template-assets.md
|-- LICENSE
`-- README.md
```

Each skill follows the Codex Skills layout:

```text
skill-name/
|-- SKILL.md
|-- agents/openai.yaml
|-- assets/.gitkeep    # user-provided templates go here
|-- references/        # optional
`-- scripts/           # optional
```

## Open Source Boundary

The MIT license applies to the code and documentation in this repository. It does not grant rights to school templates, user thesis files, generated reports, rendered pages, or any template files users place in `assets/`.

See [docs/open-source-policy.md](docs/open-source-policy.md) for details.
