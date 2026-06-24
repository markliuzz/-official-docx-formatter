---
name: official-docx-formatter
description: >-
  Format Chinese official-document files locally. Use when the user asks to classify, inspect, standardize, convert, or format 公文、红头文件、国标格式、请示、报告、通知、函、公函、纪要、字体、行距、页边距、标点或空格规范. Convert non-.docx inputs to .docx locally when possible, then classify before formatting and ask only when the document type is uncertain. Change only Word styles, layout, page setup, hierarchy fonts, signature/date placement, and conservative punctuation/spacing; do not rewrite, polish, redact, summarize, or paste full body text.
---

# Official DOCX Formatter

## Purpose

Use this skill to turn an unformatted or inconsistently formatted `.docx` into a clean Chinese official-document style with a local trusted-engine workflow. For ordinary formatting, execute locally: classify the document, identify structure, diagnose formatting issues, build a formatting plan, apply formatting, and output a local report. For existing documents, classify the document type before changing the file. Different document types have different structure rules; formatting is only the second step.

For Chinese enterprise/national standard-specification drafts that contain a clear `目次` / `前    言` / `1  范围` structure, use the standard-specification branch in `scripts/format_docx.py`. This branch is for standard/specification documents, not ordinary official notices/reports or generic formal materials. It preserves body wording and order while applying standard-specification cover, table-of-contents, chapter/clause, page-number, and table-width rules.

Do not modify legacy enterprise-specific skills unless the user explicitly asks. This skill is the general-purpose formatter.

## Safety Boundary

- Read and classify the document through local scripts.
- Do not ask the model to rewrite, polish, summarize, or optimize body text.
- Do not print or paste the full body text into the conversation for ordinary formatting tasks.
- Change Word styles, margins, line spacing, hierarchy fonts, and page setup.
- Apply only conservative text normalization: Chinese punctuation, ellipsis/dash, paired quotes, and spacing around Chinese/English/digits.
- Preserve substantive wording, company names, amounts, contract numbers, project names, facts, and paragraph order.
- Preserve visible Word automatic numbering. When Word stores visible prefixes such as `一、` in `numbering.xml` instead of paragraph text, materialize those prefixes before rebuilding the document so formatted output does not silently drop hierarchy markers.
- For extreme glued drafts with only one long paragraph, an obvious official-document title, and embedded section-heading signals, conservatively split title, body paragraphs, section headings, and trailing signature-like lines before classification and formatting.
- When the source has no explicit numbering but contains short standalone structure lines such as `存在的问题`、`解决措施` or `时间计划`, format those lines as hierarchy headings. Treat short title-continuation lines before the body as subtitles/title lines when signals are clear.
- When one long paragraph uses spaces to separate structure blocks, preserve those boundaries before text normalization and recover nested hierarchy such as `一、存在的问题` and `（一）计量单位不符合商城商品规范` when the signals are explicit.
- Protect URLs, email addresses, times, standards, and code-like identifiers during punctuation normalization.
- Do not perform any redaction/refill cycle and do not desensitize user documents.
- Do not automatically add missing body text, issuer/signature, or date.
- Do not automatically install fonts; report missing fonts or use the default fallback guidance.
- If the user asks for content rewriting, explain that this formatter is format-only and suggest doing that as a separate task or separate skill.

## 🔴 STOP Conditions

Stop and ask before editing when:

- The source file is missing or cannot be read locally.
- A non-`.docx` source cannot be converted to `.docx` with available local tools.
- The document type is ambiguous or top candidates are close.
- The user asks to create a new official-document skeleton but has not provided or confirmed the document type.
- The user asks for rewriting, polishing, summarizing, redaction, or content optimization.
- The user asks for organization-specific formatting. This formatter no longer provides custom format profiles; use the default party-government standard configuration.

Do not generate a formatted output before the required choice or missing input is resolved.

## Default Format

Always use `profiles/standard-party-government.json`.

Do not ask the user to choose between standard and custom formats. If the user asks for an organization-specific format, explain that this formatter currently only supports the default party-government standard configuration.

## Failure Handling

| Trigger | Action |
| --- | --- |
| Missing input file | Ask for the file path; do not create a placeholder document. |
| Source file is not `.docx` | Try a local conversion to `.docx` first, preserving the original file. Use available local tools such as LibreOffice/soffice, Word automation, or `textutil` where appropriate. Continue with classification only after conversion succeeds. |
| Local conversion fails or no converter is available | Report the failed conversion compactly and ask the user for a `.docx` version; do not paste source text or manually rebuild the document from memory. |
| `python-docx` or package import fails | Report the missing dependency and use the project environment if available; do not rewrite the document manually. |
| Classifier returns low confidence or close candidates | Stop and ask: `我看这篇更像是【A】或【B】。你希望怎么处理：A / B / 通用正式文本 / 其他？` |
| Formatting command fails | Preserve the source file, report the error compactly, and do not claim a formatted output exists. |
| Report JSON fails but `.docx` is generated | Return the `.docx` path and clearly say the report was not generated. |
| Font is unavailable | Do not install fonts automatically; use default fallback guidance and include a warning. |
| User requests exact character preservation | Use `--no-normalize-text`. |

## Workflow

1. Identify the task:
   - Convert an existing `.docx`
   - Convert a non-`.docx` source to `.docx`, then format it
   - Create a new formatted `.docx`
   - Inspect the default formatting configuration
   - Answer a format requirement question

2. Read and classify before formatting:
   - If the source is a standard-specification draft with `目次`, body `前    言`, and `1  范围`, allow `scripts/format_docx.py` to auto-detect it or pass `--standard-spec-text` (`--standard-text` remains as a compatibility alias). Do not force it into one of the 15 ordinary official document types.
   - For an existing `.docx`, run or follow `scripts/classify_document.py` before making style changes.
   - Classify from title, body signals, ending phrases, recipient relationship, and obvious intent.
   - Use `references/document_types.json` as the catalog. It includes the 15 official document types from `党政机关公文处理工作条例`; common enterprise use is usually 请示、报告、通知、批复、通报、函/公函、纪要、决定.
   - If confidence is low or the top candidates are close, stop and ask the user to choose. Offer `通用正式文本` as a safe fallback for ordinary materials where the user only wants clean formal typography. Do not format first and ask later.
   - Use this short question: `我看这篇更像是【A】或【B】。你希望怎么处理：A / B / 通用正式文本 / 其他？`

3. Confirm or choose the document type:
   - For conversion, proceed directly only when the type is obvious from title or fixed ending phrases.
   - For creation, ask the user for the document type before drafting the title/body skeleton.
   - Read `references/document_types.md` when the type choice affects structure or wording.
   - `通用正式文本` is not an official document type. Use it only for conversion/formatting of existing materials, not for official-document skeleton creation.

4. Select formatting configuration:
   - Use `profiles/standard-party-government.json`.
   - Do not offer or create custom organization profiles.

5. For conversion:
   - Preserve original text order.
   - Reuse existing title, recipient, body, attachment notes, issuer, and date when detectable.
   - Do not add placeholder issuer/date if the source document has no clear issuer/date and the user did not provide them.
   - Treat visible automatic numbering as content for preservation purposes. Do not rely on `paragraph.text` alone when reconstructing paragraphs, because Word/WPS may store numbering separately from text.
   - If the source has collapsed all content into one long paragraph, repair paragraph boundaries only when the title and embedded section-heading signals are clear. Preserve wording and order; do not invent missing headings, issuer, or date.
   - If paragraph hierarchy is implicit, identify only short standalone, punctuation-free structure lines as headings/subtitles. Do not promote sentence-like body text or signature-like organization lines into headings.
   - For space-delimited one-paragraph drafts, split structure blocks before applying text normalization so boundary spaces are not lost. Add visible hierarchy numbering only to recovered heading lines, not to sentence-like body paragraphs.
   - Apply the confirmed document type's structural rules and the default configuration's typography, margins, line spacing, paragraph spacing, and hierarchy rules.
   - Normalize punctuation and spacing unless the user explicitly asks to preserve characters exactly. Use `--no-normalize-text` for exact-character preservation.
   - For issuer/date placement, follow the default configuration. It uses no-seal single-issuer placement: one blank line after the body or attachment note, then issuer/date on the right. Read `references/standards.md` before changing this behavior.
   - For font availability, follow `references/fonts.md`; do not auto-install fonts.
   - Output both the formatted `.docx` and the corresponding local `.report.json`.
   - For `通用正式文本`, preserve original paragraph order, preserve detected multi-line titles as title style, and format the rest as generic body/hierarchy paragraphs. Do not treat front matter such as `版本记录`、`需求文档检查表`、`目录` as title continuation. Do not extract or add recipient, issuer, date, red-head, imprint, or seal-related structure unless the user explicitly chooses a formal document type or passes a dedicated optional feature.
   - For standard-specification drafts, preserve the existing cover/body sequence, format `目次` with clearer chapter/clause hierarchy, merge split cover labels like `中华人民共和国` + `电力企业团体标准配套稿` into one cover line when detected, and narrow a first `序号`/`编号` table column while evenly distributing remaining content columns.

6. Verify:
   - Confirm the output `.docx` path and report JSON path.
   - Report the default configuration, document type, text-normalization status, and any warnings.
   - Run a smoke check when possible by opening the generated `.docx` with `python-docx` and reporting a compact status.
   - Do not paste the full document text or full report into the conversation; summarize the key status fields only.
   - If exact visual validation is needed, tell the user that Word/WPS rendering may need manual inspection.

## Scripts

- `scripts/format_docx.py`: Convert an existing `.docx` or create a simple formatted document from text.
- `scripts/classify_document.py`: Read `.docx` or text and return likely document type candidates.
- `scripts/diagnose_docx.py`: Inspect/audit a `.docx` locally and emit document type, structure, diagnostics, planned operations, and warnings. Use `--json` for machine-readable output.
- `scripts/normalize_text.py`: Conservative punctuation and spacing normalization without rewriting substantive content.
- `evals/evals.json`: Realistic test prompts and expectations for future iterations.
- `references/document_types.json`: Machine-readable document-type structures used for new document skeletons.

Run examples:

```bash
python scripts/classify_document.py input.docx
python scripts/diagnose_docx.py input.docx --json
python scripts/format_docx.py input.docx -o output.docx
python scripts/format_docx.py input.docx -o output.docx --no-normalize-text
python scripts/format_docx.py generic-material.docx -o output.docx --report --generic-formal-text
python scripts/format_docx.py standard-spec.docx -o output.docx --report --standard-spec-text
python scripts/format_docx.py -o report.docx --doc-type 报告 --title "关于××工作的报告" --recipient "上级单位" --issuer "某单位" --date 2026年6月19日 --create-skeleton
```

## References

Read these only when needed:

- `references/standards.md`: Standard basis and what is hard standard vs Word implementation.
- `references/document_types.md`: Difference between formatting profiles and document-type structures.
- `references/fonts.md`: Font fallback and missing-font policy.

## Packaging Hygiene

Before packaging or sharing this skill, exclude generated cache files such as `__pycache__/`, `.pyc`, temporary `.docx` outputs, and local evaluation workspaces. The skill should contain source instructions, profiles, references, scripts, and eval definitions only.

## Formatting Rules

Default hierarchy:

| Element | Default font | Size | Notes |
| --- | --- | --- | --- |
| Title | 小标宋体 fallback list | 2 hao / 22 pt | Centered |
| Body | 仿宋体 fallback list | 3 hao / 16 pt | Two-character first-line indent |
| Level 1 heading `一、` | 黑体 fallback list | 3 hao / 16 pt | Bold |
| Level 2 heading `（一）` | 楷体 fallback list | 3 hao / 16 pt | Not bold |
| Level 3 heading `1.` | Body font | 3 hao / 16 pt | Not bold |
| Level 4 heading `（1）` | Body font | 3 hao / 16 pt | Not bold |
| Word automatic numbering | Materialize visible prefixes such as `一、` before rebuilding |

Standard-specification extras:

| Element | Rule |
| --- | --- |
| Cover standard label | Merge detected split labels such as `中华人民共和国` + `电力企业团体标准配套稿` |
| TOC chapter item `1  范围` | No indent, bold |
| TOC clause item `3.1  闲置物资` | Indent one level, not bold |
| Body chapter heading `1  范围` | Heiti, bold |
| Body clause heading `3.1  闲置物资` | Heiti, not bold |
| Table first column `序号`/`编号` | Use a narrow first column and evenly distribute remaining content columns |

Only apply heading fonts to short standalone heading paragraphs. If a paragraph mixes a numbered prefix with sentence content, format the whole paragraph as body text.
