---
name: cohort-protocol
description: Draft detailed, publication-ready Chinese retrospective cohort study protocol DOCX files from a user-provided protocol summary using the bundled latest corhortCRU-small.docx template, including structured summary, complete methodology sections, and recommended appendix tables. Use when the user asks for cohortProtocol/cohort-protocol, 回顾性队列研究方案, 队列研究方案, 临床研究方案, or wants a formatted .docx protocol from a 方案摘要.
---

# Cohort Protocol

## Goal

Create a complete, polished Chinese retrospective cohort study protocol from only the user's protocol summary, using the bundled template as the structural and stylistic authority.

Default output: `cohortprotocol.docx` in the current working directory, unless the user specifies another filename.

The DOCX must be publication-ready and methodologically detailed: clear section hierarchy, balanced page layout, formal Chinese academic typography, readable tables, consistent spacing, no visible drafting notes, and no unresolved placeholders. The default complete version should include the cover, an automatically updatable table-of-contents page immediately after the cover, structured `方案摘要`, sections 1-13, and the recommended appendix section with operation-definition and sensitivity-analysis tables whenever sufficient study details can be inferred.

## Bundled Resources

- Template asset: `assets/corhortCRU-small.docx`
- Template guide: `references/template-writing-guide.md`
- Extraction utility: `scripts/extract_template_map.py`
- DOCX renderer: `scripts/render_protocol_docx.py`

## Workflow

1. Read `references/template-writing-guide.md` first.
2. Resolve all bundled paths relative to this skill directory, not the user's current directory. This keeps the skill usable globally from any workspace.
3. Inspect the template before drafting. If more detail is needed, run:

```bash
python C:/Users/Lenovo/.codex/skills/cohort-protocol/scripts/extract_template_map.py C:/Users/Lenovo/.codex/skills/cohort-protocol/assets/corhortCRU-small.docx --out template-map.md
```

4. Analyze the user's summary against the template sections. Extract or infer:
   - title, study center, investigators if provided
   - target population and data source
   - exposure, comparator, index date, baseline window, follow-up and censoring
   - primary, secondary and safety outcomes
   - covariates, subgroup variables, missing-data plan and statistical models
   - ethics, data governance and conflicts of interest
   - appendix-ready operational definitions for exposure, comparator, outcomes, key covariates and sensitivity analyses
5. Draft the whole protocol in Chinese. Use the template's academic tone, section order and level hierarchy. For a full protocol, write substantive content for every required section rather than a short outline; most major sections should contain enough detail to support ethics review, data extraction and statistical analysis planning.
6. Put the user's supplied summary content into the `方案摘要` section by converting it into the template's summary table rows. Preserve all substantive information from the user summary.
7. Remove every instructional artifact from the final protocol:
   - `本方案模版使用说明`, `适用场景`, `注意`, `说明`
   - `【模板文本】`, `写作要点`, `【示例】`
   - unresolved template placeholders such as `【研究题目】`
   - square-bracket placeholders such as `[暴露名称]`
   - unused `（若适用）` markers
8. Render the DOCX through `scripts/render_protocol_docx.py` using a UTF-8 JSON content file. Prefer a file over stdin on Windows so Chinese text is not damaged by shell pipe encoding. The renderer contains the global publication layout and defaults to the bundled DOCX template when `--template` is omitted.
9. Re-open or extract the output DOCX and verify there are no template instructions, examples or unresolved placeholders.

## DOCX Publication Formatting Standard

The renderer enforces the global formatting baseline. Do not bypass it by manually composing a DOCX unless the user explicitly asks for a different format.

- Page: A4 portrait with formal research-protocol margins, running header after the cover page, and centered page numbers.
- Typography: Chinese body text uses Songti-style academic typography; headings use Heiti-style hierarchy; Latin text uses Times New Roman.
- Hierarchy: `level: 1` sections render as prominent first-level headings; `level: 2` sections render as second-level headings with keep-with-next behavior.
- Paragraphs: body paragraphs are justified, first-line indented, and spaced for readable printed review.
- Cover: render as a clean title page with centered study title/subtitle and a structured metadata table.
- Table of contents: render an automatically updatable `目录` page after the cover and before `方案摘要`, based on level 1-2 headings.
- Summary: render as a two-column structured table with emphasized row labels and readable cell spacing.
- Appendix: render the final appendix section in a slightly smaller font than the main body so it is visually distinct while remaining readable.
- Appendix tables: render operation-definition and sensitivity-analysis matrices as readable multi-column tables, not as prose-only lists.
- Cleanliness: final files must contain only study-specific protocol content, with no instructions, examples, bracket placeholders, or unused optional markers.
- Global use: always call the renderer by absolute path or from the skill root; the renderer itself locates `assets/corhortCRU-small.docx` relative to the skill directory.

## Drafting Rules

- Treat the user's summary as the source of truth. Do not contradict it.
- If a critical item is missing, infer a defensible default consistent with retrospective cohort methodology. Avoid visible placeholders; use neutral wording such as "拟提交研究实施单位伦理委员会审查" when a specific ethics committee is absent. In the complete protocol, do not expose uncertainty in core study elements when a methodologically sound definite default can be stated.
- The full protocol should be detailed enough for review. Do not answer with a skeletal template fill-in; expand the study rationale, eligibility rules, exposure/comparator algorithms, outcome validation, covariate measurement, bias control, missing data, sensitivity analyses, data governance and ethics sections.
- Core study elements in the complete protocol must use definite, review-ready wording. Do not use vague or approximate expressions such as `约`, `大概`, `预计约`, or similar uncertain phrasing for follow-up duration, observation windows, eligibility windows, sample feasibility statements, or other key design parameters.
- State time zero, index date, baseline window, follow-up start/end and censoring rules explicitly.
- In the `方案摘要` table, populate the `结局指标` row with the primary/secondary outcomes and their observation time windows. Do not list outcome names without time windows, and do not describe the observation period as a range that varies with enrollment timing.
- In the `3.3 时间零点（time zero）、基线期与随访` section, state the follow-up duration and cutoff explicitly (for example, follow patients from index date for 3 years, or observe outcomes within 6 years after index date). Use concrete time units such as years, months, or days rather than approximate wording. Do not write the total follow-up as a range that depends on inclusion timing, such as `总随访时间范围为0至6年（取决于纳入日期）` or `0–6年不等`. Do not use example-style patient-by-patient explanations such as introducing `例如` / `比如` and then describing how follow-up differs for patients enrolled on different dates. Instead, describe one cohort-level rule that applies uniformly to all included patients.
- In the `9. 结局指标定义与测量` section and its subsections, define each outcome together with its observation window from index date (for example, `自索引日期起3年内首次心衰住院`). Avoid vague wording such as `随访期间发生` without a bounded time window, and express the window as a clear number of years, months, or days. Every primary, secondary, and safety outcome should have its own explicit observation horizon rather than relying on one shared vague statement.
- For no-treatment, usual-care or historical comparators, justify the anchoring method and address immortal-time bias.
- Use target-trial emulation language when helpful, but keep the protocol readable and not over-theoretical.
- Include confounding control, covariate balance diagnostics, missing-data handling, sensitivity analyses and data privacy controls.
- Use high-standard academic Chinese with clear paragraphs. Avoid long literature-review digressions.
- Include the appendix section by default for complete protocols. Populate `附表 1` with study-specific operational definitions and `附表 2` with prespecified sensitivity analyses; omit appendix tables only when the user explicitly asks for a brief version or another format.

## Rendering Schema

Create JSON with this shape and pass it to the renderer:

```json
{
  "cover": {
    "研究题目": "string",
    "研究发起单位": "string",
    "牵头研究单位": "string",
    "主要研究者": "string",
    "参与单位/中心": "string",
    "统计分析负责人": "string",
    "数据管理负责人": "string",
    "方案编号": "string",
    "版本日期": "string",
    "版本号": "string"
  },
  "summary": [
    ["题目", "string"],
    ["研究中心", "string"],
    ["研究目的", "string"],
    ["研究设计", "string"],
    ["研究人群", "string"],
    ["暴露与对照", "string"],
    ["结局指标", "string"],
    ["样本量/可行性", "string"],
    ["主要协变量", "string"],
    ["统计分析", "string"],
    ["伦理与数据保护", "string"]
  ],
  "sections": [
    {
      "heading": "1. 研究背景与意义",
      "level": 1,
      "paragraphs": ["string"]
    },
    {
      "heading": "2.1 主要研究目的",
      "level": 2,
      "paragraphs": ["string"],
      "tables": [
        {
          "title": "string",
          "headers": ["string"],
          "rows": [["string"]]
        }
      ]
    }
  ]
}
```

Renderer usage:

```bash
python C:/Users/Lenovo/.claude/skills/cohort-protocol/scripts/render_protocol_docx.py protocol.json --output cohortprotocol.docx
```

The renderer applies the global publication-ready layout, preserves template-derived table styling where possible, and fails if obvious template artifacts remain.
