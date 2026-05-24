# Template Writing Guide

Use this guide after opening the `cohort-protocol` skill. The bundled DOCX template is the authority for formatting, section order and tone.

## Template Identity

- Name shown in template: `回顾性队列研究方案模板-龙医CRUv1.0`
- Issuer shown in template: `上海中医药大学附属龙华医院临床研究中心`
- Date shown in template: `2026年5月`
- Applicable studies: retrospective cohort studies using registries, electronic medical records, insurance claims, database linkage or multisource linked data.
- Latest bundled template update: includes a recommended appendix with operational definition and sensitivity-analysis matrix tables. Treat these appendix tables as part of the default complete protocol unless the user asks for a brief version.

## Template Cleaning Rules

Delete or replace these before final output:

- Template usage instructions and notes near the front matter.
- `【模板文本】` paragraphs as literal text.
- `写作要点` headings and their bullet guidance.
- `【示例】` content unless rewritten as study-specific text.
- `【...】` placeholder markers.
- Square-bracket placeholders from appendix examples, such as `[暴露名称]`, `[定义]`, `[编码]` and `[偏倚类型]`.
- `（若适用）` labels when the section is retained; delete inapplicable optional sections entirely when the study summary clearly does not need them.

## Publication Layout Requirements

Every generated DOCX should be suitable for direct research-protocol review, printing and circulation.

- Use A4 portrait page setup with balanced formal margins.
- Use a polished cover page: centered protocol title, clear document type, and a two-column metadata table for sponsor, site, investigators, version and date.
- Insert an automatically updatable `目录` page immediately after the cover page and before `方案摘要`, using level 1-2 headings.
- Use a structured two-column `方案摘要` table rather than free-form summary text.
- Use visually distinct heading levels so reviewers can scan the protocol quickly.
- Use justified Chinese body paragraphs, first-line indentation, readable line spacing and consistent paragraph spacing.
- Use formal Chinese academic typography: Songti-style body text, Heiti-style headings and Times New Roman for Latin text.
- Use running headers and page numbers after the cover page when rendered to DOCX.
- Render appendix definition tables as tables with clear headers and concise study-specific cell text.
- Render the final appendix section with a slightly smaller font than the main body to distinguish supporting materials from the main protocol.
- Avoid crowded pages: long sections should be divided into short paragraphs; tables should use concise but complete cell text.
- The output should not contain visible drafting instructions, example labels, placeholder brackets or unused optional-section markers.

## Required Output Structure

Preserve this order and heading hierarchy:

1. 方案封面
2. 目录
3. 方案摘要
4. 1. 研究背景与意义
5. 2. 研究目的
6. 2.1 主要研究目的
7. 2.2 次要研究目的
8. 3. 研究设计
9. 3.1 研究类型与总体框架
10. 3.2 设计选择的合理性
11. 3.3 时间零点（time zero）、基线期与随访
12. 3.4 偏倚最小化的设计考虑
13. 4. 研究对象
14. 4.1 数据来源
15. 4.2 目标人群、来源人群与队列构建
16. 4.3 纳入标准
17. 4.4 排除标准
18. 4.5 样本量估算
19. 5. 暴露因素定义与选择
20. 5.1 暴露定义的依据
21. 5.2 暴露的识别定义
22. 6. 对照组定义与选择
23. 6.1 对照组定义的依据
24. 6.2 对照组的识别定义
25. 7. 适应证、禁忌证和治疗强度的可比性
26. 8. 联合用药、换药、停药、依从性与误分类处理
27. 9. 结局指标定义与测量
28. 9.1 主要结局
29. 9.2 次要结局与安全性结局
30. 9.3 结局识别来源与验证
31. 9.4 重复事件与竞争事件
32. 9.5 替代定义、敏感性分析与特殊结局说明
33. 10. 协变量选择与测量
34. 10.1 协变量选择原则与因果框架
35. 10.2 人口学和社会经济学变量
36. 10.3 临床特征、疾病严重程度与合并症
37. 10.4 合并用药、合并治疗与医疗利用指标
38. 10.5 效应修饰因素与预设亚组变量
39. 10.6 协变量测量窗口与编码
40. 11. 统计分析方法
41. 11.1 总体分析原则与分析集
42. 11.2 描述性统计与基线平衡评估
43. 11.3 混杂控制策略与模型选择
44. 11.4 主要结局分析
45. 11.5 次要结局、亚组分析与效应异质性
46. 11.6 时间变化暴露、停药与依从性分析
47. 11.7 缺失数据处理
48. 11.8 敏感性分析
49. 11.9 统计软件、显著性水平与结果报告
50. 12. 数据管理
51. 12.1 数据提取、链接与去标识化
52. 12.2 数据管理与质量控制
53. 13. 伦理考量
54. 13.1 伦理审查
55. 13.2 知情同意获取方式或豁免
56. 13.3 隐私保护与去标识化
57. 13.4 利益冲突声明
58. 附录：建议附件与表单
59. 附表 1 暴露、对照、结局与协变量操作性定义总表
60. 附表 2 敏感性分析矩阵

Optional sections can be retained when clinically or methodologically relevant. If retained, remove the `若适用` phrase from headings.

For complete protocols, the appendix should usually be retained. It operationalizes the protocol for data extraction and analysis, and is especially useful when the user's summary includes specific exposure, comparator, outcome or covariate definitions.

## Summary Table Rows

The `方案摘要` section should be a two-column table with these rows:

- 题目
- 研究中心
- 研究目的
- 研究设计
- 研究人群
- 暴露与对照
- 结局指标
- 样本量/可行性
- 主要协变量
- 统计分析
- 伦理与数据保护

The "结局指标" row should briefly summarize primary and secondary outcomes with their observation time windows (e.g., "自索引日期起3年内首次心衰住院；次要结局为自索引日期起3年内全因死亡"). Do not list outcome names without time windows. Do not write the observation period as a range that changes with enrollment timing.

Translate the user's summary into these rows. Do not paste an unstructured abstract only; the template expects a structured summary table.

## Section Writing Expectations

Default completeness standard:
A complete protocol should read like a review-ready study protocol rather than a filled outline. Expand each section with study-specific rationale, operational definitions and implementation rules. If the user summary is sparse, infer defensible defaults and state them as planned methods, not placeholders. Major methods sections should be detailed enough for a statistician or data manager to build the extraction plan, analysis set and statistical analysis plan from the document. For core study elements, use definite, review-ready wording rather than vague approximations; avoid expressions such as `约`, `大概`, `预计约` or other uncertainty markers in the final complete protocol when a defensible methodological default can be stated.

研究背景与意义:
Explain disease burden, clinical gap, exposure rationale and why the chosen real-world data can answer the question. Keep it concise and focused on the evidence gap.

研究目的:
State one primary objective tied to the primary outcome. Secondary objectives can cover secondary clinical outcomes, safety outcomes, subgroup analyses, exposure duration, dose intensity or adherence.

研究设计:
Specify data source, study scope, retrospective cohort design, comparator type and whether a new-user design is used. Define index date, time zero, baseline assessment window, follow-up start, follow-up end and censoring. State the follow-up duration and cutoff explicitly rather than using vague wording; for example, clarify whether outcomes are observed for 1 year, 3 years, 5 years, or until a fixed database cutoff date. Express the duration with clear year/month/day units rather than approximate language. Do not describe total follow-up as a range that depends on enrollment timing, such as `总随访时间范围为0至6年（取决于纳入日期）` or `0–6年不等`; instead, frame it as observing outcomes within a fixed horizon after index date. Do not explain the rule by giving patient-by-patient examples for different enrollment dates; write one cohort-level rule that applies uniformly to all included patients. Explain measures to reduce selection bias, prevalent-user bias, immortal-time bias and confounding.

研究对象:
Describe the data source and source population. Define inclusion/exclusion criteria using information available before time zero. Explain sample size as database feasibility and expected event count, or provide a formal calculation when the summary includes assumptions.

暴露因素:
Justify the exposure definition clinically and operationally. Define codes, prescriptions, minimum duration, washout, dose/intensity, grace period and exposure risk window when relevant.

对照组:
Define active comparator, no-treatment comparator or usual-care comparator. For no-treatment/usual-care designs, explain comparable anchoring and clinical eligibility to avoid immortal-time bias.

适应证、禁忌证和治疗强度:
Use this section for drug, procedure or treatment comparisons where clinical indication and treatment intensity may differ between groups. Discuss eligibility, contraindications, baseline disease severity and treatment decision context.

联合用药、换药、停药、依从性与误分类:
Define concomitant therapies, treatment switching, discontinuation, adherence metrics such as proportion of days covered, permissible grace periods and exposure misclassification handling.

结局指标:
Define primary, secondary and safety outcomes with data sources, codes, validation strategy and adjudication if available. For every outcome, explicitly define the observation window from index date or time zero (for example, `自索引日期起3年内首次心衰事件`). Do not use unbounded wording such as `随访期间发生` without specifying the observation period. Express every observation window as a clear number of years, months, or days. State the horizon separately for each primary, secondary, and safety outcome rather than relying on one shared vague statement. Address repeated events, competing risks, proxy endpoints, composite endpoints and sensitivity definitions.

协变量:
Choose covariates based on causal reasoning and availability before time zero. Cover demographics, socioeconomic indicators, clinical severity, comorbidities, co-medications, procedures, health-care utilization and prespecified effect modifiers.

统计分析:
Describe analysis set, baseline balance, confounding control, primary model, secondary outcomes, subgroup analyses, time-varying exposure methods, missing-data handling, sensitivity analyses, software and significance/reporting principles.

数据管理:
Describe extraction, linkage, de-identification, data layers, quality checks, audit trail, script versioning, access control, encryption, backup and retention.

伦理考量:
Describe ethics review, informed-consent waiver rationale, privacy protections, de-identification, prohibition on re-identification, funding role and conflicts of interest.

附录:
Retain the appendix for full protocols. `附表 1` should list study-specific operational definitions for exposure, comparator, primary outcome, secondary/safety outcomes and core covariates. Columns should follow the template: 类别, 变量名称, 操作性定义, 数据来源, 时间窗, 编码/算法, 备注. `附表 2` should list prespecified sensitivity analyses with the columns: 敏感性分析编号, 目的, 变更内容, 对应偏倚/假设, 结果指标. Avoid appendix placeholders; every row must contain concrete study-specific or methodologically neutral content.

## Quality Checklist

Before finalizing, confirm:

- The design is clearly retrospective cohort, not case-control or cross-sectional.
- Exposure precedes outcome by design.
- Time zero is aligned across groups.
- Baseline covariates are measured before time zero.
- Follow-up and censoring rules are explicit.
- Comparator choice is justified.
- Confounding control is prespecified.
- Missing data and sensitivity analyses are included.
- The appendix operational definition table is populated when a complete protocol is requested.
- The sensitivity-analysis matrix contains prespecified, study-specific robustness checks.
- No template instruction, placeholder, writing note or example label remains.
- The output is a DOCX with the template's hierarchy and readable formatting.
