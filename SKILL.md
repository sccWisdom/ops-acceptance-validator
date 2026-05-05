---
name: ops-acceptance-validator
description: Use this skill whenever the user asks to check, review, validate, audit, or produce correction suggestions for 运维验收材料, 验收材料, 运维验收文档, 服务验收文件, or uploaded .docx/.xlsx acceptance materials. This skill reads Word and Excel files and returns only concise Markdown correction suggestions for business users.
---

# 运维验收材料检查

Use this skill to check uploaded acceptance materials and return only actionable correction suggestions. The user expects a clean business-facing answer, not a report.

## Completion Standard

Before replying, verify that:

- `.docx` and `.xlsx` files were read successfully.
- Word正文段落、Word表格、Excel所有工作表、单元格、合并单元格表述、表头、备注和说明性文字 were considered.
- The five rule families were checked:
  - 服务周期：2025-08-26 至 2026-02-28.
  - 数量一致性：19个子系统、834台服务器.
  - 子系统名称：only the 19 standard subsystem names.
  - 厂商相关泛称：厂商、服务商、运营商.
  - 具体合作伙伴名称：replace with the mapped 实施编号组.
- The final answer contains only correction suggestions, or exactly `未发现需要修改的内容。`
- The final answer does not include JSON, statistics, debug output, rule explanations, pass items, complete revised text, or a full inspection report.

## Recommended Workflow

1. Collect all uploaded or referenced `.docx` and `.xlsx` file paths. Keep the user-provided order.
2. Run the bundled checker:

```bash
python scripts/ops_acceptance_validator.py <file1.docx> <file2.xlsx>
```

3. Inspect the output before replying. Confirm it is not garbled and contains only Markdown correction suggestions.
4. Reply with the checker output exactly, unless you must remove accidental non-suggestion text.

If the checker cannot run, manually extract the document text and apply the same rules. Keep the same final output constraints.

## Output Format

When issues exist, use numbered Markdown lines:

```markdown
1. “文件名称”中“原文片段”存在问题，建议……
2. “文件名称”的“工作表名称”工作表中“原文片段”存在问题，建议……
```

When no issues exist, output exactly:

```markdown
未发现需要修改的内容。
```

## Output Rules

- Each suggestion must include the file name.
- Include the worksheet name for Excel issues when available.
- Include the original phrase that needs attention.
- Explain why it needs revision in plain business language.
- Give one clear, executable recommendation.
- Handle one issue per suggestion.
- Preserve file order and original occurrence order.
- Do not output pass items or harmless content.
- For uncertain short partner names, phrase the suggestion as a核实/替换 suggestion instead of claiming certainty.

## Rule Details

### 服务周期

Valid period: `2025-08-26` through `2026-02-28`.

Flag dates, months, and periods outside this range. Recognize common forms such as:

- `2025.8.26`
- `2025/08/26`
- `2025-08-26`
- `2025年8月26日`
- `20250826`
- `2025年8月`
- `2025-08`

For early dates, suggest changing to `2025-08-26` or later. For late dates, suggest changing to `2026-02-28` or earlier.

### 数量一致性

Flag only when the context clearly refers to the project-wide quantity:

- Subsystems must be `19个子系统`.
- Servers must be `834台服务器`.
- Recognize Arabic and Chinese numerals, such as `18个子系统`, `十八个子系统`, `833台服务器`, `八百三十三台服务器`.

### 子系统名称

Only these 19 names are standard:

1. 长三角共享交换平台
2. 运营管理子系统
3. 运维监控子系统
4. 数据资产子系统
5. 数据治理分析系统
6. 数据支撑子系统
7. 数据开发子系统
8. 数据驾驶舱
9. 数据归集子系统
10. 数据共享子系统
11. 数据标签子系统
12. 上海市数据分析平台
13. 上海市公共数据开放平台
14. 前置机管理子系统
15. 空间地理资产管理子系统
16. 国家资源平台级联系统
17. 公共数据管理门户
18. 电子证照管理系统
19. 安全管理子系统

Flag non-standard, approximate, abbreviated, duplicated, missing, or extra names. Do not automatically rewrite doubtful names; suggest verification.

### 厂商相关泛称

Flag:

- 厂商
- 服务商
- 运营商

Suggest deleting the phrase or replacing it with `实施团队`. Do not suggest another phrase containing 厂商、服务商、运营商.

### 具体合作伙伴名称

The checker script contains the complete mapping from partner names to `实施{编号}组`. Use its output as the source of truth.

Important handling:

- Match longer names before shorter names.
- Match `cert`, `CERT`, and `Cert` as `实施13组`.
- Treat Chinese and English parentheses as equivalent in company names.
- For short names that might be ordinary words, make the suggestion cautious and ask the user to核实后替换.
