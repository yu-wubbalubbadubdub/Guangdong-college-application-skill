# 官方志愿表与最终输出指引

本 skill 使用 `<skill_dir>/assets/2026年广东省普通高校招生考生志愿表（官方）.pdf` 作为官方附件原表，用于核对字段、批次和样式。最终可编辑 Excel 草稿必须基于 `<skill_dir>/assets/2026广东高考志愿表填报模板.xlsx` 复制填充；不要从零生成最终 Excel 文件。

## 最终输出策略

进入本步骤前，`recommendations-confirmed.json` 必须沿用 `<skill_dir>/scripts/build_recommendations.py` 输出的统一 schema。特别是以下字段会被 `render_volunteer_draft.py` 直接用于 Excel：

- `录取批次`
- `志愿号`
- `院校代码`
- `院校名称`
- `院校专业组代码`
- `专业代码列表` 或 `专业代码1`-`专业代码6`
- `是否服从调剂`
- `层级`
- `备注`

若用户在审阅后删除或调整志愿，只能在 `intermediate/recommendations-proposed.json` 的基础上筛选、重排和重写 `志愿号`，不要改字段名。若专业代码、专业组代码仍来自 2022-2025 历史数据，`备注` 和 `代码复核状态` 必须保留“待2026官方招生专业目录复核”的提示；最终提交前必须用 2026 官方目录补齐或确认。

优先生成可编辑 Excel 志愿表草稿。命令默认从项目根目录执行。脚本会复制 `<skill_dir>/assets/2026广东高考志愿表填报模板.xlsx` 并替换其中的工作表内容：

```bash
python <skill_dir>/scripts/render_volunteer_draft.py --input output/{考生代号}-{科类}-{位次}/recommendations-confirmed.json --template <skill_dir>/assets/2026广东高考志愿表填报模板.xlsx --output output/{考生代号}-{科类}-{位次}/volunteer-draft.xlsx --pad-to 45
```

正式交付前建议启用严格校验，默认检查已填志愿的字段完整性和覆盖率提示：

```bash
python <skill_dir>/scripts/render_volunteer_draft.py --input output/{考生代号}-{科类}-{位次}/recommendations-confirmed.json --template <skill_dir>/assets/2026广东高考志愿表填报模板.xlsx --output output/{考生代号}-{科类}-{位次}/volunteer-draft.xlsx --pad-to 45 --strict-final
```

如果用户或推荐策略明确要求目标批次满额，再追加 `--require-full`：

```bash
python <skill_dir>/scripts/render_volunteer_draft.py --input output/{考生代号}-{科类}-{位次}/recommendations-confirmed.json --template <skill_dir>/assets/2026广东高考志愿表填报模板.xlsx --output output/{考生代号}-{科类}-{位次}/volunteer-draft.xlsx --pad-to 45 --strict-final --require-full
```

如果输出为 `.xlsx` 且未显式传入 `--pad-to`，脚本会按批次自动补齐常见官方志愿位数量；普通类本科/专科自动补齐到 45 个志愿位。正式交付仍应显式使用 `--pad-to 45`，便于审阅命令意图。严格最终表要求已填志愿完整、可解释且符合资料卡；完整填写的定义是：院校代码、院校名称、院校专业组代码、至少 1 个专业代码、是否服从调剂均已确认。脚本补齐的空白可编辑行只代表官方容量，不计入已填志愿。

未填满不一定是错误，但不能成为默认交付方式。多数普通考生应在可接受范围内尽量扩大合适候选池，接近或填满官方志愿位以增加录取机会。若资料卡体现强院校/专业倾向、高分强匹配策略、明确排斥范围较大，或继续补充会明显降低匹配度，可以保留空白志愿位；但推荐清单、`填报提醒` 和最终验证报告必须同时说明完整志愿数/官方容量、未满额例外类型、已尝试扩展查询、继续补满会引入的风险，以及是否建议用户放宽约束。

如果目标批次未达到官方容量，且推荐清单没有上述例外证据，不要生成最终交付版 `.xlsx`；应回到 `<skill_dir>/references/recommendation-workflow.md` 扩展候选池。

如果只需要文本预览，可输出 Markdown：

```bash
python <skill_dir>/scripts/render_volunteer_draft.py --input output/{考生代号}-{科类}-{位次}/recommendations-confirmed.json --output output/{考生代号}-{科类}-{位次}/volunteer-draft.md --pad-to 45
```

草稿字段必须与官方表一致：

- 录取批次
- 志愿号
- 院校代码
- 院校名称
- 院校专业组代码
- 专业代码 1-6
- 是否服从调剂

Excel 草稿必须至少包含：

- `志愿表草稿` 工作表：官方字段风格表格，可直接编辑。
- `填报提醒` 工作表：官方系统确认、填报时段、2026 目录复核和章程复核提醒。
- 来源说明：该文件是从 `<skill_dir>/assets/2026广东高考志愿表填报模板.xlsx` 复制并填充得到，不是临时从零生成的空表。
- `志愿表草稿` 工作表必须含有明确的 Excel worksheet 范围元数据，避免部分预览器或 WPS/Excel 版本只显示已填推荐行而忽略补齐空白志愿位。
- 模板和 `.xlsx` 输出必须覆盖官方 PDF 中的全部录取批次，可保留“层级”和“备注”列用于风险管理。

生成 `.xlsx` 后进入最终交叉验证。内容验证命令由 `<skill_dir>/references/final-cross-validation-guide.md` 统一规定：

```bash
python <skill_dir>/scripts/validate_volunteer_output.py --file output/{考生代号}-{科类}-{位次}/volunteer-draft.xlsx --batch 普通类本科院校 --format json
```

验证脚本会按实际完整填写条数统计覆盖情况，而不是只检查 Excel 行数。默认情况下，未满额会输出警告和覆盖率提示；只有追加 `--require-full` 时，未满额才会导致脚本失败。脚本通过不等于最终可交付：最终交叉验证仍必须判断未满额是否有资料卡依据和扩展查询记录。字段完整性、政策、资料卡、风险和路径的最终门控不要写在本文件中，统一读取 `<skill_dir>/references/final-cross-validation-guide.md`。

如用户要求可打印版，再基于草稿和官方 PDF 另行制作 PDF/Word 版本。

## 填表前置条件

不要在以下条件未满足时生成最终表：

- 用户已确认推荐清单。
- 推荐清单已在资料卡约束内覆盖足够多的合适目标；多数普通考生应接近或达到官方容量。若少于官方容量，已说明未满额例外类型、完整志愿数/官方容量、已尝试扩展查询、继续补满的风险和可放宽约束建议。
- 已核对最新招生专业目录中的院校代码、院校专业组代码、专业代码。
- 已核对招生章程和体检限制。
- 已核对填报时段。
- 已核对是否服从调剂。
- 已核对专业组内所有可能调剂专业是否可接受。

## 志愿表批次摘要

普通 2026 夏季高考常见输出：

- 普通类本科院校：45 个平行志愿，第二时段。
- 普通类专科院校：45 个平行志愿，第二时段。
- 军检院校：10 个平行志愿，第一时段。
- 专科提前定向培养军士：10 个平行志愿，第一时段。
- 艺体本科：20 个平行志愿，第二时段。
- 艺体专科：20 个平行志愿，第二时段。

完整规则见 `references/2026-policy-rules.md`。

## 输出命名

所有最终产物写入项目根目录的同一个任务目录：

```text
output/{考生代号}-{科类}-{位次}/
```

建议文件名固定为：

```text
candidate.json
profile-card.md
recommendations-review.md
recommendations-confirmed.json
volunteer-draft.xlsx
volunteer-draft.md
validation-report.md
completion-summary.md
```

## 最终提示语

交付时必须提醒：

- 已按 `<skill_dir>/references/completion-summary-guide.md` 生成任务完成摘要，并把本次策略、依据、优势、风险和产物路径讲清楚。
- 这是基于资料和用户偏好生成的志愿填报草稿，不替代广东省教育考试院系统中的最终确认。
- 考生必须在官方系统规定时间内完成网上确认。
- 志愿最终以广东省教育考试院系统中最后确认版本为准。
