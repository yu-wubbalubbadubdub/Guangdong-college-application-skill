# 最终交叉验证指引

本文件只用于 `.xlsx` 志愿表草稿生成后的最终交叉验证。推荐清单生成、Excel 渲染和完成摘要撰写分别由其他 reference 负责；本文件是最终交付前的唯一验证门控。

## 读取时机

在以下文件都已经存在后读取本文件：

- `output/{考生代号}-{科类}-{位次}/candidate.json`
- `output/{考生代号}-{科类}-{位次}/profile-card.md`
- `output/{考生代号}-{科类}-{位次}/recommendations-review.md`
- `output/{考生代号}-{科类}-{位次}/recommendations-confirmed.json`
- `output/{考生代号}-{科类}-{位次}/volunteer-draft.xlsx`

如果任一文件缺失，不要进入最终交付；回退到对应前置步骤补齐。

## 验证范围

逐项核对以下内容：

- 资料卡是否包含核心字段：考生代号、年份、科类、分数、位次、目标批次、选科/专业统考类别、是否服从调剂、体检限制、单科/语种限制。
- 推荐清单是否来自最新资料卡，且用户已确认。
- `recommendations-confirmed.json` 是否只包含用户确认后的志愿，且沿用 `<skill_dir>/scripts/build_recommendations.py` 的统一 schema。
- 每条志愿是否包含录取批次、志愿号、院校代码、院校名称、院校专业组代码、至少 1 个专业代码或 `专业代码列表`、是否服从调剂、层级、备注。
- `代码复核状态` 与 `备注` 是否标明历史专业组/专业代码需要以 2026 官方招生专业目录复核；若已经使用 2026 官方目录修正，也应说明修正来源。
- 批次和填报时段是否符合 `<skill_dir>/references/2026-policy-rules.md`。
- 专业组代码、专业代码、院校代码是否标记了 2026 最新招生专业目录复核状态。
- 体检、色觉、单科、语种、性别、身高、收费、转专业和调剂范围风险是否已在备注或推荐清单中说明。
- 同一批次内艺术/体育/普通类兼报规则是否被检查。
- 高水平运动队/综合评价是否二选一；提前批、军检、定向培养军士、专项计划是否在对应栏目和时段。
- 志愿顺序是否存在明显倒挂。
- 目标批次覆盖是否与资料卡匹配；多数普通考生是否已在可接受范围内尽量接近或填满官方容量。
- 未满额时是否说明完整志愿数/官方容量、未满额例外类型、资料卡依据、已尝试扩展查询、继续补满的风险和可放宽约束建议。
- 所有运行产物是否写在项目根目录 `output/{考生代号}-{科类}-{位次}/`，没有写入 `<skill_dir>/output/`、`<skill_dir>/assets/` 或 `<skill_dir>/references/`。

## 必跑脚本

默认校验命令：

```bash
python <skill_dir>/scripts/validate_volunteer_output.py --file output/{考生代号}-{科类}-{位次}/volunteer-draft.xlsx --batch <目标批次> --format json
```

如果用户或策略明确要求目标批次满额，使用：

```bash
python <skill_dir>/scripts/validate_volunteer_output.py --file output/{考生代号}-{科类}-{位次}/volunteer-draft.xlsx --batch <目标批次> --format json --require-full
```

验证脚本只负责 Excel 内容完整性和覆盖率统计。政策、资料卡、章程、体检、调剂、费用、路径和完成摘要门控仍需按本文件人工交叉核对。尤其注意：未满额 warning 不能自动视为 PASS；只有推荐清单和资料卡能支持未满额例外时，最终验证才可通过。

## validation-report 结构

最终验证报告必须写入：

```text
output/{考生代号}-{科类}-{位次}/validation-report.md
```

使用以下结构：

```markdown
# 最终交叉验证报告

## 1. 验证结论
- overall_status: PASS / FAIL
- 是否允许生成 completion-summary.md:
- 验证时间：
- 目标批次：
- 是否启用 --require-full：

## 2. 文件完整性
- candidate.json:
- profile-card.md:
- recommendations-review.md:
- recommendations-confirmed.json:
- volunteer-draft.xlsx:
- validation-report.md:

## 3. Excel 内容校验
- validate_volunteer_output 命令：
- 脚本返回 ok：
- 官方容量：
- 表内行数：
- 完整填写数：
- 部分填写数：
- 脚本 errors：
- 脚本 warnings：

## 4. 政策与批次校验
- 目标批次：
- 填报时段：
- 志愿数量与投档方式：
- 特殊类别限制：
- 艺体/普通类兼报限制：

## 5. 资料卡匹配校验
- 分数/位次/科类：
- 院校偏好：
- 专业偏好：
- 地域偏好：
- 费用约束：
- 调剂偏好：
- 体检/单科/语种限制：
- 不得推荐项是否已排除：

## 6. 志愿策略与覆盖解释
- 采用策略：
- 各层级数量：
- 完整志愿数/官方容量：
- 是否达到官方容量：
- 未满额例外类型：strong_preference / high_score_strategy / hard_constraints / data_insufficient / policy_limited / user_requested_sparse / 无
- 未满额原因：
- 已尝试扩展查询：
- 继续补满的主要风险：
- 是否建议用户放宽约束：

## 7. 风险复核
- 专业组调剂风险：
- 招生章程待核查项：
- 2026 最新招生专业目录待核查项：
- 高收费项目：
- 体检/单科/语种风险：
- 志愿倒挂风险：

## 8. 阻塞项
- blocking_issues:
- 回退修复动作：
- 重新验证结果：

## 9. 最终交付许可
- 可以交付 volunteer-draft.xlsx：是/否
- 可以撰写 completion-summary.md：是/否
- 给用户的必要提醒：
```

## 通过标准

只有同时满足以下条件，`overall_status` 才能写 `PASS`：

- 必需文件均存在且位于项目根目录 `output/{考生代号}-{科类}-{位次}/`。
- `validate_volunteer_output.py` 返回 `ok: true`。
- 没有部分填写但不完整的志愿。
- 已填志愿字段完整，且不把空白补齐行计为已填志愿。
- 多数普通考生已在可接受范围内尽量接近或填满官方容量。
- 未满额时有清晰、可解释、与资料卡一致的原因，并给出未满额例外类型、已尝试扩展查询和可放宽约束建议。
- 对普通类本科/专科等大容量批次，如果完整志愿数明显少于官方容量且没有强偏好、高分强匹配、硬约束、数据不足或政策限制证据，不能写 `PASS`，应回退补充推荐清单。
- 没有违反用户明确排斥的专业、地区、费用或调剂风险。
- 没有未说明的体检、单科、语种、收费或章程风险。
- 没有批次、时段、兼报规则或特殊类型填报位置错误。

## 失败回退

如果发现问题，不要发送完成摘要。按问题类型回退：

- 资料卡缺字段：回到信息采集和资料卡生成步骤。
- 用户偏好冲突：回到推荐清单审阅步骤，向用户确认取舍。
- 推荐数据字段缺失：回到数据查询和推荐清单修订步骤。
- Excel 字段缺失或部分填写：修正 `recommendations-confirmed.json` 后重新渲染。
- 覆盖率解释不足：先回到推荐阶段扩展候选池；如果仍未满额，再补充未满额例外类型、扩展查询记录和可放宽约束建议后重新验证。
- 路径错误：移动或重新生成到项目根目录 `output/{考生代号}-{科类}-{位次}/`。
- `--require-full` 失败：补齐可接受志愿，或取消满额策略并说明原因后重新验证。

每次回退修复后，都必须重新运行验证脚本并更新 `validation-report.md`。
