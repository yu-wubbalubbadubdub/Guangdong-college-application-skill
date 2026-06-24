---
name: guangdong-gaokao-application
description: >-
  当用户需要为广东省2026届高考生进行志愿填报、院校专业推荐、冲稳保垫兜底分析、广东高考数据查询、生成广东省普通高校招生考生志愿表草稿时使用。支持采集考生资料、生成本地资料卡、读取广东2026志愿填报政策、查询22-25年院校/专业录取分数与招生计划数据、复核退档风险，并输出基于官方模板的可编辑Excel志愿表草稿。
---

# 广东 2026 高考志愿填报

把本文件作为流程入口。执行具体任务时，按需读取 `references/` 中的指引文件，优先使用 `scripts/` 查询结构化数据，使用 `assets/` 中的官方 PDF 作为字段依据，并使用 `assets/2026广东高考志愿表填报模板.xlsx` 作为最终可编辑 Excel 草稿的空白模板。

## 参考文件导航

必须从 `SKILL.md` 直接触达所有参考文件：

- `references/profile-card-guide.md`：考生资料卡的字段、结构和写法。
- `references/2026-policy-rules.md`：广东省 2026 年志愿填报时间、批次、确认规则、特殊类别和政策风险。
- `references/data-sources.md`：`references/data/` 内所有数据工作簿的目录、字段、脚本 source 名和使用优先级。
- `references/recommendation-workflow.md`：冲、稳、保、垫、兜底分层方法、推荐清单字段和风险表达方式。
- `references/final-form-guide.md`：最终志愿表草稿的复核步骤、字段顺序和官方系统提醒。

`references/data/` 是本技能自带的数据源目录，脚本默认从这里读取：

- `references/data/scores/`：22-25 年院校录取分数、专业录取分数。
- `references/data/plans/`：22-25 年在粤招生计划。
- `references/data/segments/`：2022-2025 年广东一分一段表。
- `references/data/charters/`：2025 年高校招生章程链接表。
- `references/data/art/`：艺术类院校/专业录取分数与招生计划。
- `references/data/spring/`：2026 春考、学考、3+证书数据。

`references/data/` 下的 `.xlsx` 大表只能通过 `scripts/query_admissions.py` 或 `scripts/equivalent_score.py` 访问。不要直接读写这些二进制数据表，不要手工猜字段。

`assets/` 存放最终填表相关资产：

- `assets/2026年广东省普通高校招生考生志愿表（官方）.pdf`：官方附件原表，用于核对字段、批次和表格样式。
- `assets/2026广东高考志愿表填报模板.xlsx`：可编辑空白 Excel 志愿表模板，最终 `.xlsx` 草稿必须复制并填充此模板，不要从零生成 Excel 文件。

## 主流程

任一步出现脚本错误、核心字段缺失、目标院校无法精确匹配、专业代码缺失、招生章程无法核查、用户偏好冲突或政策限制未确认时，先停止当前阶段，标记原因并补查或追问。不要让错误或缺失静默进入下一阶段。

1. **先采集考生信息。**  
   先收集必填信息：姓名或昵称、高考分数、全省位次、科类（物理类/历史类）、目标批次、选科组合、是否为艺术类/体育类/春考/学考/3+证书考生、是否服从调剂。缺少这些信息时不要进入正式推荐。随后收集偏好补充项：倾向院校、倾向专业、明确排斥专业、倾向地域、明确排斥地域、家庭经济承受能力、是否接受高收费项目、个人兴趣、职业规划、体检限制、单科成绩限制、外语语种限制、特殊资格，以及用户主动补充的其他择校择专业约束。

2. **生成或更新考生资料卡。**  
   读取 `references/profile-card-guide.md`，把已采集信息整理成结构化资料卡，并保存到 `profiles/`。如果用户后续补充或修正偏好，先更新资料卡，再继续查询或推荐。可使用：
   `python scripts/create_profile_card.py --input candidate.json --output profiles/candidate-profile.md`

3. **加载政策和数据说明。**  
   推荐前读取 `references/2026-policy-rules.md` 和 `references/data-sources.md`。确认填报时段、批次类别、院校专业组规则、合格考要求、艺术/体育兼报限制、特殊类型填报位置、数据年份、字段口径和数据风险。

4. **用脚本查询数据。**  
   不要手工解析 Excel。先用 `python scripts/query_admissions.py --list-sources` 查看可用 source，再用 `--inspect` 查看字段。常用查询：
   - `python scripts/equivalent_score.py --year 2025 --subject 物理类 --rank 50000`
   - `python scripts/query_admissions.py --source school-scores --year 2025 --subject 物理类 --batch 本科批 --rank-center 50000 --rank-width 20000`
   - `python scripts/query_admissions.py --source major-scores --year 2025 --subject 物理类 --major-contains 计算机`
   - `python scripts/query_admissions.py --source plans --year 2025 --subject 物理类 --school-exact 中山大学`
   - `python scripts/query_admissions.py --source plans --year 2025 --subject 物理类 --school-code 10614`
   - `python scripts/query_admissions.py --source charters --school-exact 清华大学`

5. **生成推荐清单供用户审阅。**  
   读取 `references/recommendation-workflow.md`，按冲、稳、保、垫、兜底分层输出。每条推荐包含院校名称、院校代码、院校专业组、批次、科类、目标专业、组内可接受专业、历史最低分/位次、招生计划变化、地域、学校性质、学费风险、推荐理由、风险提示、章程核查状态、是否建议服从调剂。

6. **等待用户确认推荐方案。**  
   在用户确认前，不要生成最终志愿表草稿。询问用户是否删除院校、调整风险比例、加入新偏好、排除高收费项目、改变地域或专业约束、改变是否服从调剂。

7. **最终交叉复核。**  
   读取 `references/final-form-guide.md`，并回看资料卡、`references/2026-policy-rules.md`、`references/data-sources.md`、`references/recommendation-workflow.md`。逐项核对批次、填报时段、院校代码、院校专业组代码、专业代码、选科要求、体检限制、单科/语种限制、收费、调剂范围、艺术/体育/普通类同批次兼报冲突。

8. **草拟官方志愿表。**  
   使用 `assets/2026年广东省普通高校招生考生志愿表（官方）.pdf` 作为官方字段依据，使用 `assets/2026广东高考志愿表填报模板.xlsx` 作为可编辑 Excel 空白模板。最终 `.xlsx` 草稿必须由脚本复制并填充该模板；不要让 Agent 从零生成 Excel 文件。普通类本科/专科最多 45 个志愿时，用 `--pad-to 45` 补齐可编辑空白志愿位：
   `python scripts/render_volunteer_draft.py --input recommendations.json --template assets/2026广东高考志愿表填报模板.xlsx --output output/volunteer-draft.xlsx --pad-to 45`
   明确提醒用户：草稿不等于提交，最终必须登录广东省教育考试院志愿填报系统完成网上填报与确认。

## 输出要求

推荐清单必须包含：

- 基于资料卡的考生摘要
- 使用的数据 source、年份和查询条件
- 冲、稳、保、垫、兜底分层表格
- 每层风险说明
- 每条志愿的章程、体检、专业组调剂和费用风险
- 需要用户确认的问题
- 需要以 2026 最新官方招生专业目录或高校章程再次确认的事项

最终志愿表草稿优先输出 `.xlsx` 文件，且必须基于 `assets/2026广东高考志愿表填报模板.xlsx` 复制填充。草稿必须包含：

- 官方表格字段：录取批次、志愿号、院校代码、院校名称、院校专业组代码、专业代码 1-6、是否服从调剂
- 每个志愿的层级标签和风险备注
- 官方系统填报与确认提醒
- 普通类本科/专科 45 个志愿位的可编辑空白行或已确认志愿
- “以官方系统、最新招生专业目录和高校招生章程为准”的提示

## 脚本清单

- `scripts/create_profile_card.py`：根据 JSON 生成考生资料卡 Markdown。
- `scripts/equivalent_score.py`：查询广东一分一段表和历史同位次分数。
- `scripts/query_admissions.py`：查询院校分数、专业分数、招生计划、招生章程、艺术类和春考数据。
- `scripts/render_volunteer_draft.py`：把用户确认后的推荐结果渲染成官方表格字段风格的 Markdown，或复制 `assets/2026广东高考志愿表填报模板.xlsx` 并填充为可编辑 Excel 志愿草稿。

## 禁止事项

- 不要在信息不足时直接推荐院校。
- 不要让资料卡遗漏用户补充约束；如果字段无法归入固定栏目，也必须保留在资料卡的“原始补充信息”中。
- 不要在用户同时给出分数和位次但二者与历史一分一段不一致时只按分数判断；广东志愿推荐优先使用用户明确给出的 2026 位次，分数只作辅助。
- 不要把 2022-2025 历史数据当作 2026 最终招生专业目录。
- 不要把历史专业代码、专业组代码或招生计划直接当作 2026 最终填报代码；必须标记并等待 2026 最新招生专业目录复核。
- 不要只按分数推荐；必须结合位次、等位分、专业组和计划变化。
- 不要忽略院校专业组内调剂风险。
- 不要把用户明确排斥的专业、地域或无法承担的高收费项目放入保、垫、兜底层。
- 不要跳过招生章程、体检、单科、语种、性别、色觉、身高、收费限制核查。
- 不要把 `major-scores` 查询不到结果解读为“该校没有该专业”；必须改查 `plans` 补查专业组、专业代码、招生人数、学费和备注，并在推荐清单中标记“专业分缺失”。
- 不要用 `--school-contains` 的结果直接做最终推荐或填表；遇到同名/包含关系院校时必须改用 `--school-exact` 或 `--school-code`，例如“电子科技大学”不能混入杭州电子科技大学、桂林电子科技大学等。
- 不要在未运行 `--inspect` 或未确认字段口径时手工猜 Excel 表头；`charters` 等非标准表必须依赖脚本自动表头识别并用精确院校条件核查。
- 不要在 PowerShell 中裸写带括号的 `--columns` 参数；包含 `学制(年)`、`学费(元)` 等列名时必须给整个列清单加引号。
- 不要在用户确认推荐清单前生成最终志愿表草稿。
- 不要绕过 `assets/2026广东高考志愿表填报模板.xlsx` 直接从零生成最终 Excel 草稿。
- 不要把最终 Excel 输出路径设为 `assets/2026广东高考志愿表填报模板.xlsx`；模板只能作为输入复制填充，最终文件必须写入 `output/`。
- 不要覆盖正在被 Excel 或预览程序占用的输出文件；如果出现 `~$` 锁文件或写入被拒绝，应另存新文件并告知用户。
- 不要输出少于官方志愿位数量的普通类本科/专科最终草稿；确认后用 `--pad-to 45` 补齐可编辑空白志愿位。
- 不要只输出 Markdown 作为最终交付；最终草稿优先提供基于模板填充的可编辑 `.xlsx`，Markdown 只作为预览或附属说明。
- 不要在志愿表中凭经验补空白专业代码；缺失代码必须留空并在备注或提醒页标记“待 2026 目录补齐”。
- 不要声称草稿已经提交；只有广东官方系统网上确认后的志愿才有效。
