---
name: guangdong-gaokao-application
description: >-
  当用户需要为广东省2026届高考生进行志愿填报、普通类本科/专科院校专业推荐、冲稳保垫兜底分析、广东高考数据查询、生成广东省普通高校招生考生志愿表草稿时使用。支持普通类本科/专科端到端自动推荐与Excel草稿生成；艺术、体育、春考、学考、3+证书和其他特殊类别仅提供政策与数据查询辅助，最终推荐需人工按对应规则补充复核。
---

# 广东 2026 高考志愿填报

把本文件作为流程入口。`<skill_dir>` 表示本 `SKILL.md` 所在目录；不要创建旧版中间目录。所有命令默认从项目根目录执行，脚本路径写作 `python <skill_dir>/scripts/...`，运行产物写入项目根目录 `output/...`。执行具体任务时，按需读取 `<skill_dir>/references/` 中的指引文件，优先使用 `<skill_dir>/scripts/` 查询结构化数据，使用 `<skill_dir>/assets/` 中的官方 PDF 作为字段依据，并使用 `<skill_dir>/assets/2026广东高考志愿表填报模板.xlsx` 作为最终可编辑 Excel 草稿的空白模板。

## 参考文件导航

- `<skill_dir>/references/profile-card-guide.md`：考生资料卡的字段、结构和写法。
- `<skill_dir>/references/2026-policy-rules.md`：广东省 2026 年志愿填报时间、批次、确认规则、特殊类别和政策风险。
- `<skill_dir>/references/data-sources.md`：`<skill_dir>/data/` 内所有数据工作簿的目录、字段、脚本 source 名和使用优先级。
- `<skill_dir>/references/recommendation-workflow.md`：冲、稳、保、垫、兜底分层方法、推荐清单字段和风险表达方式。
- `<skill_dir>/references/final-form-guide.md`：最终志愿表草稿的字段顺序、生成命令和官方系统提醒。
- `<skill_dir>/references/final-cross-validation-guide.md`：最终交叉验证清单、验证报告结构和失败回退规则。
- `<skill_dir>/references/completion-summary-guide.md`：最终交叉验证通过后，面向用户的结构化任务完成摘要写法。

`<skill_dir>/data/` 是本技能自带的数据源目录，脚本默认从这里读取；该目录是脚本私有数据目录，不属于按需加载文档：

- `<skill_dir>/data/scores/`：22-25 年院校录取分数、专业录取分数。
- `<skill_dir>/data/plans/`：22-25 年在粤招生计划。
- `<skill_dir>/data/segments/`：2022-2025 年广东一分一段表。
- `<skill_dir>/data/charters/`：2025 年高校招生章程链接表。
- `<skill_dir>/data/art/`：艺术类院校/专业录取分数与招生计划。
- `<skill_dir>/data/spring/`：2026 春考、学考、3+证书数据。

`<skill_dir>/data/` 下的 `.xlsx` 大表只能通过 `<skill_dir>/scripts/query_admissions.py`、`<skill_dir>/scripts/equivalent_score.py` 或封装这些脚本能力的 `<skill_dir>/scripts/build_recommendations.py` 访问。不要直接读写这些二进制数据表，不要手工猜字段。

性能说明：`query_admissions.py` 默认会把解析后的工作簿行缓存到项目根目录 `.cache/guangdong-gaokao/`。第一次读取大表仍需解析 `.xlsx` 内部 XML；后续同一工作簿、同一 sheet 的点查和批量查询应复用缓存。数据源文件路径、大小、修改时间或缓存版本变化时会自动失效；怀疑缓存异常或更新数据后，使用 `--rebuild-cache` 重建，排障时可用 `--no-cache`。

`<skill_dir>/assets/` 存放最终填表相关资产：

- `<skill_dir>/assets/2026年广东省普通高校招生考生志愿表（官方）.pdf`：官方附件原表，用于核对字段、批次和表格样式。
- `<skill_dir>/assets/2026广东高考志愿表填报模板.xlsx`：可编辑空白 Excel 志愿表模板，最终 `.xlsx` 草稿必须复制并填充此模板，不要从零生成 Excel 文件。

## 产物目录与命名

所有运行产物统一写入项目根目录的 `output/{考生代号}-{科类}-{位次}/`，不要写入 `<skill_dir>/`、`<skill_dir>/output/`、`<skill_dir>/assets/` 或 `<skill_dir>/references/`。命令默认从项目根目录执行，输出路径直接使用 `output/{考生代号}-{科类}-{位次}/...`。

单次任务至少保留以下文件名：

- `candidate.json`：用户原始信息和结构化输入。
- `profile-card.md`：考生资料卡。
- `recommendations-review.md`：给用户审阅的推荐清单。
- `recommendations-confirmed.json`：用户确认后、用于渲染志愿表的结构化推荐数据。
- `volunteer-draft.xlsx`：基于官方模板填充的可编辑志愿表草稿。
- `validation-report.md` 或 `validation-report.json`：最终交叉验证结果。
- `completion-summary.md`：最终交叉验证通过后发送给用户的任务完成摘要。

如需保留中间查询结果，写入同一目录下的 `query-log.md` 或 `intermediate/` 子目录。不要覆盖同名历史产物；同一考生重复运行时，在目录名后追加日期或轮次，例如 `chenyu-物理类-50000-20260625-r2`。

## 主流程

任一步出现脚本错误、核心字段缺失、目标院校无法精确匹配、专业代码缺失、招生章程无法核查、用户偏好冲突或政策限制未确认时，先停止当前阶段，标记原因并补查或追问。不要让错误或缺失静默进入下一阶段。

1. **先采集考生信息。**  
   先收集必填信息：姓名或昵称、高考分数、全省位次、科类（物理类/历史类）、目标批次、选科组合、是否为艺术类/体育类/春考/学考/3+证书考生、是否服从调剂。缺少这些信息时不要进入正式推荐。随后收集偏好补充项：倾向院校、倾向专业、明确排斥专业、倾向地域、明确排斥地域、家庭经济承受能力、是否接受高收费项目、个人兴趣、职业规划、体检限制、单科成绩限制、外语语种限制、特殊资格，以及用户主动补充的其他择校择专业约束。采集完成后，把原始信息结构化保存为 `output/{考生代号}-{科类}-{位次}/candidate.json`，第 2 步只读取该文件。

2. **生成或更新考生资料卡。**  
   读取 `<skill_dir>/references/profile-card-guide.md`，把已采集信息整理成结构化资料卡，并保存到项目根目录 `output/{考生代号}-{科类}-{位次}/profile-card.md`。如果用户后续补充或修正偏好，先更新资料卡，再继续查询或推荐。可使用：
   `python <skill_dir>/scripts/create_profile_card.py --input output/{考生代号}-{科类}-{位次}/candidate.json --output output/{考生代号}-{科类}-{位次}/profile-card.md --strict`

3. **加载政策和数据说明。**  
   推荐前读取 `<skill_dir>/references/2026-policy-rules.md` 和 `<skill_dir>/references/data-sources.md`。确认填报时段、批次类别、院校专业组规则、合格考要求、艺术/体育兼报限制、特殊类型填报位置、数据年份、字段口径和数据风险。

4. **用脚本查询数据。**  
   不要手工解析 Excel。先用 `python <skill_dir>/scripts/query_admissions.py --list-sources` 查看可用 source，再用 `--inspect` 查看字段。`query_admissions.py` 默认启用项目本地解析缓存；第一次读取 `major-scores`、`plans` 等大表较慢是正常现象，后续查询应复用 `.cache/guangdong-gaokao/`。如果要生成普通类本科/专科 45 个志愿或进行跨 source 交叉筛选，优先使用 `python <skill_dir>/scripts/build_recommendations.py --candidate output/{考生代号}-{科类}-{位次}/candidate.json --output-dir output/{考生代号}-{科类}-{位次}` 一次性构建候选池、推荐 JSON、审阅稿和查询日志。艺术、体育、春考、学考、3+证书和其他特殊类别暂不使用 `build_recommendations.py` 端到端自动推荐，应使用 `query_admissions.py` 查询对应数据后按政策人工复核；不要为每个考生临时重写一次性筛选脚本。
   常用查询：
   - `python <skill_dir>/scripts/equivalent_score.py --year 2025 --subject 物理类 --rank 50000`
   - `python <skill_dir>/scripts/query_admissions.py --source school-scores --year 2025 --subject 物理类 --batch 本科批 --rank-center 50000 --rank-width 20000`
   - `python <skill_dir>/scripts/query_admissions.py --source major-scores --year 2025 --subject 物理类 --major-contains 计算机`
   - `python <skill_dir>/scripts/query_admissions.py --source plans --year 2025 --subject 物理类 --school-exact 中山大学`
   - `python <skill_dir>/scripts/query_admissions.py --source plans --year 2025 --subject 物理类 --school-code 10614`
   - `python <skill_dir>/scripts/query_admissions.py --source charters --school-exact 清华大学`

5. **生成推荐清单供用户审阅。**  
   读取 `<skill_dir>/references/recommendation-workflow.md`，按冲、稳、保、垫、兜底分层输出，并保存到 `output/{考生代号}-{科类}-{位次}/recommendations-review.md`。普通类本科/专科推荐应优先由 `build_recommendations.py` 生成统一 schema 的 `intermediate/recommendations-proposed.json`，该 schema 必须与最终 Excel 字段对齐。默认采用“可接受范围内尽量填满”的覆盖策略：普通类本科/专科以 45 个官方志愿位为目标，军检/专项/艺体等批次以对应官方容量为目标。多数普通考生应先扩展合适候选池，尽量接近或填满目标批次容量，以增加录取机会；只有资料卡体现明确强院校/专业/地域偏好、高分强匹配策略、不可放宽约束，或脚本检索后合适数据不足时，才允许少量留空。每条推荐包含院校名称、院校代码、院校专业组、批次、科类、目标专业、组内可接受专业、历史最低分/位次、招生计划变化、地域、学校性质、学费风险、推荐理由、风险提示、章程核查状态、是否建议服从调剂。若推荐数少于官方容量，必须在推荐清单中写明未满额例外类型、已尝试的扩展查询和可放宽约束建议。

6. **等待用户确认推荐方案。**  
   在用户确认前，不要生成最终志愿表草稿。询问用户是否删除院校、调整风险比例、加入新偏好、排除高收费项目、改变地域或专业约束、改变是否服从调剂。用户确认后，从 `intermediate/recommendations-proposed.json` 中筛选、重排并重写 `志愿号`，把最终结构化推荐数据保存为 `output/{考生代号}-{科类}-{位次}/recommendations-confirmed.json`；必须保留与最终 Excel 对齐的统一字段，不要只保存 Markdown。

7. **草拟官方志愿表。**  
   读取 `<skill_dir>/references/final-form-guide.md`。使用 `<skill_dir>/assets/2026年广东省普通高校招生考生志愿表（官方）.pdf` 作为官方字段依据，使用 `<skill_dir>/assets/2026广东高考志愿表填报模板.xlsx` 作为可编辑 Excel 空白模板。最终 `.xlsx` 草稿必须由脚本复制并填充该模板；不要让 Agent 从零生成 Excel 文件。模板必须包含官方 PDF 中全部录取批次，允许保留“层级”和“备注”列。默认策略是输出“可接受范围内尽量填满”的最终草稿：已填志愿必须完整、可解释且符合资料卡；如果目标批次未接近或未达到官方容量，必须先回到推荐阶段扩展候选池，除非推荐清单已经给出强偏好、高分强匹配、不可放宽约束或检索数据不足等例外依据。只有当用户或策略明确要求满额时，才启用 `--require-full` 强制检查：
   `python <skill_dir>/scripts/render_volunteer_draft.py --input output/{考生代号}-{科类}-{位次}/recommendations-confirmed.json --template <skill_dir>/assets/2026广东高考志愿表填报模板.xlsx --output output/{考生代号}-{科类}-{位次}/volunteer-draft.xlsx --pad-to 45 --strict-final`

8. **最终交叉验证。**  
   在生成 `.xlsx` 后读取 `<skill_dir>/references/final-cross-validation-guide.md` 执行最终交叉验证。逐项核对批次、填报时段、院校代码、院校专业组代码、专业代码、选科要求、体检限制、单科/语种限制、收费、调剂范围、艺术/体育/普通类同批次兼报冲突、志愿顺序、覆盖率解释和输出文件位置。必须运行内容验证脚本，检查“实际完整填写数”而不是只检查行号：
   `python <skill_dir>/scripts/validate_volunteer_output.py --file output/{考生代号}-{科类}-{位次}/volunteer-draft.xlsx --batch 普通类本科院校`
   如需强制普通类本科/专科等目标批次满额，再追加 `--require-full` 到渲染脚本和验证脚本命令。把验证结果保存到 `output/{考生代号}-{科类}-{位次}/validation-report.md` 或 `.json`，格式遵循 `<skill_dir>/references/final-cross-validation-guide.md`。未满额不能只凭脚本 warning 通过：验证报告必须说明未满额例外类型、资料卡依据、已尝试扩展查询和是否建议用户放宽约束；如果这些依据不足，应直接回退到推荐阶段继续补充合适志愿。交叉验证发现任何问题时，不要先告知用户“已完成”；应直接回退到对应前置步骤补查、修正推荐清单、重新生成资料卡或重新渲染 Excel，直到最终交叉验证通过。

9. **发送任务完成摘要。**  
   最终交叉验证通过后，读取 `<skill_dir>/references/completion-summary-guide.md`，生成 `output/{考生代号}-{科类}-{位次}/completion-summary.md`，并向用户发送一份结构化任务完成摘要。摘要需要说明采用的填报策略、参考的数据源、志愿填报依据、具体填报建议、当前优势、主要风险、仍需用户在官方系统确认的事项和所有产物文件路径。明确提醒用户：草稿不等于提交，最终必须登录广东省教育考试院志愿填报系统完成网上填报与确认。

## 输出要求

推荐清单必须包含：

- 基于资料卡的考生摘要
- 使用的数据 source、年份和查询条件
- 冲、稳、保、垫、兜底分层表格
- 每层风险说明
- 每条志愿的章程、体检、专业组调剂和费用风险
- 需要用户确认的问题
- 需要以 2026 最新官方招生专业目录或高校章程再次确认的事项

最终志愿表草稿优先输出 `.xlsx` 文件，且必须基于 `<skill_dir>/assets/2026广东高考志愿表填报模板.xlsx` 复制填充。草稿必须包含：

- 官方表格字段：录取批次、志愿号、院校代码、院校名称、院校专业组代码、专业代码 1-6、是否服从调剂
- 每个志愿的层级标签和风险备注
- 官方系统填报与确认提醒
- 与官方 PDF 一致的全部录取批次志愿位
- 目标批次的覆盖统计：完整志愿数/官方容量；多数普通考生应在可接受范围内尽量接近或填满官方容量；未填满时说明例外类型、与资料卡相关的原因、已尝试扩展查询和可放宽约束建议
- “以官方系统、最新招生专业目录和高校招生章程为准”的提示

最终交叉验证通过后必须输出 `completion-summary.md`，并在回复用户时概括：

- 采用的总体填报策略和风险分层方式
- 使用的数据源、年份、脚本查询条件和官方材料
- 志愿表覆盖情况、未满额原因或满额校验结果
- 对用户的具体填报建议
- 当前方案优势、关键风险和仍需用户确认的事项
- `candidate.json`、`profile-card.md`、`recommendations-review.md`、`recommendations-confirmed.json`、`volunteer-draft.xlsx`、`validation-report` 和 `completion-summary.md` 的路径

## 脚本清单

- `<skill_dir>/scripts/create_profile_card.py`：根据 JSON 生成考生资料卡 Markdown。
- `<skill_dir>/scripts/equivalent_score.py`：查询广东一分一段表和历史同位次分数。
- `<skill_dir>/scripts/query_admissions.py`：查询院校分数、专业分数、招生计划、招生章程、艺术类和春考数据；默认启用 `.cache/guangdong-gaokao/` 解析缓存，支持 `--cache-dir`、`--no-cache`、`--rebuild-cache`。
- `<skill_dir>/scripts/build_recommendations.py`：读取 `candidate.json`，一次性交叉筛选 `school-scores`、`major-scores`、`plans`、`charters`，生成统一 schema 的推荐 JSON、审阅稿、候选池和查询日志；普通类本科/专科推荐优先使用它，避免重复编写一次性筛选脚本。
- `<skill_dir>/scripts/render_volunteer_draft.py`：把用户确认后的推荐结果渲染成官方表格字段风格的 Markdown，或复制 `<skill_dir>/assets/2026广东高考志愿表填报模板.xlsx` 并填充为可编辑 Excel 志愿草稿。
- `<skill_dir>/scripts/validate_volunteer_output.py`：读取最终 `.xlsx`，按目标批次检查实际完整填写数量、官方容量、部分填写项和覆盖率提示；仅在传入 `--require-full` 时把未满额作为失败。

## 禁止事项

- 不要在信息不足时直接推荐院校。
- 生成资料卡时，正式推荐前应使用 `<skill_dir>/scripts/create_profile_card.py --strict` 校验核心字段；严格校验失败时必须先追问用户，不要继续推荐。
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
- 不要在普通考生资料卡没有明确强偏好、不可放宽约束或数据不足证据时，只生成少量志愿并依靠“未满额 warning”通过；应回退到候选池构建，继续扩展合适院校专业组。
- 不要绕过 `<skill_dir>/assets/2026广东高考志愿表填报模板.xlsx` 直接从零生成最终 Excel 草稿。
- 不要把最终 Excel 输出路径设为 `<skill_dir>/assets/2026广东高考志愿表填报模板.xlsx`；模板只能作为输入复制填充，最终文件必须写入项目根目录 `output/{考生代号}-{科类}-{位次}/volunteer-draft.xlsx`。
- 不要把运行产物写入 `<skill_dir>/output/`；命令默认从项目根目录执行，输出路径必须使用 `output/{考生代号}-{科类}-{位次}/...`。
- 不要覆盖正在被 Excel 或预览程序占用的输出文件；如果出现 `~$` 锁文件或写入被拒绝，应另存新文件并告知用户。
- 不要为了凑满官方志愿位而加入违背资料卡的低匹配志愿，尤其是用户明确排斥的专业、地域、高收费项目或调剂风险不可接受的专业组。
- 不要把自动补齐的空白志愿位当作已填志愿；正式交付必须用 `--strict-final` 和 `<skill_dir>/scripts/validate_volunteer_output.py` 检查已填志愿完整性，并报告完整志愿数/官方容量。只有用户或策略明确要求满额时，才使用 `--require-full`。
- 不要只输出 Markdown 作为最终交付；最终草稿优先提供基于模板填充的可编辑 `.xlsx`，Markdown 只作为预览或附属说明。
- 不要在志愿表中凭经验补空白专业代码；缺失代码必须留空并在备注或提醒页标记“待 2026 目录补齐”。
- 不要在最终交叉验证失败时发送任务完成摘要；先回退修复并重新生成相关产物。
- 不要声称草稿已经提交；只有广东官方系统网上确认后的志愿才有效。
