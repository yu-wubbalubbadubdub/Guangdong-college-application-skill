# 数据源目录与字段说明

本文件说明 `<skill_dir>/data/` 中的数据源。脚本默认读取这些 skill 内部文件，不依赖项目根目录下的原始资料包。`<skill_dir>/data/` 是脚本私有数据目录，不是按需加载文档目录。

## 目录

- [数据目录结构](#数据目录结构)
- [脚本 Source 清单](#脚本-source-清单)
- [一分一段数据](#一分一段数据)
- [主表字段](#主表字段)
- [查询参数补充](#查询参数补充)
- [查询策略](#查询策略)
- [数据使用优先级](#数据使用优先级)
- [风险提示](#风险提示)

使用原则：

- 先运行 `python <skill_dir>/scripts/query_admissions.py --list-sources` 查看 source 名。
- 查询新表前先运行 `python <skill_dir>/scripts/query_admissions.py --source <source> --inspect` 查看 sheet 和字段。
- 不要凭记忆改字段名；字段以脚本 inspect 结果为准。
- 精确匹配院校时优先使用 `--school-code` 或 `--school-exact`，避免 `--school-contains` 误匹配同名/包含关系院校。
- 历史数据只用于推荐分析，最终填报必须以 2026 官方招生专业目录、广东省教育考试院更正公告和高校招生章程为准。
- `spring` source 默认读取 `院校录取信息表`，如需查询隐藏历史投档 sheet，必须显式使用 `--sheet` 指定。

## 数据目录结构

```text
<skill_dir>/data/
├── scores/
│   ├── school-scores-2022-2025.xlsx
│   └── major-scores-2022-2025.xlsx
├── plans/
│   └── admission-plans-2022-2025.xlsx
├── segments/
│   ├── guangdong-segments-2022.xlsx
│   ├── guangdong-segments-2023.xlsx
│   ├── guangdong-segments-2024.xlsx
│   └── guangdong-segments-2025.xlsx
├── charters/
│   └── charters-2025.xlsx
├── art/
│   ├── art-school-scores.xlsx
│   ├── art-major-scores.xlsx
│   └── art-plans.xlsx
└── spring/
    └── spring-2026.xlsx
```

## 脚本 Source 清单

| source | skill 内部文件 | 原始资料来源 | 用途 |
| --- | --- | --- | --- |
| `school-scores` | `<skill_dir>/data/scores/school-scores-2022-2025.xlsx` | `22-25年全国高校在广东的院校录取分数.xlsx` | 院校专业组层面的最低分、最低位次、批次线差 |
| `major-scores` | `<skill_dir>/data/scores/major-scores-2022-2025.xlsx` | `22-25年全国高校在广东的专业录取分数.xlsx` | 专业层面的最低分、最低位次、专业备注、选科要求 |
| `plans` | `<skill_dir>/data/plans/admission-plans-2022-2025.xlsx` | `22-25年全国高校在广东的招生计划.xlsx` | 2022-2025 在粤招生计划、专业组、招生人数、学费 |
| `charters` | `<skill_dir>/data/charters/charters-2025.xlsx` | `最新25年招生章程链接.xlsx` | 2025 招生章程链接、高校主页 |
| `art-school-scores` | `<skill_dir>/data/art/art-school-scores.xlsx` | `全国艺术院校在广东的院校录取分数线.xlsx` | 艺术类院校层投档线、最低位次、计算公式 |
| `art-major-scores` | `<skill_dir>/data/art/art-major-scores.xlsx` | `全国艺术院校在广东的专业录取分数线.xlsx` | 艺术类专业层最低分、控线、计算公式 |
| `art-plans` | `<skill_dir>/data/art/art-plans.xlsx` | `全国艺术院校在广东的招生计划.xlsx` | 艺术类招生计划、专业方向、计划人数 |
| `spring` | `<skill_dir>/data/spring/spring-2026.xlsx` | `广东-2026春考学考3+证书.xlsx` | 春考学考与 3+证书投档、模板、指南 |

## 一分一段数据

`<skill_dir>/scripts/equivalent_score.py` 读取：

| 年份 | 文件 |
| --- | --- |
| 2022 | `<skill_dir>/data/segments/guangdong-segments-2022.xlsx` |
| 2023 | `<skill_dir>/data/segments/guangdong-segments-2023.xlsx` |
| 2024 | `<skill_dir>/data/segments/guangdong-segments-2024.xlsx` |
| 2025 | `<skill_dir>/data/segments/guangdong-segments-2025.xlsx` |

字段：

- 年份
- 科类
- 批次
- 控制线(分)
- 分数(分)
- 本段人数(人)
- 累计人数(人)
- 排名区间
- 历史同位次考生得分

示例：

```bash
python <skill_dir>/scripts/equivalent_score.py --year 2025 --subject 物理类 --rank 50000
python <skill_dir>/scripts/equivalent_score.py --year 2025 --subject 历史类 --score 560
```

## 主表字段

`school-scores` 常用字段：

- 年份、院校名称、院校代码、科类、批次、招生类型、专业组、选科要求、录取人数、最低分数、最低分位、批次线差、学校所在、学校性质、是否985、是否211

`major-scores` 常用字段：

- 年份、院校名称、院校代码、科类、批次、专业、专业代码、所属专业组、专业备注、选科要求、录取人数、最低分数、最低位次、学校所在、学校性质、是否985、是否211

`plans` 常用字段：

- 年份、院校名称、院校代码、科类、批次、招生类型、专业名称、专业代码、所属专业组、专业备注、选科要求、招生人数、学制(年)、学费(元)

`charters` 常用字段：

- 院校名称、所在省、主管部门、院校层次、2025招生章程、院校主页链接、logo地址

说明：`charters` 原表首行是标题，脚本会自动识别真正表头。查询示例：

```bash
python <skill_dir>/scripts/query_admissions.py --source charters --school-exact 清华大学
```

艺术类和春考数据字段差异较大，使用前必须先 `--inspect`。
脚本已为 `spring` 和 `art-*` source 配置专用默认输出列，但艺术类数据可能不覆盖所有年份；若按年份查询为空，必须先核对可用年份，并以 2026 体育艺术版招生专业目录和高校章程复核。

## 查询参数补充

推荐使用：

```bash
python <skill_dir>/scripts/query_admissions.py --source plans --year 2025 --subject 物理类 --school-code 10614
python <skill_dir>/scripts/query_admissions.py --source plans --year 2025 --subject 物理类 --school-exact 电子科技大学
python <skill_dir>/scripts/query_admissions.py --source major-scores --year 2025 --subject 物理类 --major-exact 计算机类
```

Windows PowerShell 中，如果 `--columns` 包含括号字段，必须给整个列名参数加引号：

```powershell
python -X utf8 <skill_dir>\scripts\query_admissions.py --source plans --year 2025 --subject 物理类 --school-code 10614 --columns "年份,院校名称,院校代码,科类,批次,专业名称,专业代码,所属专业组,招生人数,学制(年),学费(元)"
```

## 查询策略

1. 用一分一段表定位考生当前年份分数、位次和历史同位次分数。
2. 用 `school-scores` 粗筛院校专业组，优先看 2025，再回看 2024、2023、2022 稳定性。
3. 用 `major-scores` 核查目标专业及同组其他专业最低位次，避免只看院校专业组投档线。
4. 如果 `major-scores` 对目标院校没有返回专业层结果，标记“专业分缺失”，再用 `plans` 补查招生计划、专业组、专业代码、人数、学费和备注。
5. 用 `charters` 找章程入口，再到高校官网或教育部阳光高考平台确认录取规则、体检限制、单科/语种要求、收费。
6. 艺术类、体育类、春考、学考、3+证书考生先确认类别和合成分规则，再查询对应数据。

## 数据使用优先级

1. 2026 广东省教育考试院官方通知、招生专业目录、增补更正公告和高校招生章程。
2. `<skill_dir>/data/` 中 2025 年录取、计划和一分一段数据。
3. 2024、2023、2022 年数据，用于趋势、冷热变化和等位分稳定性校验。
4. 原始资料包中的旧版历史数据只能作为补充趋势，不直接替代本 skill 整理后的主数据源。

## 风险提示

- 2026 招生计划和专业目录以官方最新发布为准，22-25 数据只代表历史。
- 同一院校不同专业组风险差异很大，不要只按院校名称推荐。
- 专业最低分可能高于专业组投档线；调剂风险必须看专业组内全部专业和章程。
- 部分高收费项目、中外合作、联合培养、学分互认、国际班、民办和独立学院必须结合家庭经济情况确认。
- 某些历史表的批次、科类命名从 2021 新高考后发生变化，比较时先统一物理类/历史类口径。
- 资料中关于专科线可能存在口径差异；遇到分数线判断时以最新官方批次线为准。
