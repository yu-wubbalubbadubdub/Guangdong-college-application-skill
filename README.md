# 广东高考志愿填报 Skill

面向广东省 2026 届高考志愿填报的 Codex Skill。它可以采集考生资料，读取广东 2026 志愿填报政策，查询整理后的 2022-2025 年录取分数、招生计划和一分一段数据，生成“冲、稳、保、垫、兜底”推荐清单，并基于官方模板输出可编辑 Excel 志愿表草稿。

## 目录结构

```text
.
├── SKILL.md
├── assets/
│   ├── 2026广东高考志愿表填报模板.xlsx
│   └── 2026年广东省普通高校招生考生志愿表（官方）.pdf
├── references/
│   ├── 2026-policy-rules.md
│   ├── data-sources.md
│   ├── final-form-guide.md
│   ├── profile-card-guide.md
│   ├── recommendation-workflow.md
│   └── data/
└── scripts/
    ├── create_profile_card.py
    ├── equivalent_score.py
    ├── query_admissions.py
    └── render_volunteer_draft.py
```

## 使用方式

在 Codex 中触发本 skill 后，从 `SKILL.md` 开始执行主流程。核心流程是：

1. 采集考生必填信息和偏好补充信息。
2. 生成本地考生资料卡。
3. 读取政策、数据源和推荐工作流说明。
4. 使用 `scripts/` 中的脚本查询数据，不直接手工解析 Excel。
5. 先输出推荐清单供用户确认。
6. 用户确认后，基于 `assets/2026广东高考志愿表填报模板.xlsx` 填充可编辑 Excel 志愿表草稿。

## 常用脚本

```powershell
python scripts/query_admissions.py --list-sources
python scripts/query_admissions.py --source school-scores --inspect
python scripts/equivalent_score.py --year 2025 --subject 物理类 --rank 50000
python scripts/render_volunteer_draft.py --input recommendations.json --template assets/2026广东高考志愿表填报模板.xlsx --output output/volunteer-draft.xlsx --pad-to 45
```

## 重要提醒

- 本 skill 的历史数据仅用于辅助分析，正式填报必须以广东省教育考试院 2026 最新招生专业目录、增补更正公告和高校招生章程为准。
- 最终志愿必须在广东省教育考试院志愿填报系统中完成网上填报与确认。
- `references/data/` 下的大型 `.xlsx` 数据表应通过脚本访问，不要直接读写。
