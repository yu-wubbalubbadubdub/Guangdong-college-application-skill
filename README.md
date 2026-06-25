# 广东高考志愿填报 Skill

面向广东省 2026 届高考志愿填报的 Codex Skill。它支持普通类本科/专科考生的端到端辅助流程：采集考生资料、生成资料卡、查询 2022-2025 历史录取与招生计划数据、生成统一格式推荐清单、渲染可编辑 Excel 志愿表草稿，并运行最终交叉验证。

范围边界：普通类本科/专科支持端到端自动化；艺术、体育、春考、学考、3+证书和其他特殊类别仅提供政策与数据查询辅助，最终推荐需要人工按对应规则补充复核。

## 目录结构

```text
Guangdong-college-application-skill
├── skills/
│   └── guangdong-gaokao-application/
│       ├── SKILL.md
│       ├── assets/
│       ├── data/
│       ├── references/
│       └── scripts/
├── README.md
├── LICENSE
├── requirements.security
└── .gitignore
```

## 安装

PowerShell:

```powershell
New-Item -ItemType Directory -Force "$env:USERPROFILE\.codex\skills" | Out-Null
Copy-Item -Recurse -Force ".\skills\guangdong-gaokao-application" "$env:USERPROFILE\.codex\skills\guangdong-gaokao-application"
```

Bash:

```bash
mkdir -p ~/.codex/skills
ln -s "$(pwd)/skills/guangdong-gaokao-application" ~/.codex/skills/guangdong-gaokao-application
```

## 常用命令

以下命令默认从仓库根目录执行：

```powershell
python skills/guangdong-gaokao-application/scripts/query_admissions.py --list-sources
python skills/guangdong-gaokao-application/scripts/query_admissions.py --source school-scores --inspect
python skills/guangdong-gaokao-application/scripts/equivalent_score.py --year 2025 --subject 物理类 --rank 50000
python skills/guangdong-gaokao-application/scripts/build_recommendations.py --candidate output/zhangsan-物理类-50000/candidate.json --output-dir output/zhangsan-物理类-50000
python skills/guangdong-gaokao-application/scripts/render_volunteer_draft.py --input output/zhangsan-物理类-50000/recommendations-confirmed.json --template skills/guangdong-gaokao-application/assets/2026广东高考志愿表填报模板.xlsx --output output/zhangsan-物理类-50000/volunteer-draft.xlsx --pad-to 45 --strict-final
python skills/guangdong-gaokao-application/scripts/validate_volunteer_output.py --file output/zhangsan-物理类-50000/volunteer-draft.xlsx --batch 普通类本科院校
```

## 产物路径

所有运行产物写入仓库根目录的 `output/{考生代号}-{科类}-{位次}/`，不要写入 skill 包内部。典型产物包括：

- `candidate.json`
- `profile-card.md`
- `recommendations-review.md`
- `intermediate/recommendations-proposed.json`
- `recommendations-confirmed.json`
- `volunteer-draft.xlsx`
- `validation-report.json`
- `completion-summary.md`

## 数据和安全

- 本仓库脚本运行时只需要 Python 标准库；`requirements.security` 仅用于维护者执行可选安全扫描。
- 历史录取数据和招生计划数据仅用于辅助分析，不能替代广东省教育考试院 2026 最新招生专业目录、增补更正公告、高校招生章程和官方志愿系统规则。
- 官方 PDF、Excel 模板、历史数据表等材料应遵循其原始来源的权利和使用限制；本仓库许可证主要覆盖本仓库原创脚本、skill 指令和配套文档。
- 考生姓名、分数、位次、偏好和限制属于敏感个人信息；请只在本地 `output/` 下保存任务产物，不要提交真实考生资料。

## 许可证

本仓库原创代码与文档采用 MIT License，详见 `LICENSE`。
