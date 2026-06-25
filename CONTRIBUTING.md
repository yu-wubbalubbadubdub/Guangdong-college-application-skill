# Contributing

欢迎改进 Guangdong College Application Skill。

## 提交前检查

- 不要提交 `__pycache__/`、`.pyc`、`output/` 等运行产物。
- 不要提交真实考生个人信息（姓名、分数、位次、身份证号等）。
- 修改 `SKILL.md` 时，保持 workflow 清晰，不要把大量示例堆进主文件；复杂说明放到 `references/`。
- 修改脚本后，至少运行一次对应脚本的 `--help` 或最小 smoke test。
- 如果改动了依赖，更新 `requirements.txt`。
- 数据文件更新时，注意保持 CSV/Excel 的编码与列结构一致。

## Pull Request 建议

请说明：

- 改了什么；
- 为什么需要改；
- 是否影响数据查询逻辑或志愿表渲染格式；
- 是否引入新的依赖或安全风险。

## 数据版权

历史录取数据和招生计划数据来源于广东省教育考试院公开发布的材料，仅用于辅助分析。请勿将未授权的第三方数据直接提交到仓库中。
