#!/usr/bin/env python3
"""Create or update a structured Guangdong application profile card."""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


SECTIONS = [
    ("基本信息", ["考生", "年份", "科类", "分数", "位次", "目标批次", "选科/专业统考类别"]),
    ("志愿目标", ["首要目标", "次要目标", "明确不接受"]),
    ("院校偏好", ["倾向", "可接受", "排斥"]),
    ("专业偏好", ["倾向", "可接受", "排斥", "特别说明"]),
    ("地域偏好", ["优先地区", "可接受地区", "排斥地区"]),
    ("经济与生活约束", ["年学费上限", "住宿/城市成本偏好", "是否接受高收费项目"]),
    ("风险偏好", ["冲刺意愿", "稳妥程度", "是否服从调剂"]),
    ("待核对事项", ["体检限制", "单科/语种限制", "特殊资格", "仍需用户补充"]),
    ("推荐使用准则", ["推荐时必须优先满足", "可在冲刺层尝试", "任何层级都不得推荐"]),
]


ALIASES = {
    "考生": ["name", "student", "姓名", "考生"],
    "年份": ["year", "年份"],
    "科类": ["subject", "科类"],
    "分数": ["score", "分数"],
    "位次": ["rank", "位次", "排位"],
    "目标批次": ["target_batch", "目标批次"],
    "选科/专业统考类别": ["selected_subjects", "selection", "选科", "选科/专业统考类别", "统考类别"],
}

SECTION_ALIASES = {
    ("志愿目标", "首要目标"): ["primary_goal", "首要目标"],
    ("志愿目标", "次要目标"): ["secondary_goal", "次要目标"],
    ("志愿目标", "明确不接受"): ["explicit_rejections", "明确不接受"],
    ("院校偏好", "倾向"): ["preferred_schools", "倾向院校", "院校倾向"],
    ("院校偏好", "可接受"): ["acceptable_schools", "可接受院校"],
    ("院校偏好", "排斥"): ["rejected_schools", "排斥院校"],
    ("专业偏好", "倾向"): ["preferred_majors", "倾向专业"],
    ("专业偏好", "可接受"): ["acceptable_majors", "可接受专业"],
    ("专业偏好", "排斥"): ["rejected_majors", "排斥专业"],
    ("专业偏好", "特别说明"): ["interests", "extra_constraints", "特别说明"],
    ("地域偏好", "优先地区"): ["preferred_regions", "倾向地区", "优先地区"],
    ("地域偏好", "可接受地区"): ["acceptable_regions", "可接受地区"],
    ("地域偏好", "排斥地区"): ["rejected_regions", "明确排斥地区", "排斥地区"],
    ("经济与生活约束", "年学费上限"): ["budget", "tuition_limit", "年学费上限"],
    ("经济与生活约束", "住宿/城市成本偏好"): ["living_cost_preference", "住宿/城市成本偏好"],
    ("经济与生活约束", "是否接受高收费项目"): ["high_fee_acceptance", "是否接受高收费项目"],
    ("风险偏好", "冲刺意愿"): ["risk_preference", "冲刺意愿"],
    ("风险偏好", "稳妥程度"): ["stability_preference", "稳妥程度"],
    ("风险偏好", "是否服从调剂"): ["adjustment_preference", "是否服从调剂"],
    ("待核对事项", "体检限制"): ["body_constraints", "体检限制"],
    ("待核对事项", "单科/语种限制"): ["single_subject_or_language_constraints", "单科/语种限制"],
    ("待核对事项", "特殊资格"): ["special_qualification", "特殊资格"],
    ("待核对事项", "仍需用户补充"): ["missing_info", "career_goals", "仍需用户补充"],
    ("推荐使用准则", "推荐时必须优先满足"): ["primary_goal", "must_have", "推荐时必须优先满足"],
    ("推荐使用准则", "可在冲刺层尝试"): ["stretch_options", "可在冲刺层尝试"],
    ("推荐使用准则", "任何层级都不得推荐"): ["explicit_rejections", "never_recommend", "任何层级都不得推荐"],
}

STRICT_REQUIRED_FIELDS = [
    ("基本信息", "考生"),
    ("基本信息", "年份"),
    ("基本信息", "科类"),
    ("基本信息", "分数"),
    ("基本信息", "位次"),
    ("基本信息", "目标批次"),
    ("基本信息", "选科/专业统考类别"),
    ("风险偏好", "是否服从调剂"),
    ("待核对事项", "体检限制"),
    ("待核对事项", "单科/语种限制"),
]


def known_input_keys() -> set[str]:
    keys = set(ALIASES)
    keys.update(alias for aliases in ALIASES.values() for alias in aliases)
    keys.update(section for section, _ in SECTIONS)
    for (section, item), aliases in SECTION_ALIASES.items():
        keys.add(section)
        keys.add(item)
        keys.update(aliases)
    return keys


def slug(value: str) -> str:
    text = re.sub(r"[\\/:*?\"<>|]+", "-", value.strip())
    text = re.sub(r"\s+", "-", text)
    return text or "anonymous"


def lookup(data: dict, key: str) -> str:
    return lookup_aliases(data, ALIASES.get(key, [key]))


def lookup_aliases(data: dict, aliases: list[str]) -> str:
    for alias in aliases:
        if alias in data and data[alias] not in (None, ""):
            return str(data[alias])
    return ""


def section_value(data: dict, section: str, item: str) -> str:
    nested = data.get(section)
    if isinstance(nested, dict) and item in nested:
        return str(nested[item])
    if (section, item) in SECTION_ALIASES:
        return lookup_aliases(data, SECTION_ALIASES[(section, item)])
    return lookup(data, item)


def render(data: dict) -> str:
    lines = ["# 2026广东高考志愿填报用户资料卡", ""]
    lines.append(f"- 创建时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("- 使用说明：后续推荐、复核和填表必须围绕本资料卡进行。")
    lines.append("")
    for section, items in SECTIONS:
        lines.append(f"## {section}")
        for item in items:
            value = section_value(data, section, item)
            lines.append(f"- {item}：{value}")
        lines.append("")
    extras = [(key, value) for key, value in data.items() if key not in known_input_keys() and value not in (None, "", [], {})]
    if extras:
        lines.append("## 原始补充信息")
        lines.append("- 以下字段未被固定栏目归档，后续推荐、查询和复核仍必须参考。")
        for key, value in extras:
            lines.append(f"- {key}：{value}")
        lines.append("")
    return "\n".join(lines)


def strict_missing(data: dict) -> list[str]:
    missing = []
    for section, item in STRICT_REQUIRED_FIELDS:
        if not section_value(data, section, item):
            missing.append(f"{section}/{item}")
    return missing


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, help="JSON file with profile data")
    parser.add_argument("--output", required=True, help="Output markdown path under the task output directory")
    parser.add_argument("--strict", action="store_true", help="Fail when core recommendation fields are missing")
    args = parser.parse_args()

    data = json.loads(Path(args.input).read_text(encoding="utf-8-sig"))
    if args.strict:
        missing = strict_missing(data)
        if missing:
            raise SystemExit("资料卡严格校验未通过，缺少核心字段：\n" + "\n".join(f"- {item}" for item in missing))
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(render(data), encoding="utf-8")
    print(output)


if __name__ == "__main__":
    main()
