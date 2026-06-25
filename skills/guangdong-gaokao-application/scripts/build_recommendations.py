#!/usr/bin/env python3
"""Build a reproducible Guangdong Gaokao recommendation review list.

The script turns a candidate.json/profile into a stable recommendation schema
that can later be confirmed and rendered by render_volunteer_draft.py.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import math
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
QUERY_SCRIPT = ROOT / "scripts" / "query_admissions.py"
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

HISTORY_YEAR = "2025"
SUBJECT_DEFAULT = "物理类"
SOURCE_BATCH_BY_TARGET = {
    "普通类本科院校": "本科批",
    "普通类本科": "本科批",
    "本科普通类": "本科批",
    "普通类专科院校": "专科批",
    "普通类专科": "专科批",
}
CAPACITY_BY_TARGET = {
    "普通类本科院校": 45,
    "普通类本科": 45,
    "本科普通类": 45,
    "普通类专科院校": 45,
    "普通类专科": 45,
}
LAYER_QUOTAS_45 = {"冲": 5, "稳": 12, "保": 14, "垫": 9, "兜底": 5}
OUTPUT_FIELDS = [
    "志愿号",
    "层级",
    "录取批次",
    "科类",
    "院校代码",
    "院校名称",
    "院校专业组代码",
    "专业代码列表",
    "专业代码1",
    "专业代码2",
    "专业代码3",
    "专业代码4",
    "专业代码5",
    "专业代码6",
    "是否服从调剂",
    "调剂条件",
    "推荐专业",
    "组内可接受专业",
    "2025专业最低分/位次",
    "2025组最低分/位次",
    "2024/2023/2022参考",
    "招生计划变化",
    "地域",
    "学校性质",
    "标签",
    "学费风险",
    "推荐理由",
    "风险提示",
    "章程核查状态",
    "代码复核状态",
    "备注",
]

DEFAULT_MAJOR_TERMS = [
    "计算机",
    "软件工程",
    "电子信息",
    "自动化",
    "网络空间安全",
    "数据科学",
    "通信工程",
    "物联网工程",
    "人工智能",
    "电气工程",
    "机械电子",
    "信息安全",
    "网络工程",
]
HIGH_FEE_TERMS = ["中外合作", "国际班", "合作办学", "内地澳门", "国际"]
REJECTION_EXPANSIONS = {
    "护理": ["护理"],
    "农学": ["农学", "农业", "农艺", "植物", "园艺", "动物", "水产", "林学", "草业"],
    "化学工程": ["化学工程", "化工"],
    "材料化学": ["材料化学"],
    "强化学实验": ["化学实验", "应用化学", "化学生物", "制药工程"],
    "高收费": HIGH_FEE_TERMS,
    "中外合作": HIGH_FEE_TERMS,
    "国际班": HIGH_FEE_TERMS,
}
REGION_EXPANSIONS = {
    "广东": {"广东"},
    "广州": {"广东"},
    "深圳": {"广东"},
    "珠三角": {"广东"},
    "上海": {"上海"},
    "杭州": {"浙江"},
    "长三角": {"上海", "江苏", "浙江", "安徽"},
    "省会城市": {"北京", "天津", "重庆", "湖北", "湖南", "江西", "福建", "四川", "山东", "广西", "海南", "河南", "河北"},
}
REMOTE_REGION_EXPANSIONS = {
    "东北偏远": {"黑龙江", "吉林"},
    "西北偏远": {"新疆", "甘肃", "宁夏", "青海", "内蒙古"},
}


def import_query_module():
    spec = importlib.util.spec_from_file_location("query_admissions", QUERY_SCRIPT)
    if spec is None or spec.loader is None:
        raise SystemExit(f"无法加载查询脚本：{QUERY_SCRIPT}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


qa = import_query_module()


def load_source(source: str) -> list[dict[str, str]]:
    path = qa.source_path(source)
    _, header, rows = qa.table_rows_cached(path, qa.effective_sheet(source, None))
    return [qa.row_dict(header, row) for row in rows]


def parse_num(value: object) -> float | None:
    if value is None:
        return None
    text = str(value).strip().replace(",", "")
    if not text or text in {"-", "—"}:
        return None
    match = re.search(r"\d+(?:\.\d+)?", text)
    return float(match.group(0)) if match else None


def split_terms(value: object) -> list[str]:
    if value in (None, "", [], {}):
        return []
    if isinstance(value, list):
        raw = "、".join(str(item) for item in value)
    else:
        raw = str(value)
    parts = [item.strip() for item in re.split(r"[、,，;；/\s]+", raw) if item.strip()]
    return parts


def compact_list(items: list[str], limit: int = 8) -> str:
    out = []
    seen = set()
    for item in items:
        item = str(item).strip()
        if item and item not in seen:
            out.append(item)
            seen.add(item)
    if len(out) > limit:
        return "、".join(out[:limit]) + f" 等{len(out)}项"
    return "、".join(out) if out else "待2026目录补齐"


def field(data: dict[str, object], *names: str) -> str:
    for name in names:
        value = data.get(name)
        if value not in (None, "", [], {}):
            return str(value)
    return ""


def slug(value: str) -> str:
    text = re.sub(r"[\\/:*?\"<>|]+", "-", value.strip())
    text = re.sub(r"\s+", "-", text)
    return text or "anonymous"


def output_dir_from_candidate(candidate: dict[str, object]) -> Path:
    name = field(candidate, "name", "考生", "姓名") or "anonymous"
    subject = field(candidate, "subject", "科类") or "unknown"
    rank = field(candidate, "rank", "位次", "排位") or "rank"
    return Path("output") / f"{slug(name).lower()}-{slug(subject)}-{slug(rank)}"


def expand_rejection_terms(candidate: dict[str, object]) -> list[str]:
    raw = "、".join(
        [
            field(candidate, "rejected_majors", "排斥专业"),
            field(candidate, "explicit_rejections", "明确不接受"),
            field(candidate, "never_recommend", "任何层级都不得推荐"),
        ]
    )
    terms = set(split_terms(raw))
    for key, values in REJECTION_EXPANSIONS.items():
        if key in raw:
            terms.update(values)
    return sorted(terms, key=len, reverse=True)


def major_terms(candidate: dict[str, object]) -> tuple[list[str], list[str]]:
    preferred = split_terms(field(candidate, "preferred_majors", "倾向专业"))
    acceptable = split_terms(field(candidate, "acceptable_majors", "可接受专业"))
    if not preferred and not acceptable:
        acceptable = list(DEFAULT_MAJOR_TERMS)
    return preferred, acceptable


def region_sets(candidate: dict[str, object]) -> tuple[set[str], set[str], set[str]]:
    preferred_text = field(candidate, "preferred_regions", "倾向地区", "优先地区")
    acceptable_text = field(candidate, "acceptable_regions", "可接受地区")
    rejected_text = field(candidate, "rejected_regions", "明确排斥地区", "排斥地区")
    preferred = expand_regions(preferred_text)
    acceptable = expand_regions(acceptable_text)
    rejected = expand_regions(rejected_text)
    for key, values in REMOTE_REGION_EXPANSIONS.items():
        if key in rejected_text:
            rejected.update(values)
    if not acceptable and ("省外" in acceptable_text or "优质" in acceptable_text):
        acceptable.update({"北京", "天津", "重庆", "湖北", "湖南", "江西", "福建", "四川", "山东", "广西", "海南", "河南", "河北"})
    return preferred, acceptable, rejected


def expand_regions(text: str) -> set[str]:
    out = set()
    for term in split_terms(text):
        out.update(REGION_EXPANSIONS.get(term, {term}))
    for key, values in REGION_EXPANSIONS.items():
        if key in text:
            out.update(values)
    return out


def tuition_limit(candidate: dict[str, object]) -> float:
    return parse_num(field(candidate, "tuition_limit", "budget", "年学费上限")) or 35000.0


def adjustment_value(candidate: dict[str, object]) -> tuple[str, str]:
    text = field(candidate, "adjustment_preference", "是否服从调剂", "服从调剂")
    if "否" in text and "是" not in text:
        return "否", text
    if "是" in text:
        return "是", text
    return "待确认", text or "未提供"


def contains_any(text: str, terms: list[str] | set[str]) -> bool:
    return any(term and term in text for term in terms)


def group_key(row: dict[str, str], group_field: str = "所属专业组") -> tuple[str, str, str]:
    return (row.get("院校代码", ""), row.get("院校名称", ""), row.get(group_field, "").strip())


def risk_layer(rank: float | None, candidate_rank: float) -> str:
    if rank is None:
        return "待定"
    if rank < candidate_rank - 3000:
        return "冲"
    if rank <= candidate_rank + 8000:
        return "稳"
    if rank <= candidate_rank + 25000:
        return "保"
    if rank <= candidate_rank + 40000:
        return "垫"
    return "兜底"


def choose_group_score(rows: list[dict[str, str]], year: str) -> dict[str, str]:
    same_year = [row for row in rows if row.get("年份") == year]
    rows = same_year or rows
    return sorted(
        rows,
        key=lambda row: (
            parse_num(row.get("最低分位")) if parse_num(row.get("最低分位")) is not None else math.inf,
            -(parse_num(row.get("最低分数")) or 0),
        ),
    )[0]


def build_indexes(school_scores, plans, charters, year: str):
    all_group_scores: dict[tuple[str, str, str], list[dict[str, str]]] = defaultdict(list)
    for row in school_scores:
        all_group_scores[(row.get("院校代码", ""), row.get("院校名称", ""), row.get("专业组", "").strip())].append(row)
    group_score = {key: choose_group_score(rows, year) for key, rows in all_group_scores.items()}

    plan_by_group: dict[tuple[str, str, str], list[dict[str, str]]] = defaultdict(list)
    for row in plans:
        if row.get("年份") == year:
            plan_by_group[group_key(row)].append(row)

    charter_by_school = {row.get("院校名称", ""): row for row in charters if row.get("院校名称")}
    return group_score, all_group_scores, plan_by_group, charter_by_school


def plan_summary(rows: list[dict[str, str]], rejected_terms: list[str], max_tuition: float) -> dict[str, object]:
    majors = [row.get("专业名称", "") for row in rows]
    blob = " ".join(majors + [row.get("专业备注", "") for row in rows] + [row.get("招生类型", "") for row in rows])
    fees = [parse_num(row.get("学费(元)")) for row in rows]
    fees = [fee for fee in fees if fee is not None]
    max_fee = max(fees) if fees else None
    bad_majors = [major for major in majors if contains_any(major, rejected_terms)]
    high_fee = bool(max_fee and max_fee > max_tuition) or contains_any(blob, HIGH_FEE_TERMS)
    plan_count = sum(int(parse_num(row.get("招生人数")) or 0) for row in rows)
    return {"majors": majors, "bad_majors": bad_majors, "high_fee": high_fee, "max_fee": max_fee, "plan_count": plan_count}


def major_fit(name: str, preferred: list[str], acceptable: list[str]) -> int:
    if contains_any(name, preferred):
        return 40
    if contains_any(name, acceptable):
        return 24
    return 0


def province_score(province: str, preferred: set[str], acceptable: set[str]) -> int:
    if province in preferred:
        return 28
    if province in acceptable:
        return 16
    return 4 if not preferred and not acceptable else 0


def historical_reference(all_group_scores, key: tuple[str, str, str]) -> str:
    by_year = {}
    for row in all_group_scores.get(key, []):
        year = row.get("年份")
        if year in {"2022", "2023", "2024"} and year not in by_year:
            by_year[year] = f"{row.get('最低分数', '-')}/{row.get('最低分位', '-')}"
    parts = [f"{year}:{by_year[year]}" for year in ["2024", "2023", "2022"] if year in by_year]
    return "；".join(parts) if parts else "同组代码近三年未匹配，需以2026目录复核"


def build_candidates(candidate: dict[str, object], args: argparse.Namespace):
    subject = args.subject or field(candidate, "subject", "科类") or SUBJECT_DEFAULT
    target_batch = args.target_batch or field(candidate, "target_batch", "目标批次") or "普通类本科院校"
    source_batch = args.source_batch or SOURCE_BATCH_BY_TARGET.get(target_batch, "本科批")
    candidate_rank = parse_num(field(candidate, "rank", "位次", "排位"))
    if candidate_rank is None:
        raise SystemExit("candidate.json 缺少可解析的位次/rank，不能构建推荐。")

    preferred_majors, acceptable_majors = major_terms(candidate)
    rejected_terms = expand_rejection_terms(candidate)
    preferred_regions, acceptable_regions, rejected_regions = region_sets(candidate)
    max_tuition = args.max_tuition or tuition_limit(candidate)
    stretch_floor = max(1, candidate_rank - args.max_stretch_advantage)

    school_scores = [r for r in load_source("school-scores") if r.get("科类") == subject and r.get("批次") == source_batch]
    major_scores = [r for r in load_source("major-scores") if r.get("科类") == subject and r.get("批次") == source_batch]
    plans = [r for r in load_source("plans") if r.get("科类") == subject and r.get("批次") == source_batch]
    charters = load_source("charters")
    group_scores, all_group_scores, plan_by_group, charter_by_school = build_indexes(school_scores, plans, charters, args.year)

    grouped: dict[tuple[str, str, str], dict[str, object]] = {}
    rejected = defaultdict(int)
    considered = 0
    for row in major_scores:
        if row.get("年份") != args.year:
            continue
        major = row.get("专业", "")
        if not major_fit(major, preferred_majors, acceptable_majors):
            continue
        considered += 1
        if contains_any(major, rejected_terms):
            rejected["专业命中排斥词"] += 1
            continue
        province = row.get("学校所在", "")
        if province in rejected_regions:
            rejected["地域命中排斥范围"] += 1
            continue
        if (preferred_regions or acceptable_regions) and province not in preferred_regions and province not in acceptable_regions:
            rejected["地域不在优先或可接受范围"] += 1
            continue
        key = group_key(row)
        group_score = group_scores.get(key)
        if not group_score:
            rejected["缺少院校专业组投档线"] += 1
            continue
        if group_score.get("年份") != args.year:
            rejected["缺少目标年份院校专业组投档线"] += 1
            continue
        if group_score.get("招生类型") and group_score.get("招生类型") != "普通类":
            rejected["非普通类或特殊招生类型"] += 1
            continue
        group_rank = parse_num(group_score.get("最低分位"))
        if group_rank is not None and group_rank < stretch_floor:
            rejected["冲刺位次差过大"] += 1
            continue
        info = plan_summary(plan_by_group.get(key, []), rejected_terms, max_tuition)
        if info["bad_majors"]:
            rejected["专业组内存在排斥专业"] += 1
            continue
        if info["high_fee"]:
            rejected["学费或招生类型触发高收费排除"] += 1
            continue
        item = grouped.setdefault(
            key,
            {
                "key": key,
                "school_code": key[0],
                "school": key[1],
                "group": key[2],
                "province": province,
                "nature": row.get("学校性质", ""),
                "is_985": row.get("是否985", ""),
                "is_211": row.get("是否211", ""),
                "group_score": group_score,
                "plan_rows": plan_by_group.get(key, []),
                "major_rows": [],
                "charter": charter_by_school.get(key[1], {}),
            },
        )
        item["major_rows"].append(row)

    candidates = []
    for item in grouped.values():
        info = plan_summary(item["plan_rows"], rejected_terms, max_tuition)
        group_rank = parse_num(item["group_score"].get("最低分位"))
        layer = risk_layer(group_rank, candidate_rank)
        majors = sorted(
            item["major_rows"],
            key=lambda row: (-major_fit(row.get("专业", ""), preferred_majors, acceptable_majors), parse_num(row.get("最低位次")) or math.inf),
        )
        best = majors[0]
        prestige = (14 if item["is_985"] == "是" else 0) + (8 if item["is_211"] == "是" else 0)
        preferred_school_bonus = 18 if item["school"] in field(candidate, "preferred_schools", "倾向院校") else 0
        layer_bonus = {"冲": -8, "稳": 14, "保": 18, "垫": 8, "兜底": 3}.get(layer, 0)
        best_rank = parse_num(best.get("最低位次")) or group_rank or candidate_rank
        score = (
            province_score(item["province"], preferred_regions, acceptable_regions)
            + preferred_school_bonus
            + prestige
            + major_fit(best.get("专业", ""), preferred_majors, acceptable_majors)
            + layer_bonus
            - abs(best_rank - candidate_rank) / 5000
        )
        candidates.append({**item, "layer": layer, "best_major": best, "info": info, "score": score, "historical": historical_reference(all_group_scores, item["key"])})
    return candidates, dict(rejected), considered, {"subject": subject, "target_batch": target_batch, "source_batch": source_batch, "rank": candidate_rank}


def choose(candidates: list[dict[str, object]], capacity: int) -> list[dict[str, object]]:
    quotas = LAYER_QUOTAS_45 if capacity == 45 else None
    buckets = defaultdict(list)
    for item in candidates:
        buckets[item["layer"]].append(item)
    for layer in buckets:
        buckets[layer].sort(key=lambda item: item["score"], reverse=True)

    selected = []
    per_school = defaultdict(int)

    def can_add(item):
        return per_school[item["school"]] < 3

    if quotas:
        for layer, quota in quotas.items():
            for item in buckets[layer]:
                if len([x for x in selected if x["layer"] == layer]) >= quota:
                    break
                if can_add(item):
                    selected.append(item)
                    per_school[item["school"]] += 1
    leftovers = [item for item in candidates if item not in selected]
    leftovers.sort(key=lambda item: item["score"], reverse=True)
    for item in leftovers:
        if len(selected) >= capacity:
            break
        if can_add(item):
            selected.append(item)
            per_school[item["school"]] += 1
    order = {"冲": 0, "稳": 1, "保": 2, "垫": 3, "兜底": 4}
    selected.sort(key=lambda item: (order.get(item["layer"], 9), -(parse_num(item["group_score"].get("最低分位")) or 0)))
    return selected[:capacity]


def rank_text(score: str, rank: str) -> str:
    return f"{score or '-'} / {rank or '-'}"


def to_output_row(item: dict[str, object], index: int, meta: dict[str, object], candidate: dict[str, object]) -> dict[str, str]:
    majors = item["major_rows"][:]
    majors.sort(key=lambda row: (-parse_num(row.get("最低分数") or 0), parse_num(row.get("最低位次")) or math.inf))
    selected = []
    seen_major_keys = set()
    for row in majors:
        major_key = (row.get("专业", ""), row.get("专业代码", ""))
        if major_key in seen_major_keys:
            continue
        selected.append(row)
        seen_major_keys.add(major_key)
        if len(selected) >= 6:
            break
    codes = [row.get("专业代码", "") for row in selected if row.get("专业代码")]
    code_values = codes[:6] + [""] * (6 - len(codes))
    names_with_codes = [f"{row.get('专业')}({row.get('专业代码')})" if row.get("专业代码") else row.get("专业", "") for row in selected]
    plan_majors = [row.get("专业名称", "") for row in item["plan_rows"]]
    acceptable = plan_majors or [row.get("专业", "") for row in selected]
    adjustment, adjustment_note = adjustment_value(candidate)
    group_score = item["group_score"]
    best = item["best_major"]
    info = item["info"]
    charter = item.get("charter") or {}
    max_fee = int(info["max_fee"]) if info["max_fee"] else None
    label = "、".join(part for part in [("985" if item["is_985"] == "是" else ""), ("211" if item["is_211"] == "是" else "")] if part) or "普通公办"
    code_status = "专业组/专业代码来自历史数据，待2026官方招生专业目录复核"
    risk = "；".join(
        [
            code_status,
            "正式填报前复核章程、体检、单科/语种、收费和专业组内调剂范围",
        ]
    )
    reason = f"{item['province']}；{compact_list([row.get('专业', '') for row in selected[:3]], 3)}方向匹配；按{HISTORY_YEAR}院校专业组投档位次归为{item['layer']}层"
    note = f"{item['layer']}；{reason}；{risk}"
    row = {
        "志愿号": str(index),
        "层级": item["layer"],
        "录取批次": str(meta["target_batch"]),
        "科类": str(meta["subject"]),
        "院校代码": item["school_code"],
        "院校名称": item["school"],
        "院校专业组代码": item["group"],
        "专业代码列表": ",".join(codes),
        "是否服从调剂": adjustment,
        "调剂条件": adjustment_note,
        "推荐专业": compact_list(names_with_codes, 6),
        "组内可接受专业": compact_list(acceptable, 10),
        "2025专业最低分/位次": rank_text(best.get("最低分数", ""), best.get("最低位次", "")),
        "2025组最低分/位次": rank_text(group_score.get("最低分数", ""), group_score.get("最低分位", "")),
        "2024/2023/2022参考": item["historical"],
        "招生计划变化": f"{HISTORY_YEAR}同组计划约{info['plan_count']}人；2026待官方目录核验",
        "地域": item["province"],
        "学校性质": item["nature"],
        "标签": label,
        "学费风险": f"{HISTORY_YEAR}同组最高约{max_fee}元/年" if max_fee else "计划表未给出或待核验",
        "推荐理由": reason,
        "风险提示": risk,
        "章程核查状态": charter.get("2025招生章程", "未匹配到2025章程链接，需高校官网复核"),
        "代码复核状态": code_status,
        "备注": note,
    }
    for idx in range(6):
        row[f"专业代码{idx + 1}"] = code_values[idx]
    return {field: str(row.get(field, "")) for field in OUTPUT_FIELDS}


def markdown_table(rows: list[dict[str, str]]) -> str:
    headers = ["志愿号", "层级", "院校名称", "院校代码", "院校专业组代码", "地域", "推荐专业", "组内可接受专业", "2025专业最低分/位次", "2025组最低分/位次", "学费风险", "推荐理由", "风险提示", "是否服从调剂"]
    lines = ["| " + " | ".join(headers) + " |", "| " + " | ".join("---" for _ in headers) + " |"]
    for row in rows:
        lines.append("| " + " | ".join(row.get(header, "").replace("|", "/") for header in headers) + " |")
    return "\n".join(lines)


def coverage_lines(rows: list[dict[str, str]], capacity: int, candidate_count: int) -> list[str]:
    if len(rows) >= capacity:
        return [
            f"- 完整推荐数/官方容量：{len(rows)}/{capacity}。",
            "- 覆盖结论：已达到官方容量。",
        ]
    exception_type = "data_insufficient" if candidate_count < capacity else "hard_constraints"
    return [
        f"- 完整推荐数/官方容量：{len(rows)}/{capacity}。",
        f"- 覆盖结论：未达到官方容量，未满额例外类型暂按 `{exception_type}` 标记。",
        "- 已尝试扩展查询：脚本已交叉读取 school-scores、major-scores、plans、charters，并按位次、专业、地域、费用、招生类型和组内排斥专业过滤。",
        "- 继续补满的主要风险：可能引入与资料卡不匹配的专业组、地域、费用或调剂风险。",
        "- 可放宽约束建议：如用户确认愿意扩大地域、相近工科专业或院校层次，可重新运行推荐脚本并扩大候选范围。",
    ]


def render_review(rows: list[dict[str, str]], candidate: dict[str, object], meta: dict[str, object], rejected: dict[str, int], considered: int, candidate_count: int, capacity: int) -> str:
    by_layer = defaultdict(list)
    for row in rows:
        by_layer[row["层级"]].append(row)
    name = field(candidate, "name", "考生", "姓名") or "anonymous"
    score = field(candidate, "score", "分数")
    lines = [
        f"# {name} 2026广东高考志愿推荐清单（审阅稿）",
        "",
        "## 考生摘要",
        f"- 考生：{name}",
        f"- 科类/位次：{meta['subject']}；{int(meta['rank']) if float(meta['rank']).is_integer() else meta['rank']}位",
        f"- 分数：{score or '未提供'}",
        f"- 目标批次：{meta['target_batch']}，官方容量 {capacity}。",
        f"- 调剂偏好：{adjustment_value(candidate)[1]}",
        "",
        "## 数据源和查询口径",
        f"- 使用脚本：`<skill_dir>/scripts/build_recommendations.py`；底层读取 `school-scores`、`major-scores`、`plans`、`charters`。",
        f"- 数据年份：{HISTORY_YEAR} 为主，2024/2023/2022 作为趋势参考。",
        f"- 源批次/科类：{meta['source_batch']}；{meta['subject']}。",
        f"- 候选筛选：专业层记录初筛 {considered} 条；有效院校专业组候选 {candidate_count} 个；剔除原因统计：{json.dumps(rejected, ensure_ascii=False)}。",
        "- 注意：历史专业组代码和专业代码不得视为2026最终代码，必须以2026官方招生专业目录复核。",
        "",
        "## 覆盖情况",
        *coverage_lines(rows, capacity, candidate_count),
        "- 用户确认前不得生成最终志愿表草稿；确认后保留同一 JSON schema 写入 `recommendations-confirmed.json`。",
        "",
    ]
    notes = {
        "冲": "历史组位次高于当前位次，少量尝试。",
        "稳": "历史组位次接近当前位次，优先匹配目标专业和城市。",
        "保": "历史组位次低于当前位次，用于提升录取稳妥性。",
        "垫": "明显低于当前位次，用于降低滑档风险。",
        "兜底": "足够低位次兜底，重点控制费用和专业组内调剂风险。",
    }
    for layer in ["冲", "稳", "保", "垫", "兜底"]:
        group_rows = by_layer.get(layer, [])
        if group_rows:
            lines.extend([f"## {layer}层（{len(group_rows)}个）", f"- 风险说明：{notes[layer]}", markdown_table(group_rows), ""])
    lines.extend(
        [
            "## 需要用户确认的问题",
            "- 是否删除任何院校专业组，或调整冲/稳/保/垫/兜底比例。",
            "- 是否确认体检、色觉、单科、语种无限制；否则必须重新筛选。",
            "- 是否确认专业组内列出的可接受专业均可接受调剂。",
            "- 是否确认不接受的费用、地域和专业约束没有变化。",
            "",
            "## 2026最终复核事项",
            "- 以广东省教育考试院2026招生专业目录、增补更正公告和官方志愿系统为准。",
            "- 逐项核对院校代码、专业组代码、专业代码、计划数、学费、办学地点、体检要求、单科/语种限制和调剂范围。",
        ]
    )
    return "\n".join(lines) + "\n"


def render_query_log(meta: dict[str, object], rejected: dict[str, int], considered: int, candidate_count: int, selected_count: int, capacity: int) -> str:
    return "\n".join(
        [
            "# 查询与筛选日志",
            "",
            f"- 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"- 科类/批次：{meta['subject']}；{meta['source_batch']}。",
            f"- 位次：{meta['rank']}。",
            "- 数据源：school-scores、major-scores、plans、charters。",
            "- 策略：先按专业偏好从 major-scores 建候选，再用 school-scores 专业组投档线分层，用 plans 检查计划、学费和组内排斥专业，用 charters 补章程链接。",
            f"- 初筛专业记录数：{considered}。",
            f"- 有效院校专业组候选数：{candidate_count}。",
            f"- 推荐覆盖：{selected_count}/{capacity}。",
            f"- 剔除原因统计：{json.dumps(rejected, ensure_ascii=False, indent=2)}",
            "",
        ]
    )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--candidate", required=True, help="Path to candidate.json")
    parser.add_argument("--output-dir", help="Task output directory; defaults to output/{name}-{subject}-{rank}")
    parser.add_argument("--year", default=HISTORY_YEAR, help="Historical data year to use")
    parser.add_argument("--subject", help="Override candidate subject")
    parser.add_argument("--target-batch", help="Override candidate target batch label")
    parser.add_argument("--source-batch", help="Override data source batch, e.g. 本科批")
    parser.add_argument("--capacity", type=int, help="Recommendation count; defaults by target batch")
    parser.add_argument("--max-tuition", type=float, help="Override tuition cap")
    parser.add_argument("--max-stretch-advantage", type=float, default=25000, help="Reject groups whose 2025 rank is more than this many places above candidate rank")
    args = parser.parse_args()

    candidate_path = Path(args.candidate)
    candidate = json.loads(candidate_path.read_text(encoding="utf-8-sig"))
    out_dir = Path(args.output_dir) if args.output_dir else output_dir_from_candidate(candidate)
    out_dir.mkdir(parents=True, exist_ok=True)
    intermediate = out_dir / "intermediate"
    intermediate.mkdir(parents=True, exist_ok=True)

    target_batch = args.target_batch or field(candidate, "target_batch", "目标批次") or "普通类本科院校"
    capacity = args.capacity or CAPACITY_BY_TARGET.get(target_batch, 45)
    candidates, rejected, considered, meta = build_candidates(candidate, args)
    selected = choose(candidates, capacity)
    rows = [to_output_row(item, index + 1, meta, candidate) for index, item in enumerate(selected)]

    (intermediate / "candidate-pool.json").write_text(
        json.dumps([to_output_row(item, index + 1, meta, candidate) for index, item in enumerate(candidates)], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (intermediate / "recommendations-proposed.json").write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "recommendations-review.md").write_text(render_review(rows, candidate, meta, rejected, considered, len(candidates), capacity), encoding="utf-8")
    (out_dir / "query-log.md").write_text(render_query_log(meta, rejected, considered, len(candidates), len(rows), capacity), encoding="utf-8")
    print(json.dumps({"candidate_pool": len(candidates), "selected": len(rows), "output_dir": str(out_dir)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
