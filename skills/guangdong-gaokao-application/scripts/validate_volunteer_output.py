#!/usr/bin/env python3
"""Validate Guangdong volunteer draft xlsx completeness and coverage."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from render_volunteer_draft import (
    FIELDS,
    canonical_batch,
    filled_official_fields,
    official_batch_count,
)


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def col_to_num(ref: str) -> int:
    match = re.match(r"([A-Z]+)", ref or "")
    if not match:
        return 0
    value = 0
    for char in match.group(1):
        value = value * 26 + ord(char) - 64
    return value


def shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    return ["".join(t.text or "" for t in si.iter(f"{{{NS['main']}}}t")) for si in root.findall("main:si", NS)]


def worksheet_path(zf: zipfile.ZipFile, sheet_name: str) -> str:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall("pkgrel:Relationship", NS)}
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        if sheet.attrib["name"] == sheet_name:
            rid = sheet.attrib[f"{{{NS['rel']}}}id"]
            target = rel_map[rid]
            path = target.lstrip("/") if target.startswith("/") else target
            return path if path.startswith("xl/") else "xl/" + path
    raise SystemExit(f"Sheet not found: {sheet_name}")


def cell_value(cell: ET.Element, shared: list[str]) -> str:
    typ = cell.attrib.get("t")
    if typ == "inlineStr":
        return "".join(t.text or "" for t in cell.iter(f"{{{NS['main']}}}t")).strip()
    value = cell.find("main:v", NS)
    if value is None:
        return ""
    text = value.text or ""
    if typ == "s":
        try:
            return shared[int(text)]
        except (ValueError, IndexError):
            return text
    return text.strip()


def sheet_rows(path: Path, sheet_name: str) -> list[list[str]]:
    with zipfile.ZipFile(path) as zf:
        shared = shared_strings(zf)
        root = ET.fromstring(zf.read(worksheet_path(zf, sheet_name)))
        out = []
        for row in root.findall("main:sheetData/main:row", NS):
            values: list[str] = []
            for cell in row.findall("main:c", NS):
                col = col_to_num(cell.attrib.get("r", ""))
                while len(values) < col - 1:
                    values.append("")
                values.append(cell_value(cell, shared))
            out.append(values)
        return out


def row_dicts(path: Path, sheet_name: str) -> list[dict[str, str]]:
    rows = sheet_rows(path, sheet_name)
    if not rows:
        raise SystemExit("志愿表草稿 sheet is empty")
    header = rows[0]
    missing = [field for field in FIELDS if field not in header]
    if missing:
        raise SystemExit("志愿表草稿缺少字段：" + "、".join(missing))
    return [
        {header[i]: row[i] if i < len(row) else "" for i in range(len(header))}
        for row in rows[1:]
        if any(row)
    ]


def has_any_entry(row: dict[str, str]) -> bool:
    return bool(row.get("院校代码") or row.get("院校名称") or row.get("院校专业组代码") or any(row.get(f"专业代码{i}") for i in range(1, 7)))


def validate(rows: list[dict[str, str]], required_batches: list[str], require_full: bool = False) -> dict[str, object]:
    errors = []
    warnings = []
    summaries = []
    touched = []
    for row in rows:
        if has_any_entry(row):
            batch = canonical_batch(row.get("录取批次", ""))
            if batch not in touched:
                touched.append(batch)
    batches = [canonical_batch(batch) for batch in required_batches] if required_batches else touched
    for batch in batches:
        batch_rows = [row for row in rows if canonical_batch(row.get("录取批次", "")) == batch]
        expected = official_batch_count(batch)
        filled = sum(1 for row in batch_rows if filled_official_fields(row))
        partial = [row for row in batch_rows if has_any_entry(row) and not filled_official_fields(row)]
        coverage_ratio = round(filled / expected, 4) if expected else None
        summaries.append(
            {
                "batch": batch,
                "expected": expected,
                "rows": len(batch_rows),
                "filled": filled,
                "partial": len(partial),
                "coverage_ratio": coverage_ratio,
                "coverage_status": "full" if expected and filled >= expected else "underfilled",
            }
        )
        if expected and len(batch_rows) < expected:
            errors.append(f"{batch} 行数不足：存在 {len(batch_rows)} 行，官方需要 {expected} 行。")
        if expected and filled < expected:
            message = f"{batch} 覆盖未满：完整志愿 {filled}/{expected}。"
            if require_full:
                errors.append(message)
            else:
                warnings.append(
                    message
                    + "最终交叉验证必须说明未满额例外类型、资料卡依据和已尝试扩展查询；"
                    "如资料卡没有明确强偏好、高分强匹配、不可放宽约束或数据不足证据，应回退补充可接受志愿。"
                )
        if not expected:
            warnings.append(f"{batch} 未识别官方志愿数量，请人工复核。")
        for row in partial[:10]:
            errors.append(f"{batch} 志愿{row.get('志愿号') or '?'} 已部分填写但不完整。")
    return {"ok": not errors, "summaries": summaries, "errors": errors, "warnings": warnings}


def print_markdown(result: dict[str, object]) -> None:
    print("# 志愿表输出校验")
    print("")
    print(f"- 结果：{'通过' if result['ok'] else '未通过'}")
    print("")
    print("| 批次 | 官方数量 | 表内行数 | 完整填写 | 覆盖率 | 覆盖状态 | 部分填写 |")
    print("| --- | ---: | ---: | ---: | ---: | --- | ---: |")
    for item in result["summaries"]:
        ratio = "" if item["coverage_ratio"] is None else f"{item['coverage_ratio']:.2%}"
        print(
            f"| {item['batch']} | {item['expected']} | {item['rows']} | {item['filled']} | "
            f"{ratio} | {item['coverage_status']} | {item['partial']} |"
        )
    if result["errors"]:
        print("")
        print("## 错误")
        for item in result["errors"]:
            print(f"- {item}")
    if result["warnings"]:
        print("")
        print("## 警告")
        for item in result["warnings"]:
            print(f"- {item}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--file", required=True, help="Volunteer draft xlsx path")
    parser.add_argument("--batch", action="append", default=[], help="Batch to validate and summarize; may be repeated")
    parser.add_argument("--sheet", default="志愿表草稿")
    parser.add_argument("--format", choices=["markdown", "json"], default="markdown")
    parser.add_argument("--require-full", action="store_true", help="Fail when selected batches are not filled to official capacity")
    args = parser.parse_args()

    result = validate(row_dicts(Path(args.file), args.sheet), args.batch, args.require_full)
    if args.format == "json":
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print_markdown(result)
    if not result["ok"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
