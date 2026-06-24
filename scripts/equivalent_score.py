#!/usr/bin/env python3
"""Look up Guangdong one-score-one-rank data and historical equivalent scores."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
SEGMENT_DIR = ROOT / "references/data/segments"
NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def col_to_num(ref: str) -> int:
    match = re.match(r"([A-Z]+)", ref or "")
    value = 0
    if not match:
        return 0
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


def first_sheet(zf: zipfile.ZipFile) -> str:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall("pkgrel:Relationship", NS)}
    sheet = workbook.find("main:sheets/main:sheet", NS)
    target = rel_map[sheet.attrib[f"{{{NS['rel']}}}id"]]
    path = target.lstrip("/") if target.startswith("/") else target
    return path if path.startswith("xl/") else "xl/" + path


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
    return text


def rows(path: Path):
    with zipfile.ZipFile(path) as zf:
        shared = shared_strings(zf)
        root = ET.fromstring(zf.read(first_sheet(zf)))
        for row in root.findall("main:sheetData/main:row", NS):
            values: list[str] = []
            for cell in row.findall("main:c", NS):
                col = col_to_num(cell.attrib.get("r", ""))
                while len(values) < col - 1:
                    values.append("")
                values.append(cell_value(cell, shared).strip())
            while values and values[-1] == "":
                values.pop()
            if any(values):
                yield values


def parse_range(text: str) -> tuple[int, int] | None:
    nums = [int(x) for x in re.findall(r"\d+", text or "")]
    if not nums:
        return None
    if len(nums) == 1:
        return nums[0], nums[0]
    return nums[0], nums[1]


def score_matches(score_cell: str, score: int) -> bool:
    rng = parse_range(score_cell)
    return bool(rng and rng[0] <= score <= rng[1])


def rank_matches(rank_cell: str, rank: int) -> bool:
    rng = parse_range(rank_cell)
    return bool(rng and rng[0] <= rank <= rng[1])


def load_year(year: int) -> tuple[list[str], list[dict[str, str]]]:
    path = SEGMENT_DIR / f"guangdong-segments-{year}.xlsx"
    if not path.exists():
        raise SystemExit(f"Missing one-score-one-rank file: {path}")
    it = rows(path)
    header = next(it)
    data = [{header[i]: row[i] if i < len(row) else "" for i in range(len(header))} for row in it]
    return header, data


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--year", type=int, required=True, choices=[2022, 2023, 2024, 2025])
    parser.add_argument("--subject", required=True, help="物理类 or 历史类")
    parser.add_argument("--score", type=int)
    parser.add_argument("--rank", type=int)
    parser.add_argument("--format", choices=["json", "markdown"], default="markdown")
    args = parser.parse_args()

    if args.score is None and args.rank is None:
        raise SystemExit("Provide --score or --rank")

    _, data = load_year(args.year)
    matches = []
    for row in data:
        if args.subject not in row.get("科类", ""):
            continue
        if args.score is not None and score_matches(row.get("分数(分)", ""), args.score):
            matches.append(row)
        elif args.rank is not None and rank_matches(row.get("排名区间", ""), args.rank):
            matches.append(row)

    if args.format == "json":
        print(json.dumps(matches, ensure_ascii=False, indent=2))
        return
    if not matches:
        print("No matching segment found.")
        return
    headers = ["年份", "科类", "批次", "控制线(分)", "分数(分)", "本段人数(人)", "累计人数(人)", "排名区间", "历史同位次考生得分"]
    print("| " + " | ".join(headers) + " |")
    print("| " + " | ".join("---" for _ in headers) + " |")
    for row in matches:
        print("| " + " | ".join(row.get(h, "") for h in headers) + " |")


if __name__ == "__main__":
    main()
