#!/usr/bin/env python3
"""Query Guangdong admission Excel workbooks without third-party packages."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import pickle
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
CACHE_VERSION = 1

SOURCES = {
    "school-scores": "data/scores/school-scores-2022-2025.xlsx",
    "major-scores": "data/scores/major-scores-2022-2025.xlsx",
    "plans": "data/plans/admission-plans-2022-2025.xlsx",
    "charters": "data/charters/charters-2025.xlsx",
    "art-school-scores": "data/art/art-school-scores.xlsx",
    "art-major-scores": "data/art/art-major-scores.xlsx",
    "art-plans": "data/art/art-plans.xlsx",
    "spring": "data/spring/spring-2026.xlsx",
}

SOURCE_DEFAULT_SHEETS = {
    "spring": "院校录取信息表",
}

SOURCE_DEFAULT_COLUMNS = {
    "spring": [
        "类型", "院校代码", "院校名称", "专业组代码", "2025计划数", "2024计划数",
        "2025最低分", "2024最低分", "2025最低位次", "2024最低位次",
        "城市", "办学性质", "招生章程",
    ],
    "art-school-scores": [
        "年份", "学校", "科类", "批次", "专业类别", "专业组", "选科要求",
        "投档线", "最低位次", "计算公式", "备注",
    ],
    "art-major-scores": [
        "年份", "学校", "科类", "批次", "专业名称", "考试类别",
        "专业控线(分)", "文化控线(分)", "最低分", "最低位次", "计算公式", "备注",
    ],
    "art-plans": [
        "年份", "学校", "科类", "批次", "专业类别", "专业组", "专业名称",
        "招生计划(人)", "学制", "学费(元/年)", "选科要求", "备注",
    ],
}

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

ALIASES = {
    "year": ["年份"],
    "school": ["院校名称", "学校", "院校"],
    "school_code": ["院校代码", "全国统一招生代码"],
    "subject": ["科类", "文理科", "计划类别名称"],
    "batch": ["批次", "批次名称"],
    "admission_type": ["招生类型", "计划类别名称", "类型"],
    "group": ["专业组", "所属专业组", "院校专业组代码", "专业组代码"],
    "major": ["专业", "专业名称"],
    "major_code": ["专业代码"],
    "major_note": ["专业备注", "备注"],
    "requirements": ["选科要求", "选考科目", "再选科目要求"],
    "plan_count": ["招生人数", "招生计划(人)", "计划数", "2025计划数", "2024计划数"],
    "score": ["最低分数", "最低分", "投档线", "投档最低分", "2025最低分", "2024最低分", "平均最低分"],
    "rank": ["最低分位", "最低位次", "最低分排名", "投档最低排位", "2025最低位次", "2024最低位次", "平均最低位次"],
    "line_diff": ["批次线差"],
    "province": ["学校所在", "省份", "所在省", "院校省份"],
    "city": ["城市", "院校城市"],
    "nature": ["学校性质", "办学性质", "性质"],
    "is_985": ["是否985", "985", "_985"],
    "is_211": ["是否211", "211", "_211"],
    "tuition": ["学费(元)", "学费", "收费标准", "学费(元/年)"],
    "charter": ["2025招生章程", "招生章程"],
    "homepage": ["院校主页链接"],
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
    values = []
    for si in root.findall("main:si", NS):
        values.append("".join(t.text or "" for t in si.iter(f"{{{NS['main']}}}t")))
    return values


def sheet_paths(zf: zipfile.ZipFile) -> list[dict[str, str]]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall("pkgrel:Relationship", NS)}
    sheets = []
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        rid = sheet.attrib[f"{{{NS['rel']}}}id"]
        target = rel_map[rid]
        path = target.lstrip("/") if target.startswith("/") else target
        if not path.startswith("xl/"):
            path = "xl/" + path
        sheets.append({"name": sheet.attrib["name"], "state": sheet.attrib.get("state", "visible"), "path": path})
    return sheets


def cell_value(cell: ET.Element, shared: list[str]) -> str:
    typ = cell.attrib.get("t")
    if typ == "inlineStr":
        return "".join(t.text or "" for t in cell.iter(f"{{{NS['main']}}}t")).strip()
    value = cell.find("main:v", NS)
    if value is None:
        formula = cell.find("main:f", NS)
        return "=" + (formula.text or "") if formula is not None else ""
    text = value.text or ""
    if typ == "s":
        try:
            return shared[int(text)]
        except (ValueError, IndexError):
            return text
    if typ == "b":
        return "TRUE" if text == "1" else "FALSE"
    return text


def normalize_header_cell(value: str) -> str:
    return re.sub(r"\s+", "", value or "")


def normalize_header_row(row: list[str]) -> list[str]:
    return [normalize_header_cell(cell) for cell in row]


def iter_rows(path: Path, sheet_name: str | None = None):
    with zipfile.ZipFile(path) as zf:
        shared = shared_strings(zf)
        sheets = sheet_paths(zf)
        if sheet_name:
            matches = [s for s in sheets if sheet_name in s["name"]]
            if not matches:
                raise SystemExit(f"Sheet not found: {sheet_name}")
            sheets = matches[:1]
        else:
            sheets = sheets[:1]
        for sheet in sheets:
            root = ET.fromstring(zf.read(sheet["path"]))
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
                    yield sheet["name"], values


def source_path(name: str) -> Path:
    if name not in SOURCES:
        raise SystemExit(f"Unknown source: {name}")
    path = ROOT / SOURCES[name]
    if not path.exists():
        raise SystemExit(f"Missing source file: {path}")
    return path


def effective_sheet(source: str, sheet_name: str | None) -> str | None:
    return sheet_name or SOURCE_DEFAULT_SHEETS.get(source)


def parse_number(value: str) -> float | None:
    if value is None:
        return None
    text = str(value).strip().replace(",", "")
    if text in {"", "-", "—"}:
        return None
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group(0)) if match else None


def header_index(header: list[str]) -> dict[str, int]:
    out = {}
    for key, names in ALIASES.items():
        for name in names:
            if name in header:
                out[key] = header.index(name)
                break
    return out


def header_score(row: list[str]) -> int:
    row = normalize_header_row(row)
    names = {name for aliases in ALIASES.values() for name in aliases}
    return sum(1 for cell in row if cell in names)


def table_rows(path: Path, sheet_name: str | None = None) -> tuple[str, list[str], list[list[str]]]:
    rows = list(iter_rows(path, sheet_name))
    if not rows:
        raise SystemExit("No rows found")
    best = 0
    best_score = -1
    for i, (_, row) in enumerate(rows[:20]):
        score = header_score(row)
        if score > best_score:
            best = i
            best_score = score
    if best_score <= 0:
        best = 0
    sheet, header = rows[best]
    return sheet, normalize_header_row(header), [row for _, row in rows[best + 1 :]]




def default_cache_dir() -> Path:
    """Project-local cache for parsed workbook rows.

    The admission workbooks are .xlsx zip archives. Parsing their XML on every
    query is the expensive part, so cache the normalized table rows outside the
    skill directory and invalidate by workbook path, sheet, size and mtime.
    """

    return Path.cwd() / ".cache" / "guangdong-gaokao"


def cache_path_for(path: Path, sheet_name: str | None, cache_dir: Path) -> Path:
    key = f"{path.resolve()}|{sheet_name or ''}"
    digest = hashlib.sha256(key.encode("utf-8")).hexdigest()[:12]
    return cache_dir / f"{path.stem}-{digest}.pickle"


def cache_key(path: Path, sheet_name: str | None) -> dict[str, object]:
    stat = path.stat()
    return {
        "version": CACHE_VERSION,
        "path": str(path.resolve()),
        "sheet_name": sheet_name,
        "mtime_ns": stat.st_mtime_ns,
        "size": stat.st_size,
    }


def table_rows_cached(
    path: Path,
    sheet_name: str | None = None,
    cache_dir: Path | None = None,
    use_cache: bool = True,
    rebuild_cache: bool = False,
) -> tuple[str, list[str], list[list[str]]]:
    if not use_cache:
        return table_rows(path, sheet_name)

    cache_dir = cache_dir or default_cache_dir()
    key = cache_key(path, sheet_name)
    cache_file = cache_path_for(path, sheet_name, cache_dir)
    if not rebuild_cache and cache_file.exists():
        try:
            with cache_file.open("rb") as fh:
                cached = pickle.load(fh)
            if cached.get("key") == key:
                return cached["sheet"], cached["header"], cached["rows"]
        except Exception:
            cache_file.unlink(missing_ok=True)

    sheet, header, rows = table_rows(path, sheet_name)
    try:
        cache_dir.mkdir(parents=True, exist_ok=True)
        temp_file = cache_file.with_suffix(".tmp")
        with temp_file.open("wb") as fh:
            pickle.dump(
                {"key": key, "sheet": sheet, "header": header, "rows": rows},
                fh,
                protocol=pickle.HIGHEST_PROTOCOL,
            )
        temp_file.replace(cache_file)
    except OSError:
        # Cache is an optimization only. Keep queries functional in read-only projects.
        pass
    return sheet, header, rows


def get(row: list[str], index: dict[str, int], key: str) -> str:
    pos = index.get(key)
    return row[pos] if pos is not None and pos < len(row) else ""


def row_dict(header: list[str], row: list[str]) -> dict[str, str]:
    return {header[i]: row[i] if i < len(row) else "" for i in range(len(header))}


def dict_get(row: dict[str, str], key: str) -> str:
    for name in ALIASES.get(key, [key]):
        value = row.get(name)
        if value not in (None, ""):
            return str(value)
    return ""


def contains_all(text: str, terms: list[str]) -> bool:
    return all(term in text for term in terms if term)


def split_terms(value: str | None) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in re.split(r"[,\s]+", value) if item.strip()]


def passes(row: list[str], idx: dict[str, int], args: argparse.Namespace) -> bool:
    simple = [
        ("year", args.year),
        ("subject", args.subject),
        ("batch", args.batch),
        ("province", args.province),
        ("nature", args.nature),
        ("admission_type", args.admission_type),
        ("group", args.group),
    ]
    for key, expected in simple:
        if expected and expected not in get(row, idx, key):
            return False
    if args.school_code and args.school_code != get(row, idx, "school_code"):
        return False
    if args.school_exact and args.school_exact != get(row, idx, "school"):
        return False
    if args.major_exact and args.major_exact != get(row, idx, "major"):
        return False
    if args.school_contains and args.school_contains not in get(row, idx, "school"):
        return False
    if args.major_contains and args.major_contains not in get(row, idx, "major"):
        return False
    blob = " ".join(row)
    if not contains_all(blob, split_terms(args.keywords)):
        return False
    if any(term in blob for term in split_terms(args.exclude_keywords)):
        return False

    score = parse_number(get(row, idx, "score"))
    rank = parse_number(get(row, idx, "rank"))
    if args.score_min is not None and (score is None or score < args.score_min):
        return False
    if args.score_max is not None and (score is None or score > args.score_max):
        return False
    if args.rank_min is not None and (rank is None or rank < args.rank_min):
        return False
    if args.rank_max is not None and (rank is None or rank > args.rank_max):
        return False
    if args.score_center is not None:
        width = args.score_width or 0
        if score is None or not (args.score_center - width <= score <= args.score_center + width):
            return False
    if args.rank_center is not None:
        width = args.rank_width or 0
        if rank is None or not (args.rank_center - width <= rank <= args.rank_center + width):
            return False
    return True


def inspect_source(args: argparse.Namespace) -> None:
    path = source_path(args.source)
    with zipfile.ZipFile(path) as zf:
        print(
            json.dumps(
                {
                    "source": args.source,
                    "file": str(path),
                    "default_sheet": SOURCE_DEFAULT_SHEETS.get(args.source),
                    "sheets": sheet_paths(zf),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    sheet, header, _ = table_rows_cached(
        path,
        effective_sheet(args.source, args.sheet),
        cache_dir=Path(args.cache_dir) if args.cache_dir else None,
        use_cache=not args.no_cache,
        rebuild_cache=args.rebuild_cache,
    )
    print(json.dumps({"sheet": sheet, "columns": header}, ensure_ascii=False, indent=2))


def list_sources() -> None:
    print(json.dumps({name: SOURCES[name] for name in sorted(SOURCES)}, ensure_ascii=False, indent=2))


def select_columns(header: list[str], rows: list[dict[str, str]], args: argparse.Namespace) -> list[dict[str, str]]:
    if args.columns:
        columns = [c.strip() for c in args.columns.split(",") if c.strip()]
    else:
        preferred = SOURCE_DEFAULT_COLUMNS.get(args.source, [
            "年份", "院校名称", "学校", "院校代码", "科类", "批次", "招生类型", "专业组", "所属专业组",
            "专业", "专业名称", "专业代码", "选科要求", "招生人数", "最低分数", "最低位次", "最低分位",
            "批次线差", "学校所在", "学校性质", "是否985", "是否211", "学费(元)", "2025招生章程", "院校主页链接",
        ])
        columns = [c for c in preferred if c in header]
    return [{c: row.get(c, "") for c in columns} for row in rows]


def output(rows: list[dict[str, str]], args: argparse.Namespace) -> None:
    if args.format == "json":
        print(json.dumps(rows, ensure_ascii=False, indent=2))
        return
    if args.format == "csv":
        if not rows:
            return
        writer = csv.DictWriter(sys.stdout, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
        return
    if not rows:
        print("No rows matched.")
        return
    headers = list(rows[0].keys())
    print("| " + " | ".join(headers) + " |")
    print("| " + " | ".join("---" for _ in headers) + " |")
    for row in rows:
        vals = []
        for h in headers:
            text = str(row.get(h, "")).replace("\n", " ").replace("|", "/")
            vals.append(text[:160])
        print("| " + " | ".join(vals) + " |")


def warn_query_risks(matches: list[dict[str, str]], args: argparse.Namespace) -> None:
    if args.school_contains and not (args.school_exact or args.school_code):
        schools = sorted({dict_get(row, "school") for row in matches if dict_get(row, "school")})
        if len(schools) > 1:
            shown = "、".join(schools[:8])
            more = "等" if len(schools) > 8 else ""
            print(
                f"WARNING: --school-contains 命中了 {len(schools)} 所院校：{shown}{more}。"
                "正式推荐或填表请改用 --school-exact 或 --school-code。",
                file=sys.stderr,
            )
    if args.source == "major-scores" and not matches:
        print(
            "WARNING: major-scores 没有匹配结果不等于该校没有相关专业。"
            "请改查 --source plans 补查专业组、专业代码、计划数、学费和备注，并在推荐清单中标记“专业分缺失”。",
            file=sys.stderr,
        )
    if args.source and args.source.startswith("art-") and not matches and args.year:
        print(
            f"WARNING: {args.source} 未匹配到年份 {args.year} 的记录。"
            "请先运行 --inspect 核对字段和可用年份；艺术类最终推荐必须以 2026 最新体育艺术版招生专业目录、合成分规则和高校章程复核。",
            file=sys.stderr,
        )
    if args.source == "charters" and matches and args.school_contains and not (args.school_exact or args.school_code):
        print(
            "WARNING: 招生章程链接应用于具体院校核验；请用 --school-exact 或 --school-code 缩小到单校。",
            file=sys.stderr,
        )


def query(args: argparse.Namespace) -> None:
    path = source_path(args.source)
    _, header, data_rows = table_rows_cached(
        path,
        effective_sheet(args.source, args.sheet),
        cache_dir=Path(args.cache_dir) if args.cache_dir else None,
        use_cache=not args.no_cache,
        rebuild_cache=args.rebuild_cache,
    )
    idx = header_index(header)
    matches = []
    for row in data_rows:
        if passes(row, idx, args):
            matches.append(row_dict(header, row))
            if len(matches) >= args.limit:
                break
    warn_query_risks(matches, args)
    output(select_columns(header, matches, args), args)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--list-sources", action="store_true")
    parser.add_argument("--inspect", action="store_true")
    parser.add_argument("--source", choices=sorted(SOURCES))
    parser.add_argument("--sheet", help="Substring of sheet name")
    parser.add_argument("--year")
    parser.add_argument("--subject")
    parser.add_argument("--batch")
    parser.add_argument("--province")
    parser.add_argument("--nature")
    parser.add_argument("--admission-type")
    parser.add_argument("--group")
    parser.add_argument("--school-code", help="Exact school code match")
    parser.add_argument("--school-exact", help="Exact school name match")
    parser.add_argument("--school-contains")
    parser.add_argument("--major-exact", help="Exact major name match")
    parser.add_argument("--major-contains")
    parser.add_argument("--keywords", help="Comma or space separated terms; all must match")
    parser.add_argument("--exclude-keywords", help="Comma or space separated terms; any match rejects row")
    parser.add_argument("--score-min", type=float)
    parser.add_argument("--score-max", type=float)
    parser.add_argument("--rank-min", type=float)
    parser.add_argument("--rank-max", type=float)
    parser.add_argument("--score-center", type=float)
    parser.add_argument("--score-width", type=float, default=0)
    parser.add_argument("--rank-center", type=float)
    parser.add_argument("--rank-width", type=float, default=0)
    parser.add_argument("--columns", help="Comma separated output columns")
    parser.add_argument("--limit", type=int, default=30)
    parser.add_argument("--format", choices=["markdown", "json", "csv"], default="markdown")
    parser.add_argument("--cache-dir", help="Parsed workbook cache directory; defaults to .cache/guangdong-gaokao under the current project")
    parser.add_argument("--no-cache", action="store_true", help="Disable parsed workbook cache for this query")
    parser.add_argument("--rebuild-cache", action="store_true", help="Rebuild the parsed workbook cache before querying")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    if args.list_sources:
        list_sources()
        return
    if not args.source:
        raise SystemExit("--source is required unless --list-sources is used")
    if args.inspect:
        inspect_source(args)
        return
    query(args)


if __name__ == "__main__":
    main()
