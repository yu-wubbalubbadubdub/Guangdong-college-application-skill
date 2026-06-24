#!/usr/bin/env python3
"""Render confirmed recommendations into an official-form-style volunteer draft.

Normal .xlsx output is created by copying the blank template in assets/ and
replacing the editable worksheets. Use --create-template only when maintaining
the skill asset itself.
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_TEMPLATE = ROOT / "assets" / "2026广东高考志愿表填报模板.xlsx"
OFFICIAL_PDF = "assets/2026年广东省普通高校招生考生志愿表（官方）.pdf"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

FIELDS = [
    "录取批次",
    "志愿号",
    "院校代码",
    "院校名称",
    "院校专业组代码",
    "专业代码1",
    "专业代码2",
    "专业代码3",
    "专业代码4",
    "专业代码5",
    "专业代码6",
    "是否服从调剂",
    "层级",
    "备注",
]

REMINDERS = [
    "本草稿只是志愿填报辅助结果，不等于已在广东省教育考试院系统提交。",
    "普通类本科院校志愿应在第二时段填报：2026-06-29 19:00 至 2026-07-04 16:00。",
    "第二时段截止后的确认窗口为 2026-07-04 16:00-18:00；该窗口只能确认，不能修改。",
    "广东普通类本科院校栏目最多可填 45 个院校专业组志愿；空白行可继续编辑补充。",
    "正式填报前必须用 2026 最新招生专业目录核对院校代码、院校专业组代码、专业代码和招生计划。",
    "正式填报前必须核对高校招生章程中的体检、色觉、单科、语种、收费、转专业和调剂范围限制。",
    "志愿最终以广东省教育考试院系统中最后一次网上确认版本为准。",
]

SHEET1_WIDTHS = [16, 8, 12, 22, 16, 11, 11, 11, 11, 11, 11, 12, 10, 58]
SHEET2_WIDTHS = [18, 120]
REQUIRED_FORM_FIELDS = ["录取批次", "志愿号", "院校代码", "院校名称", "院校专业组代码"]


def load_rows(path: Path) -> list[dict[str, str]]:
    if path.suffix.lower() == ".json":
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        if isinstance(data, dict):
            data = data.get("rows") or data.get("recommendations") or []
        return [dict(item) for item in data]
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return [dict(row) for row in csv.DictReader(handle)]


def get(row: dict[str, str], *names: str) -> str:
    for name in names:
        value = row.get(name)
        if value not in (None, ""):
            return str(value)
    return ""


def split_major_codes(value: str) -> list[str]:
    normalized = (
        value.replace("；", ",")
        .replace(";", ",")
        .replace("、", ",")
        .replace("，", ",")
        .replace("/", ",")
    )
    return [part.strip() for part in normalized.split(",") if part.strip()]


def normalize(row: dict[str, str], index: int) -> dict[str, str]:
    majors = get(row, "专业代码", "专业代码列表", "major_codes")
    codes = split_major_codes(majors)
    out = {
        "录取批次": get(row, "录取批次", "批次", "batch") or "普通类本科院校",
        "志愿号": get(row, "志愿号", "志愿序号", "order") or str(index),
        "院校代码": get(row, "院校代码", "school_code"),
        "院校名称": get(row, "院校名称", "学校", "school"),
        "院校专业组代码": get(
            row,
            "院校专业组代码",
            "专业组代码",
            "专业组",
            "所属专业组",
            "group",
        ),
        "是否服从调剂": get(row, "是否服从调剂", "服从调剂", "adjustment") or "待确认",
        "层级": get(row, "层级", "risk_level"),
        "备注": get(row, "备注", "推荐理由", "风险提示", "note"),
    }
    for i in range(6):
        out[f"专业代码{i + 1}"] = (
            get(row, f"专业代码{i + 1}", f"major_code_{i + 1}")
            or (codes[i] if i < len(codes) else "")
        )
    return out


def render_markdown(rows: list[dict[str, str]]) -> str:
    lines = [
        "# 2026广东省普通高校招生考生志愿表草稿",
        "",
        f"- 官方附件：{OFFICIAL_PDF}",
        f"- 可编辑模板：assets/{DEFAULT_TEMPLATE.name}",
        "- 提醒：本草稿不替代广东省教育考试院系统中的最终填报和确认。",
        "",
        "| " + " | ".join(FIELDS) + " |",
        "| " + " | ".join("---" for _ in FIELDS) + " |",
    ]
    for row in rows:
        lines.append("| " + " | ".join(str(row.get(field, "")).replace("|", "/") for field in FIELDS) + " |")
    lines.extend(["", "## 最终确认提醒", ""])
    lines.extend(f"- {item}" for item in REMINDERS)
    warnings = validation_warnings(rows)
    if warnings:
        lines.extend(["", "## 草稿复核警告", ""])
        lines.extend(f"- {item}" for item in warnings)
    return "\n".join(lines) + "\n"


def blank_rows(count: int, batch: str = "普通类本科院校") -> list[dict[str, str]]:
    rows = []
    for index in range(1, count + 1):
        row = {field: "" for field in FIELDS}
        row["录取批次"] = batch
        row["志愿号"] = str(index)
        rows.append(row)
    return rows


def pad_rows(rows: list[dict[str, str]], count: int) -> list[dict[str, str]]:
    if count <= 0 or len(rows) >= count:
        return rows
    batch = rows[0].get("录取批次", "普通类本科院校") if rows else "普通类本科院校"
    out = list(rows)
    for index in range(len(rows) + 1, count + 1):
        row = {field: "" for field in FIELDS}
        row["录取批次"] = batch
        row["志愿号"] = str(index)
        row["是否服从调剂"] = "待确认"
        row["备注"] = "待补充：请结合2026招生专业目录、用户偏好和梯度继续完善。"
        out.append(row)
    return out


def has_recommendation(row: dict[str, str]) -> bool:
    return bool(row.get("院校名称") or row.get("院校代码") or row.get("院校专业组代码"))


def validation_warnings(rows: list[dict[str, str]]) -> list[str]:
    warnings = []
    for row in rows:
        if not has_recommendation(row):
            continue
        label = f"志愿{row.get('志愿号') or '?'} {row.get('院校名称') or row.get('院校代码') or '未命名院校'}"
        missing = [field for field in REQUIRED_FORM_FIELDS if not row.get(field)]
        if missing:
            warnings.append(f"{label} 缺少官方填报必需字段：{'、'.join(missing)}。正式填报前必须补齐。")
        major_codes = [row.get(f"专业代码{i}") for i in range(1, 7)]
        if not any(major_codes):
            warnings.append(f"{label} 未填写任何专业代码；不得凭历史备注猜代码，必须以 2026 招生专业目录补齐。")
    return warnings[:30]


def col_name(index: int) -> str:
    name = ""
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def cell_xml(row: int, col: int, value: str, style: int = 0) -> str:
    attrs = f' r="{col_name(col)}{row}" t="inlineStr"'
    if style:
        attrs += f' s="{style}"'
    text = escape(str(value or ""))
    return f"<c{attrs}><is><t>{text}</t></is></c>"


def sheet_xml(rows: list[list[str]], widths: list[float] | None = None) -> str:
    cols = ""
    if widths:
        cols = "<cols>" + "".join(
            f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>'
            for i, width in enumerate(widths, 1)
        ) + "</cols>"
    body = []
    for r_idx, row in enumerate(rows, 1):
        cells = "".join(
            cell_xml(r_idx, c_idx, value, 1 if r_idx == 1 else 0)
            for c_idx, value in enumerate(row, 1)
        )
        body.append(f'<row r="{r_idx}">{cells}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        f"{cols}<sheetData>{''.join(body)}</sheetData></worksheet>"
    )


def workbook_parts(rows: list[dict[str, str]]) -> dict[str, str]:
    table = [FIELDS] + [[row.get(field, "") for field in FIELDS] for row in rows]
    reminders = [["项目", "内容"], ["官方附件", OFFICIAL_PDF], ["可编辑空白模板", f"assets/{DEFAULT_TEMPLATE.name}"]]
    reminders.extend([["最终确认提醒", item] for item in REMINDERS])
    reminders.extend([["草稿复核警告", item] for item in validation_warnings(rows)])
    reminders.extend(
        [
            ["使用说明", "请直接编辑“志愿表草稿”工作表中的空白志愿位。"],
            ["数据限制", "历史数据仅作辅助，最终以2026招生专业目录和高校章程为准。"],
        ]
    )
    return {
        "xl/worksheets/sheet1.xml": sheet_xml(table, SHEET1_WIDTHS),
        "xl/worksheets/sheet2.xml": sheet_xml(reminders, SHEET2_WIDTHS),
    }


def new_xlsx_package(rows: list[dict[str, str]]) -> dict[str, str]:
    files = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            "</Types>"
        ),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>"
        ),
        "xl/workbook.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="志愿表草稿" sheetId="1" r:id="rId1"/><sheet name="填报提醒" sheetId="2" r:id="rId2"/></sheets>'
            "</workbook>"
        ),
        "xl/_rels/workbook.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            "</Relationships>"
        ),
        "xl/styles.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="2"><font><name val="Arial"/></font><font><b/><name val="Arial"/></font></fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/></patternFill></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            '<xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>'
            "</styleSheet>"
        ),
    }
    files.update(workbook_parts(rows))
    return files


def write_blank_template(output: Path, count: int) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, content in new_xlsx_package(blank_rows(count)).items():
            zf.writestr(name, content)


def fill_template_xlsx(rows: list[dict[str, str]], template: Path, output: Path) -> None:
    if not template.exists():
        raise SystemExit(f"找不到 Excel 空白模板：{template}")
    if template.resolve() == output.resolve():
        raise SystemExit("输出路径不能与空白模板路径相同，请写入 output/ 下的新文件。")
    output.parent.mkdir(parents=True, exist_ok=True)
    replacements = workbook_parts(rows)
    try:
        with zipfile.ZipFile(template, "r") as source, zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as target:
            replaced = set()
            for item in source.infolist():
                content = replacements.get(item.filename)
                if content is None:
                    content = source.read(item.filename)
                target.writestr(item, content)
                replaced.add(item.filename)
            for name, content in replacements.items():
                if name not in replaced:
                    target.writestr(name, content)
    except PermissionError as exc:
        raise SystemExit(f"无法写入输出文件，可能已被 Excel 或预览程序占用：{output}") from exc


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", help="Confirmed recommendations as JSON or CSV")
    parser.add_argument("--output", help="Output markdown or xlsx path")
    parser.add_argument("--pad-to", type=int, default=0, help="Pad volunteer rows to this count with editable blanks")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="Blank xlsx template to copy and fill")
    parser.add_argument("--create-template", help="Create or refresh a blank xlsx template asset at this path")
    args = parser.parse_args()

    if args.create_template:
        count = args.pad_to if args.pad_to > 0 else 45
        write_blank_template(Path(args.create_template), count)
        print(Path(args.create_template))
        return

    if not args.input:
        raise SystemExit("--input is required unless --create-template is used")

    raw_rows = load_rows(Path(args.input))
    rows = [normalize(row, i + 1) for i, row in enumerate(raw_rows)]
    rows = pad_rows(rows, args.pad_to)

    output = Path(args.output) if args.output else ROOT / "output" / "2026广东高考志愿表草稿.md"
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.suffix.lower() == ".xlsx":
        fill_template_xlsx(rows, Path(args.template), output)
    else:
        output.write_text(render_markdown(rows), encoding="utf-8")
    print(output)


if __name__ == "__main__":
    main()
