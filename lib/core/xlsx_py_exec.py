# pyright: reportMissingModuleSource=false
"""openpyxl 写入/编辑执行器(由 xlsx-py.js 通过 stdin 的 JSON 调用)。

只做机械动作:JS 侧已经把值编码成带类型标签({k, v})、把样式与规则校验成普通 JSON。
落盘前会重新打开产物自校验,通过后 os.replace 到目标,失败时目标文件不受影响。
stdout 只回一行 JSON:成功 {"ok": true, ...},失败 {"ok": false, "error": ..., "kind": ...}。
"""

from __future__ import annotations

import datetime
import json
import os
import sys
import zipfile
from pathlib import Path
from typing import Any, Callable, cast

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule, IconSetRule, Rule
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

MAX_NAME = 31
THIN = Side(style="thin", color=Color(rgb="FF000000"))


class OpError(Exception):
    """单个操作失败(optional 时可跳过)。"""


# ---------------------------------------------------------------------------
# 单元格值与样式
# ---------------------------------------------------------------------------
def cell_value(cell: dict[str, Any]) -> Any:
    """带标签的单元格 -> Python 值。"""
    kind = cell.get("k")
    value = cell.get("v")
    if kind == "n":
        return value
    if kind == "b":
        return bool(value)
    if kind == "f":
        return "=" + str(value)
    if kind == "d":
        return datetime.datetime.fromisoformat(str(value))
    return "" if value is None else str(value)


def _merged_font(cell, style: dict[str, Any]) -> Font | None:
    current = cell.font
    changed = False
    font = Font(
        name=current.name,
        sz=current.sz,
        b=bool(style.get("bold", current.b)),
        i=bool(style.get("italic", current.i)),
        u=current.u,
    )
    if "size" in style:
        font.sz = float(style["size"])
        changed = True
    if "color" in style:
        font.color = Color(rgb="FF" + str(style["color"]))
        changed = True
    if changed or "bold" in style or "italic" in style:
        return font
    return None


def apply_style(cell, style: dict[str, Any] | None) -> None:
    """把规范化样式落到单元格上与已有格式合并(不整体覆盖其他属性)。"""
    if not style:
        return
    font = _merged_font(cell, style)
    if font is not None:
        cell.font = font
    if "fill" in style:
        cell.fill = PatternFill(fill_type="solid", start_color=Color(rgb="FF" + str(style["fill"])))
    if "numFmt" in style:
        cell.number_format = str(style["numFmt"])
    if "borderColor" in style or "border" in style:
        color = Color(rgb="FF" + str(style.get("borderColor", "000000")))
        side = Side(style="thin", color=color)
        cell.border = Border(left=side, right=side, top=side, bottom=side)
    align = Alignment(
        horizontal=style.get("align", cell.alignment.horizontal),
        vertical=style.get("valign", cell.alignment.vertical),
        wrap_text=bool(style.get("wrap", cell.alignment.wrap_text)),
    )
    if any(key in style for key in ("align", "valign", "wrap")):
        cell.alignment = align


def display_width(value: Any) -> int:
    """估算列宽:中文按两个字符宽(入参可能是带标签的单元格)。"""
    inner = value.get("v") if isinstance(value, dict) else value
    text = "" if inner is None else str(inner)
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in text)


# ---------------------------------------------------------------------------
# 工作表定位与写入
# ---------------------------------------------------------------------------
def at(sheet: Any, ref: str) -> Any:
    """单个单元格:`sheet['A1']` 的静态类型是 Cell|tuple,这里收敛成单格。"""
    return sheet[ref]


def sheet_of(wb: Workbook, selector: Any):
    if isinstance(selector, int):
        sheets = wb.worksheets
        if 1 <= selector <= len(sheets):
            return sheets[selector - 1]
        raise OpError(f"工作表序号 {selector} 不存在（共 {len(sheets)} 个，序号从 1 开始）")
    name = str(selector)
    if name in wb.sheetnames:
        return wb[name]
    raise OpError(f"工作表「{name}」不存在（现有: {'、'.join(wb.sheetnames)}）")


def write_rows(sheet, start_row: int, start_col: int, rows: list) -> None:
    for r, row in enumerate(rows):
        if row is None:
            continue
        for c, cell in enumerate(row):
            if not cell:
                continue
            target = sheet.cell(row=start_row + r, column=start_col + c)
            target.value = cell_value(cell)


def apply_cell_styles(sheet, rows: list) -> None:
    """逐格样式:单元格自带 style 时落到对应格。"""
    for r, row in enumerate(rows):
        for c, cell in enumerate(row or []):
            if cell and cell.get("style"):
                apply_style(sheet.cell(row=1 + r, column=1 + c), cell["style"])


def apply_header_style(sheet, style: Any, width: int) -> None:
    """首行整行套表头样式(加粗底纹等)。"""
    for c in range(width):
        apply_style(sheet.cell(row=1, column=1 + c), style)


def fill_sheet(sheet, spec: dict[str, Any]) -> None:
    """按 spec 写一张表:值、逐格样式、表头美化、列宽行高。"""
    rows = spec.get("rows") or []
    write_rows(sheet, 1, 1, rows)
    apply_cell_styles(sheet, rows)
    width = max((len(row) for row in rows), default=0)
    if spec.get("header") and rows:
        apply_header_style(sheet, spec.get("headerStyle"), width)
    apply_widths(sheet, spec.get("columnWidths"), rows, width)
    for index, height in enumerate(spec.get("rowHeights") or []):
        if height:
            sheet.row_dimensions[index + 1].height = float(height)


def apply_widths(sheet, widths: Any, rows: list, width: int) -> None:
    if isinstance(widths, list):
        for index, value in enumerate(widths):
            if isinstance(value, (int, float)) and 0 < value <= 255:
                sheet.column_dimensions[get_column_letter(index + 1)].width = value
        return
    # 没给列宽时按内容估算,保证中文可见(与 ExcelJS 后端的下限/上限一致)
    for c in range(width):
        longest = max((display_width((row or [None] * (c + 1))[c] if c < len(row or []) else "") for row in rows), default=0)
        sheet.column_dimensions[get_column_letter(c + 1)].width = max(10, min(50, longest + 2))


# ---------------------------------------------------------------------------
# 编辑操作
# ---------------------------------------------------------------------------
def op_set_value(wb, op) -> None:
    cell = at(sheet_of(wb, op.get("sheet", 1)), op["ref"])
    cell.value = cell_value(op.get("cell") or {"k": "n", "v": None})
    apply_style(cell, op.get("style"))


def op_set_formula(wb, op) -> None:
    cell = at(sheet_of(wb, op.get("sheet", 1)), op["ref"])
    cell.value = "=" + str(op.get("formula", ""))
    apply_style(cell, op.get("style"))


def op_set_cells(wb, op) -> None:
    sheet = sheet_of(wb, op.get("sheet", 1))
    start = at(sheet, op.get("start", "A1"))
    write_rows(sheet, start.row, start.column, op.get("values") or [])
    if op.get("style"):
        for r, row in enumerate(op.get("values") or []):
            for c in range(len(row or [])):
                apply_style(sheet.cell(row=start.row + r, column=start.column + c), op["style"])


def op_style_range(wb, op) -> None:
    sheet = sheet_of(wb, op.get("sheet", 1))
    ref = str(op.get("range") or "")
    if not ref:
        raise OpError("style_range 需要 range（如 A1:C10）")
    cells = [at(sheet, ref)] if ":" not in ref else [cell for row in sheet[ref] for cell in row]
    for cell in cells:
        apply_style(cell, op.get("style"))


def op_merge(wb, op) -> None:
    sheet = sheet_of(wb, op.get("sheet", 1))
    ref = str(op.get("range") or "")
    if not ref:
        raise OpError("merge/unmerge 需要 range（如 A1:C3）")
    if op.get("op") == "merge":
        sheet.merge_cells(ref)
    else:
        sheet.unmerge_cells(ref)


def op_resize(wb, op) -> None:
    sheet = sheet_of(wb, op.get("sheet", 1))
    at = int(op.get("at", 1))
    count = max(1, int(op.get("count", 1)))
    kind = op.get("op")
    if kind == "insert_rows":
        sheet.insert_rows(at, count)
    elif kind == "delete_rows":
        sheet.delete_rows(at, count)
    elif kind == "insert_cols":
        sheet.insert_cols(at, count)
    else:
        sheet.delete_cols(at, count)


def op_add_sheet(wb, op) -> None:
    spec = op.get("sheet") or {}
    name = str(spec.get("name") or "")[:MAX_NAME]
    if name in wb.sheetnames:
        raise OpError(f"工作表「{name}」已存在")
    fill_sheet(wb.create_sheet(title=name), spec)


def op_rename_sheet(wb, op) -> None:
    sheet = sheet_of(wb, op.get("sheet", 1))
    name = str(op.get("name"))[:MAX_NAME]
    if name in wb.sheetnames and wb[name] is not sheet:
        raise OpError(f"工作表「{name}」已存在")
    sheet.title = name


def op_delete_sheet(wb, op) -> None:
    sheet = sheet_of(wb, op.get("sheet", 1))
    if len(wb.worksheets) <= 1:
        raise OpError("不能删除最后一个工作表")
    wb.remove(sheet)


def op_set_col_width(wb, op) -> None:
    sheet = sheet_of(wb, op.get("sheet", 1))
    col = op.get("col")
    letters = get_column_letter(col) if isinstance(col, int) else str(col).upper()
    sheet.column_dimensions[letters].width = float(op.get("width", 8.43))


def op_set_row_height(wb, op) -> None:
    sheet = sheet_of(wb, op.get("sheet", 1))
    sheet.row_dimensions[int(op.get("row", 1))].height = float(op.get("height", 15))


def op_freeze(wb, op) -> None:
    sheet = sheet_of(wb, op.get("sheet", 1))
    rows = int(op.get("rows") or 0)
    cols = int(op.get("cols") or 0)
    sheet.freeze_panes = f"{get_column_letter(cols + 1)}{rows + 1}" if rows or cols else None


def op_auto_filter(wb, op) -> None:
    sheet_of(wb, op.get("sheet", 1)).auto_filter.ref = str(op.get("range") or "")


def op_add_image(wb, op) -> None:
    sheet = sheet_of(wb, op.get("sheet", 1))
    path = Path(str(op.get("path") or ""))
    if not path.is_file():
        raise OpError(f"图片不存在: {path}")
    try:
        image = Image(str(path))
    except ImportError as err:
        raise OpError("add_image 需要 Pillow（pip3 install Pillow）") from err
    except OSError as err:
        raise OpError(f"图片无法读取: {err}") from err
    width, height = int(op.get("width") or 0), int(op.get("height") or 0)
    if width > 0:
        image.width, image.height = width, height or width
    anchor = op.get("cell") or str(op.get("range") or "A1").split(":")[0]
    sheet.add_image(image, anchor)


def _dxf(style: Any) -> DifferentialStyle | None:
    if not style:
        return None
    out = DifferentialStyle(
        font=Font(bold=bool(style.get("bold")), italic=bool(style.get("italic")),
                  size=float(style["size"]) if style.get("size") else None,
                  color=Color(rgb="FF" + str(style["color"])) if style.get("color") else None),
    )
    if style.get("fill"):
        out.fill = PatternFill(fill_type="solid", bgColor=Color(rgb="FF" + str(style["fill"])))
    if style.get("numFmt"):
        out.numFmt = NumberFormat(numFmtId=0, formatCode=str(style["numFmt"]))
    if style.get("borderColor"):
        color = Color(rgb="FF" + str(style["borderColor"]))
        side = Side(style="thin", color=color)
        out.border = Border(left=side, right=side, top=side, bottom=side)
    return out


def _with_dxf(rule: Any, style: Any) -> Any:
    """FormulaRule/CellIsRule 的 dxf 参数只认 font/fill/border,数字格式得自己挂上去。"""
    dxf = _dxf(style)
    if dxf is not None:
        rule.dxf = dxf
    return rule


def _rule(kind: str, rule: dict[str, Any]) -> Any:
    style = rule.get("dxf")
    if kind == "cellIs":
        return _with_dxf(CellIsRule(operator=rule["operator"], formula=rule["formula"]), style)
    if kind == "expression":
        return _with_dxf(FormulaRule(formula=rule["formula"]), style)
    if kind == "colorScale":
        stops = rule["stops"]
        colors = ["FF" + color for color in rule["colors"]]
        scale = ColorScaleRule(start_type=stops[0]["type"], start_color=colors[0],
                               end_type=stops[-1]["type"], end_color=colors[-1])
        if len(colors) == 3:
            scale = ColorScaleRule(start_type=stops[0]["type"], start_color=colors[0],
                                   mid_type=stops[1]["type"], mid_value=stops[1].get("value"), mid_color=colors[1],
                                   end_type=stops[-1]["type"], end_color=colors[-1])
        return scale
    if kind == "dataBar":
        return DataBarRule(start_type="min", end_type="max", color="FF" + str(rule.get("color")))
    if kind == "iconSet":
        icon = str(rule.get("iconSet"))
        return IconSetRule(icon, "percent", rule["stops"], showValue=bool(rule.get("showValue", True)))
    if kind == "top10":
        return Rule(type="top10", rank=int(rule.get("rank", 10)), percent=bool(rule.get("percent")),
                    bottom=bool(rule.get("bottom")), dxf=_dxf(style))
    if kind == "aboveAverage":
        return Rule(type="aboveAverage", aboveAverage=bool(rule.get("aboveAverage", True)), dxf=_dxf(style))
    return Rule(type="timePeriod", timePeriod=cast(Any, str(rule.get("timePeriod"))))


def op_conditional_format(wb, op) -> None:
    sheet = sheet_of(wb, op.get("sheet", 1))
    for rule in op.get("rules") or []:
        sheet.conditional_formatting.add(str(op["range"]), _rule(str(rule["kind"]), rule))


def op_data_validation(wb, op) -> None:
    sheet = sheet_of(wb, op.get("sheet", 1))
    rule = op.get("validation") or {}
    validation = DataValidation(
        type=rule.get("type"),
        formula1=rule.get("formula1"),
        formula2=rule.get("formula2"),
        allowBlank=bool(rule.get("allowBlank")),
        operator=rule.get("operator"),
        showErrorMessage=bool(rule.get("showErrorMessage")),
        showInputMessage=bool(rule.get("showInputMessage")),
        errorTitle=rule.get("errorTitle"),
        error=rule.get("error"),
        errorStyle=rule.get("errorStyle"),
        promptTitle=rule.get("promptTitle"),
        prompt=rule.get("prompt"),
    )
    sheet.add_data_validation(validation)
    validation.add(str(op["range"]))


OPS: dict[str, Callable] = {
    "set_value": op_set_value,
    "set_formula": op_set_formula,
    "set_cells": op_set_cells,
    "style_range": op_style_range,
    "merge": op_merge,
    "unmerge": op_merge,
    "insert_rows": op_resize,
    "delete_rows": op_resize,
    "insert_cols": op_resize,
    "delete_cols": op_resize,
    "add_sheet": op_add_sheet,
    "rename_sheet": op_rename_sheet,
    "delete_sheet": op_delete_sheet,
    "set_col_width": op_set_col_width,
    "set_row_height": op_set_row_height,
    "freeze": op_freeze,
    "auto_filter": op_auto_filter,
    "add_image": op_add_image,
    "conditional_format": op_conditional_format,
    "data_validation": op_data_validation,
}


def run_op(wb, op: dict[str, Any]) -> dict[str, Any]:
    """执行一条操作;optional 操作失败时改成跳过记录(带原始顺序号)。"""
    kind = str(op.get("op"))
    handler = OPS.get(kind)
    if handler is None:
        raise OpError(f"未知操作类型: {kind}")
    try:
        handler(wb, op)
        text = kind
    except OpError as err:
        if not op.get("optional"):
            raise
        text = f"{kind}(跳过: {err})"
    return {"index": op.get("changeIndex"), "text": text}


# ---------------------------------------------------------------------------
# 自校验
# ---------------------------------------------------------------------------
def sheet_shape(wb, name: str) -> dict[str, Any]:
    sheet = wb[name]
    return {"name": name, "rows": max(1, sheet.max_row), "cols": max(1, sheet.max_column)}


def check_names(actual: list[str], expected: list[str]) -> None:
    if actual == expected:
        return
    raise VerifyError(f"工作表与预期不符(实际 {'/'.join(actual)},预期 {'/'.join(expected)})")


def check_build(actual: list[dict[str, Any]], expected: list[dict[str, Any]]) -> None:
    check_names([s["name"] for s in actual], [s["name"] for s in expected])
    for got, want in zip(actual, expected):
        if got["rows"] < want["rows"] or got["cols"] < want["cols"]:
            raise VerifyError(f"工作表「{want['name']}」落盘不全(期望至少 {want['rows']}×{want['cols']},实际 {got['rows']}×{got['cols']})")


def check_written(wb, written: list[dict[str, Any]]) -> None:
    for entry in written:
        # 名字可能已被后续改名,退回到登记时的序号;整张表被删掉则跳过
        if entry["sheet"] in wb.sheetnames:
            sheet = wb[entry["sheet"]]
        elif int(entry.get("index", 0)) <= len(wb.worksheets):
            sheet = wb.worksheets[int(entry["index"]) - 1]
        else:
            continue
        if "rows" in entry:
            start = at(sheet, entry.get("start", "A1"))
            refs = [sheet.cell(row=start.row + r, column=start.column + c)
                    for r, row in enumerate(entry["rows"]) for c in range(len(row or [])) if (row or [])[c]]
        else:
            refs = [at(sheet, entry["ref"])]
        for cell in refs:
            if cell.value is None:
                raise VerifyError(f"回读校验失败:{cell.coordinate} 仍是空单元格")


class VerifyError(Exception):
    """产物回读与预期不符。"""


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def save_atomic(wb: Workbook, target: Path) -> None:
    temp = target.with_suffix(".tmp.xlsx")
    wb.save(temp)
    os.replace(temp, target)


def has_part(path: Path, prefix: str) -> bool:
    """包里是否含某类部件(用来提示 openpyxl 会丢图表)。"""
    with zipfile.ZipFile(path) as package:
        return any(name.startswith(prefix) for name in package.namelist())


def build(payload: dict[str, Any]) -> dict[str, Any]:
    wb = Workbook()
    if wb.active is not None:
        wb.remove(wb.active)
    for spec in payload.get("sheets") or []:
        fill_sheet(wb.create_sheet(title=spec["name"]), spec)
    target = Path(payload["output"])
    save_atomic(wb, target)
    opened = load_workbook(target)
    actual = [sheet_shape(opened, name) for name in opened.sheetnames]
    check_build(actual, payload["expect"]["sheets"])
    return {"ok": True, "output": str(target), "sheets": actual}


def edit(payload: dict[str, Any]) -> dict[str, Any]:
    source = Path(payload["input"])
    charts = has_part(source, "xl/charts/")
    try:
        opened = load_workbook(source)
    except ImportError as err:
        raise OpError("该工作簿含图片，openpyxl 处理图片需要 Pillow（pip3 install Pillow）") from err
    changes = [run_op(opened, op) for op in payload.get("ops") or []]
    target = Path(payload["output"])
    save_atomic(opened, target)
    check = load_workbook(target)
    check_names(check.sheetnames, payload["expect"]["sheets"])
    check_written(check, payload["expect"].get("written") or [])
    sheets = [sheet_shape(check, name) for name in check.sheetnames]
    return {"ok": True, "output": str(target), "changes": changes, "sheets": sheets, "droppedCharts": charts}


def main() -> int:
    try:
        payload = json.loads(sys.stdin.read() or "{}")
        result = edit(payload) if payload.get("mode") == "edit" else build(payload)
    except VerifyError as err:
        print(json.dumps({"ok": False, "kind": "verify", "error": str(err)}, ensure_ascii=False))
        return 2
    except OpError as err:
        print(json.dumps({"ok": False, "kind": "op", "error": str(err)}, ensure_ascii=False))
        return 2
    except Exception as err:  # NOSONAR
        print(json.dumps({"ok": False, "kind": "op", "error": f"{type(err).__name__}: {err}"}, ensure_ascii=False))
        return 2
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
