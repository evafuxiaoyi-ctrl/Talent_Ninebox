from __future__ import annotations

import re
import shutil
import zipfile
import copy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

HEADER_ROW = 3
DATA_START_ROW = HEADER_ROW + 1
MAX_SPLIT_GROUPS = 200
ALLOWED_SPLIT_FIELDS = [
    "员工姓名",
    "工号",
    "邮箱",
    "岗位",
    "一级部门",
    "二级部门",
    "三级部门",
    "四级部门",
]


@dataclass(frozen=True)
class SplitField:
    name: str
    column_index: int
    column_letter: str


@dataclass(frozen=True)
class SplitResult:
    output_file: Path
    summary: dict[str, object]


def _clean_header(value: object) -> str:
    return str(value or "").strip()


def _safe_filename_part(value: object) -> str:
    text = str(value or "").strip() or "未填写"
    text = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text[:60] or "未填写"


def _find_template_sheet(workbook) -> object:
    for sheet in workbook.worksheets:
        headers = {_clean_header(cell.value) for cell in sheet[HEADER_ROW]}
        if any(field in headers for field in ALLOWED_SPLIT_FIELDS):
            return sheet
    return workbook.active


def list_split_fields(workbook_path: Path) -> list[SplitField]:
    try:
        workbook = load_workbook(workbook_path, read_only=True, data_only=False)
    except Exception as exc:
        raise ValueError("文件无法读取，可能已损坏、加密，或并非有效的 .xlsx 文件。") from exc
    try:
        sheet = _find_template_sheet(workbook)
        fields: list[SplitField] = []
        seen: set[str] = set()
        for cell in sheet[HEADER_ROW]:
            header = _clean_header(cell.value)
            if header in ALLOWED_SPLIT_FIELDS and header not in seen:
                fields.append(SplitField(header, cell.column, get_column_letter(cell.column)))
                seen.add(header)
        if not fields:
            raise ValueError("第 3 行未找到可用于拆分的字段，请确认模板包含员工姓名、工号、邮箱、岗位或部门字段。")
        return fields
    finally:
        workbook.close()


def _group_values(workbook_path: Path, field_name: str) -> tuple[str, int, dict[str, int]]:
    workbook = load_workbook(workbook_path, read_only=True, data_only=False)
    try:
        sheet = _find_template_sheet(workbook)
        field_column = None
        for cell in sheet[HEADER_ROW]:
            if _clean_header(cell.value) == field_name:
                field_column = cell.column
                break
        if field_column is None:
            raise ValueError(f"第 3 行未找到拆分字段「{field_name}」。")

        groups: dict[str, int] = {}
        for row in range(DATA_START_ROW, sheet.max_row + 1):
            row_values = [sheet.cell(row=row, column=column).value for column in range(1, sheet.max_column + 1)]
            if not any(value not in (None, "") for value in row_values):
                continue
            value = str(sheet.cell(row=row, column=field_column).value or "").strip() or "未填写"
            groups[value] = groups.get(value, 0) + 1

        if not groups:
            raise ValueError("未找到可拆分的数据行，请确认第 4 行开始存在人员数据。")
        if len(groups) > MAX_SPLIT_GROUPS:
            raise ValueError(f"拆分后将生成 {len(groups)} 个文件，超过 {MAX_SPLIT_GROUPS} 个上限，请更换拆分字段。")
        return sheet.title, field_column, groups
    finally:
        workbook.close()


def _copy_cell(source, target, source_row: int, target_row: int) -> None:
    value = source.value
    if isinstance(value, str) and value.startswith("=") and source_row != target_row:
        try:
            value = Translator(value, origin=f"{get_column_letter(source.column)}{source_row}").translate_formula(
                f"{get_column_letter(target.column)}{target_row}"
            )
        except Exception:
            pass
    target.value = value
    if source.has_style:
        target.font = copy.copy(source.font)
        target.fill = copy.copy(source.fill)
        target.border = copy.copy(source.border)
        target.alignment = copy.copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy.copy(source.protection)
    if source.hyperlink:
        target._hyperlink = copy.copy(source.hyperlink)
    if source.comment:
        target.comment = copy.copy(source.comment)


def _write_group_workbook(source_path: Path, target_path: Path, sheet_name: str, field_column: int, group_value: str) -> int:
    source_workbook = load_workbook(source_path, data_only=False)
    target_workbook = load_workbook(target_path, data_only=False)
    try:
        source_sheet = source_workbook[sheet_name]
        target_sheet = target_workbook[sheet_name]
        kept = 0
        for source_row in range(DATA_START_ROW, source_sheet.max_row + 1):
            row_values = [source_sheet.cell(row=source_row, column=column).value for column in range(1, source_sheet.max_column + 1)]
            if not any(value not in (None, "") for value in row_values):
                continue
            value = str(source_sheet.cell(row=source_row, column=field_column).value or "").strip() or "未填写"
            if value != group_value:
                continue
            target_row = DATA_START_ROW + kept
            for column in range(1, source_sheet.max_column + 1):
                _copy_cell(source_sheet.cell(row=source_row, column=column), target_sheet.cell(row=target_row, column=column), source_row, target_row)
            if source_row in source_sheet.row_dimensions:
                target_sheet.row_dimensions[target_row].height = source_sheet.row_dimensions[source_row].height
            kept += 1

        first_extra_row = DATA_START_ROW + kept
        if first_extra_row <= target_sheet.max_row:
            target_sheet.delete_rows(first_extra_row, target_sheet.max_row - first_extra_row + 1)

        target_workbook.save(target_path)
        return kept
    finally:
        source_workbook.close()
        target_workbook.close()


def split_workbook_by_field(workbook_path: Path, output_dir: Path, field_name: str) -> SplitResult:
    available_fields = {field.name for field in list_split_fields(workbook_path)}
    if field_name not in ALLOWED_SPLIT_FIELDS:
        raise ValueError("请选择系统支持的拆分字段。")
    if field_name not in available_fields:
        raise ValueError(f"当前模板第 3 行不包含拆分字段「{field_name}」。")

    sheet_name, field_column, groups = _group_values(workbook_path, field_name)
    split_dir = output_dir / "split_excels"
    split_dir.mkdir(parents=True, exist_ok=True)

    generated: list[dict[str, object]] = []
    for index, group_value in enumerate(sorted(groups.keys()), start=1):
        filename = f"{index:02d}_{_safe_filename_part(field_name)}_{_safe_filename_part(group_value)}.xlsx"
        target_path = split_dir / filename
        shutil.copy2(workbook_path, target_path)
        kept = _write_group_workbook(workbook_path, target_path, sheet_name, field_column, group_value)
        generated.append({"字段值": group_value, "人数": kept, "文件名": filename})

    output_file = output_dir / f"人才盘点拆分结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    with zipfile.ZipFile(output_file, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for item in generated:
            archive.write(split_dir / str(item["文件名"]), arcname=str(item["文件名"]))

    summary = {
        "处理类型": "表格拆分",
        "拆分依据": field_name,
        "拆分 Sheet": sheet_name,
        "生成文件数": len(generated),
        "总人员数": sum(int(item["人数"]) for item in generated),
        "拆分明细": generated,
    }
    return SplitResult(output_file=output_file, summary=summary)
