from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from .field_utils import has_field, has_placement_field
from .models import Issue, ProcessOptions, WorkbookInfo

REQUIRED_HEADERS = ("姓名", "人才九宫格落位")


def normalize_header(value: object) -> str:
    return str(value).strip() if value is not None else ""


def inspect_workbook(source_name: str, path: Path, options: ProcessOptions) -> tuple[WorkbookInfo | None, list[Issue]]:
    issues: list[Issue] = []
    try:
        wb = load_workbook(path, data_only=False)
    except Exception as exc:
        return None, [Issue("error", "文件读取异常", f"文件无法读取：{exc}", source_file=source_name)]

    for ws in wb.worksheets:
        max_row = min(options.max_header_scan_rows, ws.max_row)
        for row_idx in range(1, max_row + 1):
            values = [normalize_header(cell.value) for cell in ws[row_idx]]
            if has_field(values, "姓名") and has_placement_field(values, options.placement_mode):
                headers = values
                while headers and headers[-1] == "":
                    headers.pop()
                normalized = [normalize_header(v) for v in headers]
                return WorkbookInfo(source_name, path, ws.title, row_idx, headers, normalized), issues

    return None, [Issue("error", "模板异常", "未找到同时包含「姓名」和指定九宫格落位字段的表头行", source_file=source_name)]
