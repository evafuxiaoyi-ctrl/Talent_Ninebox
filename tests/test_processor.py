from __future__ import annotations

import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from talent_ninebox.core.processor import process_zip
from talent_ninebox.core.ninebox_mapper import NINEBOX_KEYS


def _make_workbook(path: Path, rows: list[tuple[str, str, str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "盘点"
    ws.append(["说明", "", "", ""])
    ws.append(["工号", "姓名", "所属部门", "人才九宫格落位", "评分"])
    for idx, row in enumerate(rows, start=3):
        emp_id, name, dept, place = row
        ws.append([emp_id, name, dept, place, f"=LEN(B{idx})"])
    ws.column_dimensions["B"].width = 18
    wb.save(path)


def test_process_zip_generates_output(tmp_path: Path) -> None:
    first = tmp_path / "first.xlsx"
    second = tmp_path / "nested" / "second.xlsx"
    second.parent.mkdir()
    _make_workbook(first, [("001", "张三", "产品部", "高潜高绩"), ("002", "李四", "产品部", "B2")])
    _make_workbook(second, [("003", "王五", "研发部", "低潜低绩"), ("004", "赵六", "研发部", "未知落位")])

    zip_path = tmp_path / "input.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.write(first, "first.xlsx")
        zf.write(second, "nested/second.xlsx")
        zf.writestr("__MACOSX/ignored.xlsx", "ignored")
        zf.writestr("~$temp.xlsx", "ignored")

    result = process_zip(zip_path, tmp_path / "out")

    assert result.output_file.exists()
    assert result.summary.excel_file_count == 2
    assert result.summary.processed_file_count == 2
    assert result.summary.total_people == 4
    assert result.summary.placed_people == 3
    assert result.summary.unplaced_people == 1

    wb = load_workbook(result.output_file, data_only=False)
    assert set(["使用说明", "处理摘要", "整合总表", "人才九宫格", "异常报告", "文件来源映射"]).issubset(wb.sheetnames)
    merged = wb["整合总表"]
    assert merged.max_row == 6
    assert merged["A1"].value == "说明"
    assert merged["E3"].value == "=LEN(B3)"
    assert merged["E4"].value == "=LEN(B4)"
    assert merged["B3"].value == "张三"
    assert merged["B6"].value == "赵六"

    assert list(result.summary.ninebox_counts) == NINEBOX_KEYS
    assert NINEBOX_KEYS == [
        "高潜力 / 低绩效",
        "高潜力 / 中绩效",
        "高潜力 / 高绩效",
        "中潜力 / 低绩效",
        "中潜力 / 中绩效",
        "中潜力 / 高绩效",
        "低潜力 / 低绩效",
        "低潜力 / 中绩效",
        "低潜力 / 高绩效",
    ]

    summary = wb["处理摘要"]
    rows = list(summary.iter_rows(values_only=True))
    ninebox_row_index = next(idx for idx, row in enumerate(rows) if row[0] == "九宫格人数")
    assert rows[ninebox_row_index + 1] == ("  高潜力 / 低绩效", 0, 0)
    assert rows[ninebox_row_index + 3] == ("  高潜力 / 高绩效", 1, 0.25)


def test_process_zip_skips_sample_row_style_for_merged_data(tmp_path: Path) -> None:
    workbook_path = tmp_path / "with_sample.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "盘点"
    ws.append(["说明", "", "", ""])
    ws.append(["工号", "姓名", "所属部门", "人才九宫格落位", "评分"])
    ws.append(["示例行", "样例", "样例部门", "高潜高绩", "=LEN(B3)"])
    for cell in ws[3]:
        cell.font = Font(color="C00000")
    ws.append(["001", "张三", "产品部", "高潜高绩", "=LEN(B4)"])
    workbook_path.parent.mkdir(exist_ok=True)
    wb.save(workbook_path)

    zip_path = tmp_path / "input.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.write(workbook_path, "with_sample.xlsx")

    result = process_zip(zip_path, tmp_path / "out")

    assert result.summary.total_people == 1
    output = load_workbook(result.output_file, data_only=False)
    try:
        merged = output["整合总表"]
        assert merged["B3"].value == "张三"
        assert merged["B3"].font.color is None or merged["B3"].font.color.rgb != "00C00000"
    finally:
        output.close()


def test_too_many_excels_stops(tmp_path: Path) -> None:
    zip_path = tmp_path / "input.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        for idx in range(31):
            zf.writestr(f"{idx}.xlsx", "not real")

    try:
        process_zip(zip_path, tmp_path / "out")
    except ValueError as exc:
        assert "超过 30 个 Excel" in str(exc)
    else:
        raise AssertionError("expected ValueError")
