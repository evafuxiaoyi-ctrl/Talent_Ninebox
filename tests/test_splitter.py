from __future__ import annotations

import zipfile

from openpyxl import Workbook, load_workbook

from talent_ninebox.core.splitter import list_split_fields, list_split_sheets, split_workbook_by_field


def _make_split_workbook(path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "人才盘点"
    sheet.append(["说明"])
    sheet.append(["请填写基础字段"])
    sheet.append(["员工姓名", "工号", "邮箱", "岗位", "一级部门", "二级部门", "盘点人", "备注公式"])
    sheet.append(["示例行", "000", "sample@example.com", "示例", "XXX", "XXX", "李四", '=A4&"-"&E4'])
    sheet.append(["张三", "001", "a@example.com", "顾问", "销售部", "华东", "傅肖忆", '=A5&"-"&E5'])
    sheet.append(["李四", "002", "b@example.com", "主管", "教研部", "数学", "赵桐", '=A6&"-"&E6'])
    sheet.append(["王五", "003", "c@example.com", "顾问", "销售部", "华南", "傅肖忆", '=A7&"-"&E7'])
    params = workbook.create_sheet("参数")
    params["A1"] = "保留"
    workbook.save(path)


def test_list_split_fields_reads_allowed_headers_from_row_3(tmp_path) -> None:
    workbook_path = tmp_path / "template.xlsx"
    _make_split_workbook(workbook_path)

    fields = list_split_fields(workbook_path)

    assert [field.name for field in fields] == ["一级部门", "二级部门", "盘点人"]
    assert fields[0].column_letter == "E"


def test_list_split_sheets_returns_sheet_fields(tmp_path) -> None:
    workbook_path = tmp_path / "template.xlsx"
    _make_split_workbook(workbook_path)

    sheets = list_split_sheets(workbook_path)

    assert len(sheets) == 1
    assert sheets[0].name == "人才盘点"
    assert [field.name for field in sheets[0].fields] == ["一级部门", "二级部门", "盘点人"]


def test_split_workbook_by_field_creates_zip_with_grouped_excels(tmp_path) -> None:
    workbook_path = tmp_path / "template.xlsx"
    output_dir = tmp_path / "out"
    output_dir.mkdir()
    _make_split_workbook(workbook_path)

    result = split_workbook_by_field(workbook_path, output_dir, "一级部门", "人才盘点")

    assert result.output_file.suffix == ".zip"
    assert result.summary["生成文件数"] == 2
    assert result.summary["总人员数"] == 3

    with zipfile.ZipFile(result.output_file) as archive:
        names = archive.namelist()
        assert len(names) == 2
        assert any("销售部" in name for name in names)
        archive.extractall(tmp_path / "unzipped")

    sales_file = next((tmp_path / "unzipped").glob("*销售部*.xlsx"))
    workbook = load_workbook(sales_file, data_only=False)
    try:
        sheet = workbook["人才盘点"]
        assert sheet.max_row == 5
        assert sheet["A4"].value == "张三"
        assert sheet["A5"].value == "王五"
        assert sheet["H5"].value == '=A5&"-"&E5'
        assert "参数" in workbook.sheetnames
    finally:
        workbook.close()
