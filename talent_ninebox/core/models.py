from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class ProcessOptions:
    max_excel_files: int = 30
    max_header_scan_rows: int = 20
    add_source_columns: bool = True
    placement_mode: str = "auto"
    stage_label: str = "自动处理"


@dataclass
class Issue:
    level: str
    category: str
    message: str
    source_file: str = ""
    sheet: str = ""
    row: int | None = None
    field: str = ""
    raw_value: Any = None


@dataclass
class WorkbookInfo:
    source_name: str
    path: Path
    sheet_name: str
    header_row: int
    headers: list[str]
    normalized_headers: list[str]

    @property
    def signature(self) -> tuple[str, ...]:
        return tuple(self.normalized_headers)


@dataclass
class RowRecord:
    values: list[Any]
    source_file: str
    sheet: str
    source_row: int
    is_duplicate: bool = False
    notes: list[str] = field(default_factory=list)
    display_name: str = ""
    ninebox_key: str = ""
    high_attrition_risk: bool = False
    value_score_level: str = ""


@dataclass
class ProcessSummary:
    upload_file: str
    stage_label: str = ""
    placement_field: str = ""
    output_file: str = ""
    excel_file_count: int = 0
    processed_file_count: int = 0
    skipped_file_count: int = 0
    total_people: int = 0
    placed_people: int = 0
    unplaced_people: int = 0
    duplicate_count: int = 0
    issue_count: int = 0
    ninebox_counts: dict[str, int] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        return {
            "处理阶段": self.stage_label,
            "九宫格依据字段": self.placement_field,
            "上传文件": self.upload_file,
            "输出文件": self.output_file,
            "识别Excel数量": self.excel_file_count,
            "成功处理文件数": self.processed_file_count,
            "跳过文件数": self.skipped_file_count,
            "总人员数": self.total_people,
            "已落位人数": self.placed_people,
            "未落位人数": self.unplaced_people,
            "重复人员数": self.duplicate_count,
            "异常数量": self.issue_count,
            "九宫格人数": self.ninebox_counts,
        }


@dataclass
class ProcessResult:
    output_file: Path
    summary: ProcessSummary
    issues: list[Issue]
