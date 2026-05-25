from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class GenericSheetConfig:
    name: str
    header_row: int = 1
    data_start_row: int = 2
    split_field_aliases: dict[str, list[str]] = field(default_factory=dict)


@dataclass(frozen=True)
class GenericTemplateConfig:
    template_id: str
    name: str
    sheets: list[GenericSheetConfig]
    split_fields: dict[str, list[str]]


@dataclass(frozen=True)
class GenericSplitResult:
    output_file: Path
    summary: dict[str, Any]


@dataclass(frozen=True)
class GenericMergeResult:
    output_file: Path
    summary: dict[str, Any]
