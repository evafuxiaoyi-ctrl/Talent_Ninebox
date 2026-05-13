from __future__ import annotations

import re
from typing import Any

FIELD_ALIASES = {
    "姓名": ["姓名", "员工姓名"],
    "工号": ["工号", "员工工号"],
    "所属部门": ["所属部门", "一级部门", "部门"],
    "一级部门": ["一级部门"],
    "二级部门": ["二级部门"],
    "人才九宫格落位": ["人才九宫格落位", "校准后九宫格落位", "初始九宫格落位"],
    "离职风险": ["离职风险"],
    "价值观综合得分": ["价值观综合得分"],
    "成就客户": ["成就客户", "Customer Focus"],
    "激情人生": ["激情人生", "Passion"],
    "拥抱变化": ["拥抱变化", "Game Changer"],
    "团队协作": ["团队协作", "Teamwork"],
}

PLACEMENT_ALIASES = {
    "initial": ["初始九宫格落位", "人才九宫格落位"],
    "final": ["校准后九宫格落位", "人才九宫格落位"],
    "auto": ["校准后九宫格落位", "初始九宫格落位", "人才九宫格落位"],
}


def normalize_field(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", "", text)
    return text


def find_field_index(headers: list[str], canonical_name: str) -> int | None:
    aliases = FIELD_ALIASES.get(canonical_name, [canonical_name])
    normalized_headers = [normalize_field(header) for header in headers]
    for alias in aliases:
        normalized_alias = normalize_field(alias)
        for idx, normalized_header in enumerate(normalized_headers):
            if normalized_header == normalized_alias:
                return idx
    for alias in aliases:
        normalized_alias = normalize_field(alias)
        for idx, normalized_header in enumerate(normalized_headers):
            if normalized_alias and normalized_alias in normalized_header:
                return idx
    return None


def has_field(headers: list[str], canonical_name: str) -> bool:
    return find_field_index(headers, canonical_name) is not None


def find_placement_index(headers: list[str], placement_mode: str) -> int | None:
    aliases = PLACEMENT_ALIASES.get(placement_mode, PLACEMENT_ALIASES["auto"])
    normalized_headers = [normalize_field(header) for header in headers]
    for alias in aliases:
        normalized_alias = normalize_field(alias)
        for idx, normalized_header in enumerate(normalized_headers):
            if normalized_header == normalized_alias:
                return idx
    for alias in aliases:
        normalized_alias = normalize_field(alias)
        for idx, normalized_header in enumerate(normalized_headers):
            if normalized_alias and normalized_alias in normalized_header:
                return idx
    return None


def placement_field_label(headers: list[str], placement_mode: str) -> str:
    idx = find_placement_index(headers, placement_mode)
    if idx is None:
        return ""
    return str(headers[idx]).strip()


def has_placement_field(headers: list[str], placement_mode: str) -> bool:
    return find_placement_index(headers, placement_mode) is not None
