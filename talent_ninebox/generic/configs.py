from __future__ import annotations

from .models import GenericSheetConfig, GenericTemplateConfig


ORG_INVENTORY_TEMPLATE = GenericTemplateConfig(
    template_id="org_inventory",
    name="组织盘点表",
    split_fields={
        "一级部门": ["一级部门", "一级组织"],
        "二级部门": ["二级部门", "二级组织"],
        "三级部门": ["三级部门", "三级组织"],
        "负责人邮箱": ["负责人邮箱"],
    },
    sheets=[
        GenericSheetConfig(name="组织架构盘点表", header_row=1, data_start_row=2),
        GenericSheetConfig(name="组织职责盘点表", header_row=1, data_start_row=2),
    ],
)

