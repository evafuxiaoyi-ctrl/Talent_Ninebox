from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from .models import Issue, ProcessOptions


def is_valid_excel_member(name: str) -> bool:
    parts = Path(name).parts
    basename = Path(name).name
    return (
        name.lower().endswith(".xlsx")
        and "__MACOSX" not in parts
        and not basename.startswith("~$")
        and not basename.startswith(".")
    )


def extract_excels(zip_path: Path, work_dir: Path, options: ProcessOptions) -> tuple[list[tuple[str, Path]], list[Issue]]:
    issues: list[Issue] = []
    if zip_path.suffix.lower() != ".zip":
        raise ValueError("仅支持 .zip 文件")

    extract_dir = work_dir / "extracted"
    extract_dir.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(zip_path) as zf:
            members = [m for m in zf.infolist() if not m.is_dir() and is_valid_excel_member(m.filename)]
            if not members:
                raise ValueError("未找到有效的 Excel 文件，请确认 zip 中包含 .xlsx 文件。")
            if len(members) > options.max_excel_files:
                raise ValueError("当前压缩包内包含超过 30 个 Excel 文件，请拆分后重新上传。")

            extracted: list[tuple[str, Path]] = []
            for idx, member in enumerate(members, start=1):
                safe_name = f"{idx:03d}_{Path(member.filename).name}"
                target = extract_dir / safe_name
                resolved = target.resolve()
                if extract_dir.resolve() not in resolved.parents:
                    issues.append(Issue("error", "文件格式异常", "zip 路径不安全，已跳过", member.filename))
                    continue
                with zf.open(member) as src, target.open("wb") as dst:
                    shutil.copyfileobj(src, dst)
                extracted.append((member.filename, target))
            return extracted, issues
    except zipfile.BadZipFile as exc:
        raise ValueError("压缩包损坏或不是有效 zip 文件。") from exc
