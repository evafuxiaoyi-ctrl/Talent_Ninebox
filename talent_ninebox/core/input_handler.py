from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from .models import Issue, ProcessOptions

BAD_MOJIBAKE_CHARS = set("ÃÂâ€µ╬╩╠╚╔╝╦═║┌┐└┘├┤┬┴┼─│▒░▓σπΓΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙√ⁿ²")


def _looks_cjk(char: str) -> bool:
    return "\u3400" <= char <= "\u9fff"


def _filename_score(name: str) -> int:
    score = 0
    for char in name:
        if _looks_cjk(char):
            score += 8
        elif char.isascii() and (char.isalnum() or char in "._-/ "):
            score += 1
        elif char in BAD_MOJIBAKE_CHARS:
            score -= 6
        elif ord(char) < 32 or char == "\ufffd":
            score -= 20
    return score


def repair_zip_member_name(name: str) -> str:
    try:
        raw = name.encode("cp437")
    except UnicodeEncodeError:
        return name

    candidates = [name]
    for encoding in ("utf-8", "gb18030", "gbk", "big5"):
        try:
            decoded = raw.decode(encoding)
        except UnicodeDecodeError:
            continue
        if decoded not in candidates:
            candidates.append(decoded)
    return max(candidates, key=_filename_score)


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
                display_name = repair_zip_member_name(member.filename)
                safe_name = f"{idx:03d}_{Path(display_name).name}"
                target = extract_dir / safe_name
                resolved = target.resolve()
                if extract_dir.resolve() not in resolved.parents:
                    issues.append(Issue("error", "文件格式异常", "zip 路径不安全，已跳过", display_name))
                    continue
                with zf.open(member) as src, target.open("wb") as dst:
                    shutil.copyfileobj(src, dst)
                extracted.append((display_name, target))
            return extracted, issues
    except zipfile.BadZipFile as exc:
        raise ValueError("压缩包损坏或不是有效 zip 文件。") from exc
