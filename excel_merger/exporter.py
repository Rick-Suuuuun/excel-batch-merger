from pathlib import Path
from typing import Any

import pandas as pd


def build_output_path(
    output_dir: Path,
    deduplication_fields: list[str],
) -> Path:
    """根据去重字段生成对应的输出文件名。"""
    output_dir = Path(output_dir)
    if deduplication_fields:
        field_suffix = "_".join(deduplication_fields)
        file_name = f"合并结果_按{field_suffix}去重.xlsx"
    else:
        file_name = "合并结果_未去重.xlsx"
    return output_dir / file_name


def export_to_xlsx(
    headers: list[str],
    rows: list[list[Any]],
    output_path: Path,
) -> Path:
    """创建输出目录，并将内存数据导出为 Excel 文件。"""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    dataframe = pd.DataFrame(rows, columns=headers)
    dataframe.to_excel(output_path, index=False, engine="openpyxl")
    return output_path.resolve()
