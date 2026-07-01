from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd


@dataclass(frozen=True)
class ExcelReadResult:
    """单个 Excel 文件第一张工作表的读取结果。"""

    file_path: Path
    sheet_name: str
    headers: list[str]
    rows: list[list[Any]]

    @property
    def is_empty(self) -> bool:
        """没有数据行时视为空表；即使文件中只有表头也返回 True。"""
        return len(self.rows) == 0


def read_first_sheet(file_path: Path) -> ExcelReadResult:
    """读取 Excel 的第一张工作表，返回表头和数据行。"""
    file_path = Path(file_path)

    with pd.ExcelFile(file_path, engine="openpyxl") as workbook:
        sheet_name = workbook.sheet_names[0]
        dataframe = pd.read_excel(workbook, sheet_name=0, dtype=object)

    headers = [str(column).strip() for column in dataframe.columns]

    # 将 pandas 的缺失值统一转成 None，便于后续阶段处理。
    normalized = dataframe.astype(object).where(pd.notna(dataframe), None)
    rows = normalized.values.tolist()

    return ExcelReadResult(
        file_path=file_path,
        sheet_name=sheet_name,
        headers=headers,
        rows=rows,
    )
