from dataclasses import dataclass
from typing import Any

from excel_merger.reader import ExcelReadResult


SOURCE_COLUMN_NAME = "来源文件"


@dataclass(frozen=True)
class MergeResult:
    """保存在内存中的合并数据和统计信息。"""

    headers: list[str]
    rows: list[list[Any]]
    input_row_count: int
    file_row_counts: dict[str, int]

    @property
    def participating_file_count(self) -> int:
        return len(self.file_row_counts)

    @property
    def merged_row_count(self) -> int:
        return len(self.rows)


@dataclass(frozen=True)
class DeduplicationResult:
    """去重后的内存数据和统计信息。"""

    rows: list[list[Any]]
    fields: tuple[str, ...]
    input_row_count: int

    @property
    def row_count(self) -> int:
        return len(self.rows)

    @property
    def removed_row_count(self) -> int:
        return self.input_row_count - self.row_count


def merge_excel_data(read_results: list[ExcelReadResult]) -> MergeResult:
    """纵向合并表头一致的数据，并为每行追加来源文件名。"""
    if not read_results:
        return MergeResult([], [], 0, {})

    standard_headers = read_results[0].headers
    if SOURCE_COLUMN_NAME in standard_headers:
        raise ValueError(f"原始表头中已存在“{SOURCE_COLUMN_NAME}”列")

    merged_rows: list[list[Any]] = []
    file_row_counts: dict[str, int] = {}

    for read_result in read_results:
        if read_result.headers != standard_headers:
            raise ValueError(
                f"{read_result.file_path.name} 的表头与标准表头不一致"
            )

        file_name = read_result.file_path.name
        file_row_counts[file_name] = len(read_result.rows)
        merged_rows.extend(
            [*row, file_name]
            for row in read_result.rows
        )

    input_row_count = sum(file_row_counts.values())
    return MergeResult(
        headers=[*standard_headers, SOURCE_COLUMN_NAME],
        rows=merged_rows,
        input_row_count=input_row_count,
        file_row_counts=file_row_counts,
    )


def deduplicate_rows(
    headers: list[str],
    rows: list[list[Any]],
    fields: list[str],
) -> DeduplicationResult:
    """按一个或多个字段去重，保留第一次出现的数据。"""
    if not fields:
        return DeduplicationResult(
            rows=[row.copy() for row in rows],
            fields=(),
            input_row_count=len(rows),
        )

    missing_fields = [field for field in fields if field not in headers]
    if missing_fields:
        raise ValueError(f"去重字段不存在：{', '.join(missing_fields)}")

    field_indexes = [headers.index(field) for field in fields]
    seen_keys: set[tuple[Any, ...]] = set()
    deduplicated_rows: list[list[Any]] = []

    for row in rows:
        deduplication_key = tuple(row[index] for index in field_indexes)
        if deduplication_key in seen_keys:
            continue
        seen_keys.add(deduplication_key)
        deduplicated_rows.append(row.copy())

    return DeduplicationResult(
        rows=deduplicated_rows,
        fields=tuple(fields),
        input_row_count=len(rows),
    )
