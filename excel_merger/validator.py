from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

from excel_merger.reader import ExcelReadResult


HeaderTuple = tuple[str, ...]


@dataclass(frozen=True)
class HeaderCandidate:
    """一种表头及其出现次数和对应文件。"""

    headers: HeaderTuple
    count: int
    file_names: tuple[str, ...]


@dataclass(frozen=True)
class SkippedFile:
    """未参与后续处理的文件及其原因。"""

    file_path: Path
    reason: str


@dataclass
class HeaderValidationResult:
    """一批 Excel 文件的表头校验结果。"""

    standard_headers: list[str] | None = None
    standard_header_count: int = 0
    mergeable_files: list[ExcelReadResult] = field(default_factory=list)
    skipped_files: list[SkippedFile] = field(default_factory=list)


def find_top_header_candidates(
    read_results: list[ExcelReadResult],
) -> list[HeaderCandidate]:
    """返回非空文件中出现次数最多的一个或多个表头。"""
    header_counts: Counter[HeaderTuple] = Counter()
    header_files: dict[HeaderTuple, list[str]] = defaultdict(list)

    for read_result in read_results:
        if read_result.is_empty or not read_result.headers:
            continue

        headers = tuple(read_result.headers)
        header_counts[headers] += 1
        header_files[headers].append(read_result.file_path.name)

    if not header_counts:
        return []

    highest_count = max(header_counts.values())
    return [
        HeaderCandidate(
            headers=headers,
            count=count,
            file_names=tuple(header_files[headers]),
        )
        for headers, count in header_counts.items()
        if count == highest_count
    ]


def validate_headers(
    read_results: list[ExcelReadResult],
    standard_headers: HeaderTuple | None,
    skipped_files: list[SkippedFile] | None = None,
) -> HeaderValidationResult:
    """按选定的标准表头划分可合并文件和跳过文件。"""
    validation = HeaderValidationResult(
        standard_headers=list(standard_headers) if standard_headers else None,
        skipped_files=list(skipped_files or []),
    )

    for read_result in read_results:
        if read_result.is_empty:
            if read_result.headers:
                reason = "空表：只有表头，没有数据行"
            else:
                reason = "空表：没有表头和数据行"
            validation.skipped_files.append(
                SkippedFile(read_result.file_path, reason)
            )
            continue

        if not read_result.headers:
            validation.skipped_files.append(
                SkippedFile(read_result.file_path, "没有可用表头")
            )
            continue

        if standard_headers is None:
            validation.skipped_files.append(
                SkippedFile(read_result.file_path, "没有选定标准表头")
            )
            continue

        if tuple(read_result.headers) == standard_headers:
            validation.mergeable_files.append(read_result)
            validation.standard_header_count += 1
        else:
            validation.skipped_files.append(
                SkippedFile(
                    read_result.file_path,
                    f"表头不一致：实际表头为 {read_result.headers}",
                )
            )

    return validation
