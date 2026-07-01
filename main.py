from pathlib import Path
from typing import Callable

from excel_merger.exporter import build_output_path, export_to_xlsx
from excel_merger.merger import deduplicate_rows, merge_excel_data
from excel_merger.reader import read_first_sheet
from excel_merger.scanner import scan_xlsx_files
from excel_merger.validator import (
    HeaderCandidate,
    find_top_header_candidates,
    validate_headers,
)


PROJECT_ROOT = Path(__file__).resolve().parent
INPUT_DIR = PROJECT_ROOT / "sample_data" / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"
PREVIEW_ROW_COUNT = 5


def choose_standard_header(
    candidates: list[HeaderCandidate],
    input_func: Callable[[str], str] = input,
) -> tuple[str, ...] | None:
    """唯一最高频时自动选择，并列时提示用户在终端选择。"""
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0].headers

    print("检测到多个表头出现次数相同，请手动选择标准表头：")
    for index, candidate in enumerate(candidates, start=1):
        print(f"  {index}. {list(candidate.headers)}")
        print(f"     出现次数：{candidate.count}")
        print(f"     文件：{', '.join(candidate.file_names)}")

    while True:
        choice = input_func(f"请输入序号（1-{len(candidates)}）：").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(candidates):
            return candidates[int(choice) - 1].headers
        print("输入无效，请输入列表中的数字序号。")


def prompt_deduplication_fields(
    available_fields: list[str],
    input_func: Callable[[str], str] = input,
) -> list[str]:
    """询问去重字段；直接回车表示不去重。"""
    while True:
        raw_fields = input_func(
            "是否需要去重？请输入去重字段"
            "（多个字段用逗号分隔，直接回车不去重）："
        ).strip()

        if not raw_fields:
            return []

        fields = [
            field.strip()
            for field in raw_fields.replace("，", ",").split(",")
            if field.strip()
        ]
        fields = list(dict.fromkeys(fields))

        if not fields:
            print("没有识别到有效字段，请重新输入。")
            continue

        missing_fields = [
            field for field in fields if field not in available_fields
        ]
        if missing_fields:
            print(f"字段不存在：{', '.join(missing_fields)}")
            print(f"可用字段：{', '.join(available_fields)}")
            continue

        return fields


def main(input_func: Callable[[str], str] = input) -> None:
    """筛选、合并、按需去重并导出 Excel 数据。"""
    print(f"输入文件夹：{INPUT_DIR}")

    try:
        excel_files = scan_xlsx_files(INPUT_DIR)
    except (FileNotFoundError, NotADirectoryError) as exc:
        print(f"扫描失败：{exc}")
        return

    if not excel_files:
        print("未发现可读取的 .xlsx 文件。")
        return

    print(f"发现 {len(excel_files)} 个 .xlsx 文件。\n")

    read_results = []
    failed_files: list[tuple[Path, str]] = []
    for file_path in excel_files:
        try:
            read_results.append(read_first_sheet(file_path))
        except Exception as exc:
            failed_files.append((file_path, str(exc)))

    candidates = find_top_header_candidates(read_results)
    standard_headers = choose_standard_header(candidates, input_func)
    validation = validate_headers(
        read_results,
        standard_headers,
    )

    print("标准表头：")
    if validation.standard_headers is None:
        print("  未找到可用表头")
    else:
        print(f"  {validation.standard_headers}")
    print(f"标准表头出现次数：{validation.standard_header_count}")

    print("\n可合并文件列表：")
    if validation.mergeable_files:
        for read_result in validation.mergeable_files:
            print(f"  - {read_result.file_path.name}")
    else:
        print("  无")

    print("\n被跳过文件列表：")
    if validation.skipped_files:
        for skipped_file in validation.skipped_files:
            print(f"  - {skipped_file.file_path.name}")
            print(f"    原因：{skipped_file.reason}")
    else:
        print("  无")

    print("\n读取失败文件列表：")
    if failed_files:
        for file_path, reason in failed_files:
            print(f"  - {file_path.name}")
            print(f"    原因：{reason}")
    else:
        print("  无")

    if not validation.mergeable_files:
        print("\n没有可参与合并的数据。")
        return

    merge_result = merge_excel_data(validation.mergeable_files)
    deduplication_fields = prompt_deduplication_fields(
        merge_result.headers,
        input_func,
    )
    deduplication_result = deduplicate_rows(
        merge_result.headers,
        merge_result.rows,
        deduplication_fields,
    )
    output_path = build_output_path(
        OUTPUT_DIR,
        deduplication_fields,
    )
    output_path = export_to_xlsx(
        merge_result.headers,
        deduplication_result.rows,
        output_path,
    )

    print("\n处理统计：")
    print(f"  成功读取文件数量：{len(validation.mergeable_files)}")
    print(f"  跳过文件数量：{len(validation.skipped_files)}")
    print(f"  失败文件数量：{len(failed_files)}")
    print(f"  参与合并的文件数量：{merge_result.participating_file_count}")
    print("  每个文件贡献的数据行数：")
    for file_name, row_count in merge_result.file_row_counts.items():
        print(f"    - {file_name}：{row_count} 行")
    print(f"  合并前总行数：{merge_result.input_row_count}")
    if deduplication_result.fields:
        print(f"  去重字段：{', '.join(deduplication_result.fields)}")
    else:
        print("  去重字段：未去重")
    print(f"  去重后总行数：{deduplication_result.row_count}")
    print(f"  删除重复行数：{deduplication_result.removed_row_count}")
    print(f"  输出文件路径：{output_path}")

    print(f"\n处理结果前 {PREVIEW_ROW_COUNT} 行预览：")
    print(f"  {merge_result.headers}")
    for row in deduplication_result.rows[:PREVIEW_ROW_COUNT]:
        print(f"  {row}")


if __name__ == "__main__":
    main()
