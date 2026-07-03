from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
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
OUTPUT_DIR = PROJECT_ROOT / "output"

HeaderSelector = Callable[
    [list[HeaderCandidate]],
    tuple[str, ...] | None,
]


class UserCancelledError(Exception):
    """用户取消了需要确认的操作。"""


@dataclass(frozen=True)
class GuiMergeResult:
    """GUI 需要展示的处理结果。"""

    success_file_count: int
    skipped_file_count: int
    failed_file_count: int
    input_row_count: int
    deduplicated_row_count: int
    removed_row_count: int
    output_path: Path


def parse_deduplication_fields(
    raw_fields: str,
    available_fields: list[str],
) -> list[str]:
    """解析并检查 GUI 中输入的去重字段。"""
    if not raw_fields.strip():
        return []

    fields = [
        field.strip()
        for field in raw_fields.replace("，", ",").split(",")
        if field.strip()
    ]
    fields = list(dict.fromkeys(fields))

    if not fields:
        return []

    missing_fields = [
        field for field in fields if field not in available_fields
    ]
    if missing_fields:
        raise ValueError(
            f"去重字段不存在：{', '.join(missing_fields)}\n"
            f"可用字段：{', '.join(available_fields)}"
        )

    return fields


def process_folder(
    input_dir: Path,
    raw_deduplication_fields: str,
    output_dir: Path = OUTPUT_DIR,
    header_selector: HeaderSelector | None = None,
) -> GuiMergeResult:
    """处理一个输入目录，并返回 GUI 展示所需的统计结果。"""
    excel_files = scan_xlsx_files(Path(input_dir))
    if not excel_files:
        raise ValueError("所选文件夹中没有可读取的 .xlsx 文件。")

    read_results = []
    failed_file_count = 0
    for file_path in excel_files:
        try:
            read_results.append(read_first_sheet(file_path))
        except Exception:
            failed_file_count += 1

    candidates = find_top_header_candidates(read_results)
    if not candidates:
        standard_headers = None
    elif len(candidates) == 1:
        standard_headers = candidates[0].headers
    else:
        if header_selector is None:
            raise ValueError("存在多个高频表头，需要手动选择标准表头。")
        standard_headers = header_selector(candidates)
        if standard_headers is None:
            raise UserCancelledError("已取消标准表头选择。")

    validation = validate_headers(read_results, standard_headers)
    if not validation.mergeable_files:
        raise ValueError(
            "没有可参与合并的 Excel 文件。请检查空表、文件内容和表头。"
        )

    merge_result = merge_excel_data(validation.mergeable_files)
    deduplication_fields = parse_deduplication_fields(
        raw_deduplication_fields,
        merge_result.headers,
    )
    deduplication_result = deduplicate_rows(
        merge_result.headers,
        merge_result.rows,
        deduplication_fields,
    )

    output_path = build_output_path(output_dir, deduplication_fields)
    output_path = export_to_xlsx(
        merge_result.headers,
        deduplication_result.rows,
        output_path,
    )

    return GuiMergeResult(
        success_file_count=len(validation.mergeable_files),
        skipped_file_count=len(validation.skipped_files),
        failed_file_count=failed_file_count,
        input_row_count=merge_result.input_row_count,
        deduplicated_row_count=deduplication_result.row_count,
        removed_row_count=deduplication_result.removed_row_count,
        output_path=output_path,
    )


class ExcelMergerApp:
    """Excel 批量合并工具的简单 tkinter 界面。"""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel 批量合并工具")
        self.root.geometry("760x500")
        self.root.minsize(680, 440)

        self.folder_path = tk.StringVar()
        self.deduplication_fields = tk.StringVar()
        self.result_values = {
            "success": tk.StringVar(value="—"),
            "skipped": tk.StringVar(value="—"),
            "failed": tk.StringVar(value="—"),
            "input_rows": tk.StringVar(value="—"),
            "deduplicated_rows": tk.StringVar(value="—"),
            "removed_rows": tk.StringVar(value="—"),
            "output_path": tk.StringVar(value="—"),
        }

        self._build_layout()

    def _build_layout(self) -> None:
        container = ttk.Frame(self.root, padding=20)
        container.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)

        ttk.Label(
            container,
            text="Excel 批量合并、去重与来源标记工具",
            font=("Microsoft YaHei UI", 15, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 18))

        ttk.Label(container, text="输入文件夹：").grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 10),
        )
        folder_entry = ttk.Entry(
            container,
            textvariable=self.folder_path,
            state="readonly",
        )
        folder_entry.grid(row=1, column=1, sticky="ew")
        ttk.Button(
            container,
            text="选择文件夹",
            command=self.select_folder,
        ).grid(row=1, column=2, padx=(10, 0))

        ttk.Label(container, text="去重字段：").grid(
            row=2,
            column=0,
            sticky="w",
            padx=(0, 10),
            pady=(16, 0),
        )
        ttk.Entry(
            container,
            textvariable=self.deduplication_fields,
        ).grid(row=2, column=1, columnspan=2, sticky="ew", pady=(16, 0))
        ttk.Label(
            container,
            text="示例：订单号 / 手机号 / 订单号,手机号；留空表示不去重",
            foreground="#666666",
        ).grid(row=3, column=1, columnspan=2, sticky="w", pady=(4, 16))

        self.start_button = ttk.Button(
            container,
            text="开始合并",
            command=self.start_merge,
        )
        self.start_button.grid(row=4, column=0, columnspan=3, pady=(0, 18))

        result_frame = ttk.LabelFrame(container, text="处理结果", padding=14)
        result_frame.grid(row=5, column=0, columnspan=3, sticky="nsew")
        result_frame.columnconfigure(1, weight=1)
        container.rowconfigure(5, weight=1)

        result_rows = [
            ("成功读取文件数量", "success"),
            ("跳过文件数量", "skipped"),
            ("失败文件数量", "failed"),
            ("合并前总行数", "input_rows"),
            ("去重后总行数", "deduplicated_rows"),
            ("删除重复行数", "removed_rows"),
            ("输出文件路径", "output_path"),
        ]
        for row_index, (label_text, result_key) in enumerate(result_rows):
            ttk.Label(result_frame, text=f"{label_text}：").grid(
                row=row_index,
                column=0,
                sticky="nw",
                padx=(0, 12),
                pady=3,
            )
            ttk.Label(
                result_frame,
                textvariable=self.result_values[result_key],
                wraplength=520 if result_key == "output_path" else 0,
            ).grid(row=row_index, column=1, sticky="w", pady=3)

    def select_folder(self) -> None:
        selected_folder = filedialog.askdirectory(
            title="选择包含 Excel 文件的文件夹",
            mustexist=True,
            parent=self.root,
        )
        if selected_folder:
            self.folder_path.set(selected_folder)

    def select_standard_header(
        self,
        candidates: list[HeaderCandidate],
    ) -> tuple[str, ...] | None:
        lines = ["检测到多个出现次数相同的表头，请选择标准表头：", ""]
        for index, candidate in enumerate(candidates, start=1):
            lines.append(f"{index}. {list(candidate.headers)}")
            lines.append(f"   文件：{', '.join(candidate.file_names)}")

        selected_index = simpledialog.askinteger(
            "选择标准表头",
            "\n".join(lines),
            parent=self.root,
            minvalue=1,
            maxvalue=len(candidates),
        )
        if selected_index is None:
            return None
        return candidates[selected_index - 1].headers

    def start_merge(self) -> None:
        selected_folder = self.folder_path.get().strip()
        if not selected_folder:
            messagebox.showerror(
                "未选择文件夹",
                "请先选择包含 Excel 文件的文件夹。",
                parent=self.root,
            )
            return

        self.start_button.configure(state="disabled")
        self.root.update_idletasks()
        try:
            result = process_folder(
                Path(selected_folder),
                self.deduplication_fields.get(),
                OUTPUT_DIR,
                self.select_standard_header,
            )
        except UserCancelledError:
            return
        except Exception as exc:
            messagebox.showerror(
                "处理失败",
                str(exc),
                parent=self.root,
            )
            return
        finally:
            self.start_button.configure(state="normal")

        self.result_values["success"].set(str(result.success_file_count))
        self.result_values["skipped"].set(str(result.skipped_file_count))
        self.result_values["failed"].set(str(result.failed_file_count))
        self.result_values["input_rows"].set(str(result.input_row_count))
        self.result_values["deduplicated_rows"].set(
            str(result.deduplicated_row_count)
        )
        self.result_values["removed_rows"].set(str(result.removed_row_count))
        self.result_values["output_path"].set(str(result.output_path))

        messagebox.showinfo(
            "处理完成",
            f"Excel 合并完成。\n输出文件：{result.output_path}",
            parent=self.root,
        )


def main() -> None:
    root = tk.Tk()
    ExcelMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
