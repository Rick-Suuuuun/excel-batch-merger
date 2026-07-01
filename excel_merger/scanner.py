from pathlib import Path


def scan_xlsx_files(input_dir: Path) -> list[Path]:
    """扫描目录顶层的 .xlsx 文件，并排除 Excel 临时文件。"""
    input_dir = Path(input_dir)

    if not input_dir.exists():
        raise FileNotFoundError(f"输入文件夹不存在：{input_dir}")
    if not input_dir.is_dir():
        raise NotADirectoryError(f"输入路径不是文件夹：{input_dir}")

    excel_files = [
        path
        for path in input_dir.iterdir()
        if path.is_file()
        and path.suffix.lower() == ".xlsx"
        and not path.name.startswith("~$")
    ]

    return sorted(excel_files, key=lambda path: path.name.lower())
