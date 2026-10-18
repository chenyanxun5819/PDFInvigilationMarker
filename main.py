from __future__ import annotations

import argparse
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox

from app.config_store import ensure_user_config, import_excel_to_user_config, load_config
from app.version import APP_NAME, APP_VERSION

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=f"{APP_NAME} v{APP_VERSION}")
    parser.add_argument("--input", type=Path, help="输入 PDF / TIFF 路径")
    parser.add_argument("--output", type=Path, help="输出 PDF 路径")
    parser.add_argument("--config", type=Path, help="指定 JSON 设定档路径")
    parser.add_argument("--import-excel", type=Path, help="把老师名单 Excel 匯入成目前使用的设定档")
    parser.add_argument("--version", action="version", version=f"{APP_NAME} v{APP_VERSION}")
    return parser


def run_cli(args) -> int:
    if args.import_excel:
        target_path = import_excel_to_user_config(args.import_excel)
        print(f"已匯入设定：{target_path}")
        if not args.input and not args.output:
            return 0

    if args.input or args.output:
        if not args.input or not args.output:
            raise ValueError("命令列模式下必须同时提供 --input 与 --output。")
        from app.pdf_processor import process_pdf

        config = load_config(args.config if args.config else ensure_user_config())
        result = process_pdf(args.input, args.output, config, progress_callback=print)
        print(f"完成：{result.output_path}")
        print(f"共输出 {len(result.matched_teachers)} 位老师。")
        for warning in result.warnings:
            print(f"WARNING: {warning}")
        return 0

    from app.main_window import MainWindow

    window = MainWindow()
    window.run()
    return 0


def main() -> int:
    parser = build_argument_parser()
    args = parser.parse_args()
    try:
        return run_cli(args)
    except Exception as exc:
        if not sys.stderr:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("程式错误", str(exc), parent=root)
            root.destroy()
        print(f"错误：{exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
