from __future__ import annotations

import os
import threading
import traceback
import tkinter as tk
from pathlib import Path
from queue import Empty, Queue
from tkinter import filedialog, messagebox, ttk

from .config_store import (
    ensure_user_config,
    get_user_data_dir,
    import_excel_to_user_config,
    load_config,
    restore_default_config,
    update_ocr_settings,
    update_render_settings,
)
from .pdf_processor import resolve_output_teachers
from .template_generator import create_teacher_list_template
from .version import build_app_title


class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(build_app_title())
        self.root.geometry("920x760")
        self.root.report_callback_exception = self._handle_callback_exception

        self.message_queue: Queue[tuple] = Queue()
        self.worker_thread: threading.Thread | None = None

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.config_status_var = tk.StringVar()
        self.code_check_var = tk.StringVar()
        self.label_x_var = tk.StringVar()
        self.label_y_var = tk.StringVar()
        self.label_font_size_var = tk.StringVar()
        self.label_gap_var = tk.StringVar()
        self.label_anchor_var = tk.StringVar(value="bottom_left")
        self.star_scale_var = tk.StringVar()
        self.star_gap_var = tk.StringVar()
        self.star_y_offset_var = tk.StringVar()
        self.text_only_mode_var = tk.BooleanVar(value=False)

        self._build_ui()
        self.code_check_var.set(
            "请选择输入档案。\n"
            "输入档案去重人数：-\n"
            "老师名单成功对照：-\n"
            "未知代码：-"
        )
        self._refresh_config_status()
        self.root.after(150, self._poll_queue)

    def run(self):
        self.root.mainloop()

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(6, weight=1)

        title = ttk.Label(self.root, text=build_app_title(), font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, sticky="w", padx=14, pady=(14, 4))

        description = ttk.Label(self.root, text="载入 PDF 或 TIFF，系统会先检查本次监考代码，再输出新的 PDF。")
        description.grid(row=1, column=0, sticky="w", padx=14, pady=(0, 10))

        config_frame = ttk.LabelFrame(self.root, text="设定资料")
        config_frame.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 10))
        config_frame.columnconfigure(0, weight=1)

        status_label = ttk.Label(
            config_frame,
            textvariable=self.config_status_var,
            justify="left",
        )
        status_label.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 8))

        button_row = ttk.Frame(config_frame)
        button_row.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 10))
        ttk.Button(button_row, text="下载老师名单模板", command=self._download_template_excel).pack(side="left", padx=(0, 8))
        ttk.Button(button_row, text="匯入老师名单 Excel", command=self._import_excel_config).pack(side="left", padx=(0, 8))
        ttk.Button(button_row, text="开启设定资料夹", command=self._open_config_folder).pack(side="left", padx=(0, 8))
        ttk.Button(button_row, text="还原内建设定", command=self._restore_default_config).pack(side="left")
        ttk.Checkbutton(button_row, text="仅文字层模式", variable=self.text_only_mode_var, command=self._save_processing_mode).pack(side="left", padx=(8, 0))

        layout_frame = ttk.LabelFrame(self.root, text="标签与星号设定")
        layout_frame.grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 10))
        for index in range(10):
            layout_frame.columnconfigure(index, weight=0)
        layout_frame.columnconfigure(9, weight=1)

        ttk.Label(layout_frame, text="标签锚点").grid(row=0, column=0, sticky="w", padx=(10, 4), pady=(10, 8))
        anchor_combo = ttk.Combobox(
            layout_frame,
            textvariable=self.label_anchor_var,
            values=("top_left", "bottom_left"),
            width=12,
            state="readonly",
        )
        anchor_combo.grid(row=0, column=1, sticky="w", padx=(0, 10), pady=(10, 8))
        ttk.Label(layout_frame, text="X").grid(row=0, column=2, sticky="w", padx=(0, 4), pady=(10, 8))
        ttk.Entry(layout_frame, textvariable=self.label_x_var, width=8).grid(row=0, column=3, sticky="w", padx=(0, 10), pady=(10, 8))
        ttk.Label(layout_frame, text="Y").grid(row=0, column=4, sticky="w", padx=(0, 4), pady=(10, 8))
        ttk.Entry(layout_frame, textvariable=self.label_y_var, width=8).grid(row=0, column=5, sticky="w", padx=(0, 10), pady=(10, 8))
        ttk.Label(layout_frame, text="字体大小").grid(row=0, column=6, sticky="w", padx=(0, 4), pady=(10, 8))
        ttk.Entry(layout_frame, textvariable=self.label_font_size_var, width=8).grid(row=0, column=7, sticky="w", padx=(0, 10), pady=(10, 8))
        ttk.Label(layout_frame, text="栏位间距").grid(row=0, column=8, sticky="w", padx=(0, 4), pady=(10, 8))
        ttk.Entry(layout_frame, textvariable=self.label_gap_var, width=8).grid(row=0, column=9, sticky="w", padx=(0, 10), pady=(10, 8))

        ttk.Label(layout_frame, text="星号倍率").grid(row=1, column=0, sticky="w", padx=(10, 4), pady=(0, 10))
        ttk.Entry(layout_frame, textvariable=self.star_scale_var, width=8).grid(row=1, column=1, sticky="w", padx=(0, 10), pady=(0, 10))
        ttk.Label(layout_frame, text="星号距离").grid(row=1, column=2, sticky="w", padx=(0, 4), pady=(0, 10))
        ttk.Entry(layout_frame, textvariable=self.star_gap_var, width=8).grid(row=1, column=3, sticky="w", padx=(0, 10), pady=(0, 10))
        ttk.Label(layout_frame, text="星号Y偏移").grid(row=1, column=4, sticky="w", padx=(0, 4), pady=(0, 10))
        ttk.Entry(layout_frame, textvariable=self.star_y_offset_var, width=8).grid(row=1, column=5, sticky="w", padx=(0, 10), pady=(0, 10))
        ttk.Button(layout_frame, text="保存标签/星号设定", command=self._save_render_settings).grid(row=1, column=9, sticky="e", padx=(0, 10), pady=(0, 10))

        file_frame = ttk.LabelFrame(self.root, text="PDF 处理")
        file_frame.grid(row=4, column=0, sticky="ew", padx=14, pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="输入档案").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 6))
        ttk.Entry(file_frame, textvariable=self.input_var).grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=(10, 6))
        ttk.Button(file_frame, text="选择档案", command=self._choose_input_pdf).grid(row=0, column=2, sticky="ew", padx=(0, 10), pady=(10, 6))

        ttk.Label(file_frame, text="输出 PDF").grid(row=1, column=0, sticky="w", padx=10, pady=(0, 10))
        ttk.Entry(file_frame, textvariable=self.output_var).grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=(0, 10))
        ttk.Button(file_frame, text="另存为", command=self._choose_output_pdf).grid(row=1, column=2, sticky="ew", padx=(0, 10), pady=(0, 10))
        self.run_button = ttk.Button(file_frame, text="开始处理", command=self._start_processing)
        self.run_button.grid(row=2, column=2, sticky="e", padx=(0, 10), pady=(0, 10))

        check_frame = ttk.LabelFrame(self.root, text="监考人员代码检查")
        check_frame.grid(row=5, column=0, sticky="ew", padx=14, pady=(0, 10))
        check_frame.columnconfigure(0, weight=1)
        ttk.Label(check_frame, textvariable=self.code_check_var, justify="left").grid(row=0, column=0, sticky="w", padx=10, pady=10)

        self.progress_bar = ttk.Progressbar(self.root, mode="determinate", maximum=100)
        self.progress_bar.grid(row=6, column=0, sticky="ew", padx=14, pady=(0, 8))

        log_frame = ttk.LabelFrame(self.root, text="处理日志")
        log_frame.grid(row=7, column=0, sticky="nsew", padx=14, pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, wrap="word", height=18)
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns", pady=10)
        self.log_text.configure(yscrollcommand=scrollbar.set, state="disabled")

        self.root.rowconfigure(7, weight=1)

    def _choose_input_pdf(self):
        file_path = filedialog.askopenfilename(
            parent=self.root,
            title="选择监考档案",
            filetypes=[("Supported Files", "*.pdf *.tif *.tiff"), ("PDF Files", "*.pdf"), ("TIFF Files", "*.tif *.tiff")],
        )
        if not file_path:
            return
        self.input_var.set(file_path)
        if not self.output_var.get().strip():
            input_path = Path(file_path)
            self.output_var.set(str(input_path.with_name(f"{input_path.stem}_marked.pdf")))
        self._analyze_selected_pdf(Path(file_path))

    def _choose_output_pdf(self):
        initial_path = self.output_var.get().strip() or self.input_var.get().strip()
        file_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="选择输出 PDF",
            initialfile=Path(initial_path).name if initial_path else "output_marked.pdf",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
        )
        if file_path:
            self.output_var.set(file_path)

    def _download_template_excel(self):
        """下载老师名单 Excel 模板。"""
        file_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="保存老师名单模板",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="老师名单模板.xlsx",
        )
        if not file_path:
            return
        try:
            create_teacher_list_template(Path(file_path))
            self._append_log(f"已下载模板：{file_path}")
            messagebox.showinfo("下载成功", f"模板已保存到：\n{file_path}", parent=self.root)
        except Exception as exc:
            messagebox.showerror("下载失败", str(exc), parent=self.root)

    def _import_excel_config(self):
        file_path = filedialog.askopenfilename(
            parent=self.root,
            title="选择老师名单 Excel",
            filetypes=[("Excel Files", "*.xlsx *.xlsm")],
        )
        if not file_path:
            return
        try:
            target_path = import_excel_to_user_config(Path(file_path))
            self._refresh_config_status()
            self._append_log(f"已匯入设定：{target_path}")
        except Exception as exc:
            messagebox.showerror("匯入失败", str(exc), parent=self.root)

    def _open_config_folder(self):
        config_dir = get_user_data_dir()
        config_dir.mkdir(parents=True, exist_ok=True)
        os.startfile(str(config_dir))

    def _restore_default_config(self):
        try:
            target_path = restore_default_config()
            self._refresh_config_status()
            self._append_log(f"已还原内建设定：{target_path}")
        except Exception as exc:
            messagebox.showerror("还原失败", str(exc), parent=self.root)

    def _save_render_settings(self):
        try:
            render_settings = self._collect_render_settings()
            target_path = update_render_settings(render_settings)
            self._refresh_config_status()
            self._append_log(f"已保存标签设定：{target_path}")
        except ValueError:
            messagebox.showerror("储存失败", "标签设定必须是数字。", parent=self.root)
        except Exception as exc:
            messagebox.showerror("储存失败", str(exc), parent=self.root)

    def _save_processing_mode(self):
        try:
            target_path = update_ocr_settings({"text_only_mode": bool(self.text_only_mode_var.get())})
            self._refresh_config_status()
            self._append_log(f"已保存处理模式设定：{target_path}")
        except Exception as exc:
            messagebox.showerror("储存失败", str(exc), parent=self.root)

    def _start_processing(self):
        input_path = Path(self.input_var.get().strip())
        output_path = Path(self.output_var.get().strip())

        if not input_path.exists():
            messagebox.showwarning("资料不足", "请先选择有效的输入档案。")
            return
        if output_path.suffix.lower() != ".pdf":
            messagebox.showwarning("资料不足", "输出档案必须是 .pdf。")
            return

        try:
            ensure_user_config()
            target_path = update_render_settings(self._collect_render_settings())
            update_ocr_settings({"text_only_mode": bool(self.text_only_mode_var.get())})
            load_config()
            self._append_log(f"本次处理已套用目前标签设定：{target_path}")
        except Exception as exc:
            messagebox.showerror("设定错误", str(exc))
            return

        self.progress_bar["value"] = 5
        self._set_log_text("")
        self.run_button.state(["disabled"])

        def worker():
            try:
                from .pdf_processor import process_pdf

                config = load_config()
                self.message_queue.put(("progress", 10))
                result = process_pdf(
                    input_path,
                    output_path,
                    config,
                    progress_callback=lambda message: self.message_queue.put(("log", message)),
                )
                summary_lines = [
                    f"完成：{result.output_path}",
                    f"共输出 {len(result.matched_teachers)} 位老师。",
                ]
                if result.unresolved_teacher_codes:
                    summary_lines.append("以下代码在老师名单中找不到，已照常输出并以 !代码 标示：")
                    summary_lines.extend(f"- {code}" for code in result.unresolved_teacher_codes)
                if result.warnings:
                    summary_lines.append("警告：")
                    summary_lines.extend(f"- {item}" for item in result.warnings)
                self.message_queue.put(("finished", True, "\n".join(summary_lines)))
            except Exception as exc:
                self.message_queue.put(("finished", False, str(exc)))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    def _refresh_config_status(self):
        try:
            config_path = ensure_user_config()
            config = load_config(config_path)
            self.config_status_var.set(
                "目前已载入老师名单与座位区块设定。\n"
                f"标签锚点：{'左下角' if config.label_anchor == 'bottom_left' else '左上角'}；"
                f"仅文字层模式：{'开' if config.text_only_mode else '关'}"
            )
            self.label_anchor_var.set(config.label_anchor)
            self.label_x_var.set(f"{config.label_x:.1f}")
            self.label_y_var.set(f"{config.label_y:.1f}")
            self.label_font_size_var.set(f"{config.label_font_size:.1f}")
            self.label_gap_var.set(f"{config.label_gap:.1f}")
            self.star_scale_var.set(f"{config.star_scale:.2f}")
            self.star_gap_var.set(f"{config.star_gap:.1f}")
            self.star_y_offset_var.set(f"{config.star_y_offset:.1f}")
            self.text_only_mode_var.set(config.text_only_mode)
        except Exception as exc:
            self.config_status_var.set(
                "目前没有可用设定。\n"
                "请先匯入 Excel，或将 default_config.json 放到 data 目录。\n"
                f"错误：{exc}"
            )
            for var in (
                self.label_x_var,
                self.label_y_var,
                self.label_font_size_var,
                self.label_gap_var,
                self.star_scale_var,
                self.star_gap_var,
                self.star_y_offset_var,
            ):
                var.set("")
            self.label_anchor_var.set("bottom_left")
            self.text_only_mode_var.set(False)
            self.code_check_var.set(
                "请选择输入档案。\n"
                "输入档案去重人数：-\n"
                "老师名单成功对照：-\n"
                "未知代码：-"
            )

    def _collect_render_settings(self) -> dict[str, float]:
        return {
            "label_anchor": self.label_anchor_var.get().strip(),
            "label_x": float(self.label_x_var.get().strip()),
            "label_y": float(self.label_y_var.get().strip()),
            "label_font_size": float(self.label_font_size_var.get().strip()),
            "label_gap": float(self.label_gap_var.get().strip()),
            "star_scale": float(self.star_scale_var.get().strip()),
            "star_gap": float(self.star_gap_var.get().strip()),
            "star_y_offset": float(self.star_y_offset_var.get().strip()),
        }

    def _append_log(self, message: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _handle_callback_exception(self, exc_type, exc_value, exc_traceback):
        error_text = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback)).strip()
        self._append_log(error_text)
        messagebox.showerror("程式错误", error_text, parent=self.root)

    def _set_log_text(self, text: str):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        if text:
            self.log_text.insert("1.0", text)
        self.log_text.configure(state="disabled")

    def _poll_queue(self):
        try:
            while True:
                event = self.message_queue.get_nowait()
                event_type = event[0]
                if event_type == "log":
                    self._append_log(event[1])
                elif event_type == "progress":
                    self.progress_bar["value"] = event[1]
                elif event_type == "finished":
                    success = event[1]
                    message = event[2]
                    self.run_button.state(["!disabled"])
                    self.progress_bar["value"] = 100 if success else 0
                    self._append_log(message)
                    if success:
                        messagebox.showinfo("处理完成", message, parent=self.root)
                    else:
                        messagebox.showerror("处理失败", message)
        except Empty:
            pass
        finally:
            self.root.after(150, self._poll_queue)

    def _analyze_selected_pdf(self, input_path: Path):
        try:
            config = load_config()
            teachers = resolve_output_teachers(input_path, config.teachers, config)
            unresolved_codes = [teacher.initial for teacher in teachers if teacher.unresolved]
            matched_count = len(teachers) - len(unresolved_codes)
            if unresolved_codes:
                self.code_check_var.set(
                    "已完成输入档案代码检查。\n"
                    f"输入档案去重人数：{len(teachers)}\n"
                    f"老师名单成功对照：{matched_count}\n"
                    f"未知代码：{len(unresolved_codes)}（{', '.join(unresolved_codes)}）\n"
                    "未知代码仍会照常输出，并以 !代码 标示。"
                )
            else:
                self.code_check_var.set(
                    "已完成输入档案代码检查。\n"
                    f"输入档案去重人数：{len(teachers)}\n"
                    f"老师名单成功对照：{matched_count}\n"
                    "未知代码：0"
                )
        except Exception as exc:
            self.code_check_var.set(f"代码检查失败：{exc}")
