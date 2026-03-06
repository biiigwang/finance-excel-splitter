#!/usr/bin/env python3
"""
Finance Excel Splitter - GUI Application

A graphical user interface for splitting Excel files by department.
Uses tkinter for cross-platform compatibility.
"""

import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from typing import Tuple
from tkinter import ttk, filedialog, messagebox, font

from openpyxl import load_workbook

from core.sheet_analyzer import SheetAnalyzer
from core.department_collector import DepartmentCollector
from core.department_index import DepartmentIndex
from core.workbook_builder import WorkbookBuilder


class FinanceSplitterGUI:
    """GUI application for splitting Excel files by department."""

    def __init__(self, root: tk.Tk):
        """
        Initialize the GUI application.

        Args:
            root: The root tkinter window
        """
        self.root = root
        self.root.title("甜甜的财务数据拆分工具")
        self.root.geometry("700x450")
        self.root.minsize(600, 400)

        # Set default font
        self._setup_default_font()

        # Set window icon (if available)
        self._setup_icon()

        try:
            if sys.platform == 'darwin':  # macOS
                self.root.tk.call('tk', 'scaling', 1.0)
        except Exception:
            pass

        # Variables
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar(value=str(Path.cwd() / 'output'))
        self.status_text = tk.StringVar(value="Ready")
        self.progress_value = tk.DoubleVar(value=0)
        self.remove_empty_sheets = tk.BooleanVar(value=True)
        # Display value to internal value mapping for style mode
        # Include both display values and internal values for compatibility
        self._style_display_map = {
            "统一样式（默认）": "unified",
            "unified": "unified",
            "保留原样式": "original",
            "original": "original",
            "无样式": "none",
            "none": "none"
        }
        # Use display value for StringVar, will be converted when passed to WorkbookBuilder
        self.style_mode = tk.StringVar(value="统一样式（默认）")

        self.create_widgets()
        self.layout_widgets()

    def create_widgets(self) -> None:
        """Create all GUI widgets."""
        # Title label
        self.title_label = ttk.Label(
            self.root,
            text="甜甜的财务数据拆分工具",
            font=('Arial', 16, 'bold')
        )

        self.subtitle_label = ttk.Label(
            self.root,
            text="按科室拆分 Excel 财务数据",
            font=('Arial', 10)
        )

        # Input file section
        self.input_frame = ttk.LabelFrame(self.root, text="输入文件", padding=10)
        self.input_entry = ttk.Entry(
            self.input_frame,
            textvariable=self.input_path,
            width=50
        )
        self.input_button = ttk.Button(
            self.input_frame,
            text="浏览...",
            command=self.browse_input
        )

        # Output directory section
        self.output_frame = ttk.LabelFrame(self.root, text="输出目录", padding=10)
        self.output_entry = ttk.Entry(
            self.output_frame,
            textvariable=self.output_path,
            width=50
        )
        self.output_button = ttk.Button(
            self.output_frame,
            text="浏览...",
            command=self.browse_output
        )

        # Options section
        self.options_frame = ttk.LabelFrame(self.root, text="选项", padding=10)
        self.remove_empty_checkbox = ttk.Checkbutton(
            self.options_frame,
            text="移除空白子表",
            variable=self.remove_empty_sheets
        )

        # Style mode selection
        self.style_label = ttk.Label(self.options_frame, text="样式模式:")
        self.style_combobox = ttk.Combobox(
            self.options_frame,
            textvariable=self.style_mode,
            values=["统一样式（默认）", "保留原样式", "无样式"],
            state="readonly",
            width=15
        )
        self.style_combobox.current(0)
        # Bind selection to convert display value to internal value
        self.style_combobox.bind('<<ComboboxSelected>>', self._on_style_select)

        # Progress section
        self.progress_frame = ttk.LabelFrame(self.root, text="处理进度", padding=10)
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            variable=self.progress_value,
            maximum=100,
            mode='determinate',
            length=400
        )
        self.status_label = ttk.Label(
            self.progress_frame,
            textvariable=self.status_text
        )

        # Button section
        self.button_frame = ttk.Frame(self.root)
        self.start_button = ttk.Button(
            self.button_frame,
            text="开始处理",
            command=self.start_processing,
            width=15
        )
        self.exit_button = ttk.Button(
            self.button_frame,
            text="退出",
            command=self.root.quit,
            width=15
        )

    def layout_widgets(self) -> None:
        """Layout all widgets in the window."""
        # Title
        self.title_label.pack(pady=(20, 5))
        self.subtitle_label.pack(pady=(0, 15))

        # Input frame
        self.input_frame.pack(fill='x', padx=20, pady=5)
        self.input_entry.pack(side='left', fill='x', expand=True, padx=(0, 5))
        self.input_button.pack(side='right')

        # Output frame
        self.output_frame.pack(fill='x', padx=20, pady=5)
        self.output_entry.pack(side='left', fill='x', expand=True, padx=(0, 5))
        self.output_button.pack(side='right')

        # Options frame
        self.options_frame.pack(fill='x', padx=20, pady=5)
        # Make options horizontal layout
        self.remove_empty_checkbox.pack(side='left', padx=(0, 20))
        self.style_label.pack(side='left', padx=(0, 5))
        self.style_combobox.pack(side='left')

        # Progress frame
        self.progress_frame.pack(fill='x', padx=20, pady=10)
        self.progress_bar.pack(fill='x', pady=5)
        self.status_label.pack()

        # Button frame
        self.button_frame.pack(pady=15)
        self.start_button.pack(side='left', padx=5)
        self.exit_button.pack(side='left', padx=5)

    def browse_input(self) -> None:
        """Open file dialog to select input Excel file."""
        file_path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.input_path.set(file_path)

            # Auto-set output directory to 'output' folder in the same directory as input file
            default_output = Path(file_path).parent / 'output'
            self.output_path.set(str(default_output))

    def browse_output(self) -> None:
        """Open directory dialog to select output directory."""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_path.set(dir_path)

    def _setup_default_font(self) -> None:
        """Setup default font for the application."""
        # Use system default font for better appearance
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(size=11)
        self.root.option_add("*Font", default_font)

        # Also configure ttk styles
        style = ttk.Style()
        style.configure("TLabel", font=("TkDefaultFont", 11))
        style.configure("TButton", font=("TkDefaultFont", 11))
        style.configure("TCheckbutton", font=("TkDefaultFont", 11))
        style.configure("TEntry", font=("TkDefaultFont", 11))
        style.configure("TCombobox", font=("TkDefaultFont", 11))

    def _setup_icon(self) -> None:
        """Setup window icon if available."""
        # Try to load icon from various possible locations
        icon_paths = [
            Path(__file__).parent / "icon.png",
            Path(__file__).parent / "icon.ico",
            Path(__file__).parent / "resources" / "icon.png",
            Path(__file__).parent / "resources" / "icon.ico",
        ]

        for icon_path in icon_paths:
            if icon_path.exists():
                try:
                    if icon_path.suffix == ".ico":
                        # Windows ICO format
                        self.root.iconbitmap(str(icon_path))
                    else:
                        # PNG format (works on macOS and Windows)
                        self.root.iconphoto(False, tk.PhotoImage(file=str(icon_path)))
                    break
                except Exception:
                    # If icon loading fails, continue silently
                    pass

    def _on_style_select(self, event) -> None:
        """
        Convert display value to internal style mode value.

        Args:
            event: The combobox selection event
        """
        display_to_internal = {
            "统一样式（默认）": "unified",
            "保留原样式": "original",
            "无样式": "none"
        }
        current_display = self.style_mode.get()
        if current_display in display_to_internal:
            self.style_mode.set(display_to_internal[current_display])

    def update_status(self, message: str, progress: float = None) -> None:
        """
        Update status text and progress bar.

        Args:
            message: Status message to display
            progress: Progress percentage (0-100), optional
        """
        self.status_text.set(message)
        if progress is not None:
            self.progress_value.set(progress)
        self.root.update_idletasks()

    def validate_inputs(self) -> Tuple[Path, Path]:
        """
        Validate input and output paths.

        Returns:
            Tuple of (input_file, output_dir) as Path objects

        Raises:
            ValueError: If validation fails
        """
        input_str = self.input_path.get().strip()
        output_str = self.output_path.get().strip()

        if not input_str:
            raise ValueError("请选择输入文件")

        if not output_str:
            raise ValueError("请选择输出目录")

        input_file = Path(input_str)
        output_dir = Path(output_str)

        if not input_file.exists():
            raise ValueError(f"输入文件不存在: {input_file}")

        if not input_file.is_file():
            raise ValueError(f"输入路径不是文件: {input_file}")

        if input_file.suffix.lower() not in ['.xlsx', '.xlsm']:
            raise ValueError(f"输入文件必须是 .xlsx 或 .xlsm 格式")

        return input_file, output_dir

    def process_files(self, input_file: Path, output_dir: Path) -> None:
        """
        Process the Excel file in a separate thread.

        Args:
            input_file: Path to input Excel file
            output_dir: Path to output directory
        """
        try:
            # Create output directory
            output_dir.mkdir(parents=True, exist_ok=True)

            # Load workbook
            self.root.after(0, lambda: self.update_status("正在加载 Excel 文件...", 10))
            wb = load_workbook(str(input_file), data_only=True)

            # Analyze sheets
            self.root.after(0, lambda: self.update_status("正在分析表格结构...", 20))
            analyzer = SheetAnalyzer(wb)
            sheet_structures = analyzer.analyze_all_sheets()

            if not sheet_structures:
                raise ValueError("未找到包含'科室'列的表格")

            # Build department index (caches row indices)
            self.root.after(0, lambda: self.update_status("正在建立索引...", 30))
            dept_index = DepartmentIndex(wb, sheet_structures)
            dept_index.build_index()

            # Get departments from index
            all_departments = dept_index.get_departments()

            if not all_departments:
                raise ValueError("未找到任何科室数据")

            # Build workbooks using index
            self.root.after(0, lambda: self.update_status(f"找到 {len(all_departments)} 个科室，正在生成文件...", 40))
            # Convert display style mode to internal value
            style_mode_value = self._style_display_map.get(self.style_mode.get(), "unified")
            builder = WorkbookBuilder(
                wb, sheet_structures, output_dir, dept_index,
                remove_empty_sheets=self.remove_empty_sheets.get(),
                style_mode=style_mode_value
            )

            total = len(all_departments)
            success_count = 0

            for i, dept in enumerate(sorted(all_departments), 1):
                try:
                    output_file = builder.build_workbook_for_department(dept)
                    progress = 40 + (i / total * 50)
                    self.root.after(0, lambda p=progress, d=dept: self.update_status(f"正在处理: {d}", p))
                    success_count += 1
                except Exception as e:
                    print(f"Error creating file for {dept}: {e}")

            wb.close()

            # Complete
            self.root.after(0, lambda: self.update_status(f"完成！成功生成 {success_count}/{total} 个文件", 100))
            self.root.after(0, lambda: messagebox.showinfo(
                "完成",
                f"处理完成！\n\n成功生成 {success_count}/{total} 个科室文件\n\n输出目录: {output_dir}"
            ))

        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda: self.update_status(f"错误: {error_msg}", 0))
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))

        finally:
            self.root.after(0, lambda: self.start_button.config(state='normal'))

    def start_processing(self) -> None:
        """Start the processing in a separate thread."""
        try:
            input_file, output_dir = self.validate_inputs()
        except ValueError as e:
            messagebox.showerror("验证错误", str(e))
            return

        # Disable start button
        self.start_button.config(state='disabled')
        self.update_status("开始处理...", 0)

        # Start processing in a thread
        thread = threading.Thread(
            target=self.process_files,
            args=(input_file, output_dir)
        )
        thread.daemon = True
        thread.start()


def main():
    """Main entry point for the GUI application."""
    root = tk.Tk()
    app = FinanceSplitterGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
