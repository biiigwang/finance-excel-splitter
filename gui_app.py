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
from tkinter import ttk, filedialog, messagebox

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
        self.root.title("财务数据拆分工具 - Finance Excel Splitter")
        self.root.geometry("600x400")
        self.root.minsize(500, 350)

        # Set window icon (if available)
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

        self.create_widgets()
        self.layout_widgets()

    def create_widgets(self) -> None:
        """Create all GUI widgets."""
        # Title label
        self.title_label = ttk.Label(
            self.root,
            text="财务数据拆分工具",
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

    def validate_inputs(self) -> tuple[Path, Path]:
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
            builder = WorkbookBuilder(wb, sheet_structures, output_dir, dept_index)

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
