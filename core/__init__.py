"""
Core module for finance Excel splitter.

This package provides core functionality for splitting Excel files by department.
"""

from .sheet_structure import SheetStructure
from .style_utils import copy_cell_style
from .sheet_analyzer import SheetAnalyzer
from .department_collector import DepartmentCollector
from .department_index import DepartmentIndex
from .sheet_filter import SheetFilter
from .workbook_builder import WorkbookBuilder

__all__ = [
    'SheetStructure',
    'copy_cell_style',
    'SheetAnalyzer',
    'DepartmentCollector',
    'DepartmentIndex',
    'SheetFilter',
    'WorkbookBuilder',
]
