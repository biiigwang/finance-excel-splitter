"""
Department collector module.

Collects all unique departments from all sheets in a workbook.
"""

from typing import Set, List, Dict
from openpyxl import Workbook
from .sheet_structure import SheetStructure


class DepartmentCollector:
    """
    Collects all unique departments from all sheets in a workbook.
    """

    def __init__(self, workbook: Workbook, sheet_structures: Dict[str, SheetStructure]):
        """
        Initialize the collector with a workbook and pre-analyzed structures.

        Args:
            workbook: The openpyxl workbook to read from
            sheet_structures: Dictionary of sheet structures from SheetAnalyzer
        """
        self.workbook = workbook
        self.sheet_structures = sheet_structures

    def collect_all_departments(self) -> Set[str]:
        """
        Collect all unique departments from all sheets.

        Returns:
            Set of unique department names
        """
        all_departments: Set[str] = set()

        for sheet_name, structure in self.sheet_structures.items():
            if not structure.has_data:
                continue

            worksheet = self.workbook[sheet_name]

            # Collect departments from this sheet
            for row_idx in range(structure.data_start_row, worksheet.max_row + 1):
                dept_value = worksheet.cell(
                    row=row_idx, column=structure.dept_col
                ).value

                if self._is_valid_department(dept_value):
                    all_departments.add(str(dept_value).strip())

        return all_departments

    def _is_valid_department(self, dept_value: str) -> bool:
        """
        Check if a department value is valid (not a header or number).

        Args:
            dept_value: The department value to check

        Returns:
            True if valid, False otherwise
        """
        if not dept_value:
            return False

        dept = str(dept_value).strip()

        # Skip empty strings
        if not dept:
            return False

        # Skip header names
        if dept in ['科室', '绩效科室', '序号']:
            return False

        # Skip pure numbers (like "1", "2", etc.)
        if dept.isdigit():
            return False

        return True

    def get_sorted_departments(self) -> List[str]:
        """
        Get all departments sorted alphabetically.

        Returns:
            Sorted list of department names
        """
        return sorted(list(self.collect_all_departments()))
