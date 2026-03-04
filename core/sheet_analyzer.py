"""
Sheet analyzer module.

Analyzes Excel sheets to find department columns and data structure.
"""

import re
from typing import Optional, List, Dict
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from .sheet_structure import SheetStructure


class SheetAnalyzer:
    """
    Analyzes Excel sheets to identify department columns and data structure.
    """

    # Pattern to match department column headers (exact match or common variations)
    DEPT_HEADERS = {'科室', '绩效科室', '所属科室', '部门'}

    def __init__(self, workbook: Workbook):
        """
        Initialize the analyzer with a workbook.

        Args:
            workbook: The openpyxl workbook to analyze
        """
        self.workbook = workbook

    def analyze_all_sheets(self) -> Dict[str, SheetStructure]:
        """
        Analyze all sheets in the workbook.

        Returns:
            Dictionary mapping sheet names to their SheetStructure
        """
        structures = {}

        for sheet_name in self.workbook.sheetnames:
            worksheet = self.workbook[sheet_name]
            structure = self._analyze_sheet(worksheet)
            if structure.has_data:
                structures[sheet_name] = structure

        return structures

    def _analyze_sheet(self, worksheet: Worksheet) -> SheetStructure:
        """
        Analyze a single worksheet.

        Args:
            worksheet: The worksheet to analyze

        Returns:
            SheetStructure object containing the analysis results
        """
        structure = SheetStructure(sheet_name=worksheet.title)

        # Find the header row and department column
        header_row, dept_col = self._find_header_and_dept_col(worksheet)

        if header_row and dept_col:
            structure.header_row = header_row
            structure.dept_col = dept_col
            structure.dept_col_letter = self._get_column_letter(dept_col)
            structure.data_start_row = header_row + 1
            structure.has_data = True

        return structure

    def _find_header_and_dept_col(self, worksheet: Worksheet) -> tuple[Optional[int], Optional[int]]:
        """
        Find the header row and department column in the sheet.

        Args:
            worksheet: The worksheet to analyze

        Returns:
            Tuple of (header_row, dept_col) or (None, None) if not found
        """
        max_rows_to_check = min(10, worksheet.max_row)

        for row_idx in range(1, max_rows_to_check + 1):
            for col_idx in range(1, worksheet.max_column + 1):
                cell_value = worksheet.cell(row=row_idx, column=col_idx).value

                # Check for exact match with known department headers
                if cell_value and str(cell_value).strip() in self.DEPT_HEADERS:
                    return row_idx, col_idx

        return None, None

    def _get_column_letter(self, col_idx: int) -> str:
        """
        Convert column index to column letter.

        Args:
            col_idx: 1-based column index

        Returns:
            Column letter (e.g., 'A', 'B', 'AA')
        """
        result = ""
        while col_idx > 0:
            col_idx, remainder = divmod(col_idx - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def get_department_values(self, worksheet: Worksheet, structure: SheetStructure) -> List[str]:
        """
        Get all unique department values from the sheet.

        Args:
            worksheet: The worksheet to read from
            structure: The sheet structure containing department column info

        Returns:
            List of unique department names
        """
        if not structure.has_data or not structure.dept_col:
            return []

        departments = set()
        for row_idx in range(structure.data_start_row, worksheet.max_row + 1):
            dept_value = worksheet.cell(
                row=row_idx, column=structure.dept_col
            ).value

            if dept_value and str(dept_value).strip():
                departments.add(str(dept_value).strip())

        return sorted(list(departments))
