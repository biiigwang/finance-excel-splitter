"""
Sheet structure data class.

Defines the structure information for an Excel sheet.
"""

from dataclasses import dataclass
from typing import Optional


@dataclass
class SheetStructure:
    """
    Data class representing the structure of an Excel sheet.

    Attributes:
        sheet_name: Name of the sheet
        header_row: Row index of the header (1-based)
        dept_col: Column index of the department column (1-based)
        dept_col_letter: Column letter of the department column (e.g., 'B')
        data_start_row: First data row index (1-based)
        has_data: Whether the sheet has any data
    """

    sheet_name: str
    header_row: Optional[int] = None
    dept_col: Optional[int] = None
    dept_col_letter: Optional[str] = None
    data_start_row: Optional[int] = None
    has_data: bool = False

    def __post_init__(self):
        """Validate the structure after initialization."""
        if self.header_row is not None and self.data_start_row is None:
            self.data_start_row = self.header_row + 1
