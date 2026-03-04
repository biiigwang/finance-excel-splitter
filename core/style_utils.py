"""
Style utilities for copying cell styles between Excel files.
"""

from openpyxl.styles import Font, Border, Side, PatternFill, Alignment, Protection
from openpyxl.cell import Cell
from copy import copy


def copy_cell_style(source_cell: Cell, target_cell: Cell) -> None:
    """
    Copy all style attributes from source cell to target cell.

    Args:
        source_cell: The cell to copy styles from
        target_cell: The cell to copy styles to
    """
    if source_cell.has_style:
        # Copy font
        target_cell.font = copy(source_cell.font)

        # Copy border
        target_cell.border = copy(source_cell.border)

        # Copy fill
        target_cell.fill = copy(source_cell.fill)

        # Copy number format
        target_cell.number_format = copy(source_cell.number_format)

        # Copy protection
        target_cell.protection = copy(source_cell.protection)

        # Copy alignment
        target_cell.alignment = copy(source_cell.alignment)
