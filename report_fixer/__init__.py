"""Printing Daily Report correction engine."""

from .processor import Sheet, Table, process_workbook
from .writer import write_workbook


def fix_report(data: bytes, zero_is_empty: bool = True) -> tuple[bytes, list[Sheet]]:
    """Return (corrected workbook bytes, processed sheets) for an uploaded report."""
    sheets = process_workbook(data, zero_is_empty=zero_is_empty)
    return write_workbook(sheets), sheets


__all__ = ["fix_report", "process_workbook", "write_workbook", "Sheet", "Table"]
