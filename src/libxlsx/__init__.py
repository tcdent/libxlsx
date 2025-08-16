from __future__ import annotations
from pathlib import Path

from .workbook import Workbook
from .sheet import Sheet
from .column import Column
from .types import formula, NativeTypes


def load_workbook(filename: str | Path) -> Workbook:
    """Load an Excel workbook from file.
    
    Args:
        filename: Path to the XLSX file
        
    Returns:
        Workbook object for surgical editing
    """
    return Workbook(Path(filename))
