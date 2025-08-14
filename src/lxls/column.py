from __future__ import annotations
from typing import Iterator, TYPE_CHECKING
from . import NativeTypes

if TYPE_CHECKING:
    from .sheet import Sheet


class Column:
    """Represents a column in a worksheet with row access and slicing support."""
    
    def __init__(self, sheet: Sheet, column: str) -> None:
        """Initialize a column.
        
        Args:
            sheet: The parent sheet
            column: Column letter (A, B, C, etc.)
        """
        self.sheet = sheet
        self.column = column
    
    def __getitem__(self, row: int | slice) -> NativeTypes | Iterator[NativeTypes]:
        """Get value(s) from specific row(s).
        
        Args:
            row: Row number (1-based) or slice
            
        Returns:
            Cell value for int, iterator for slice
        """
        if isinstance(row, int):
            # Single row access
            if row < 1:
                raise ValueError("Row numbers must be 1-based (>= 1)")
            return self.sheet.get_cell_value(self.column, row)
        
        elif isinstance(row, slice):
            # Row range access
            start = row.start or 1
            stop = row.stop
            
            if start < 1:
                raise ValueError("Row numbers must be 1-based (>= 1)")
            if stop is not None and stop < start:
                raise ValueError("Slice stop must be >= start")
            
            def row_iterator() -> Iterator[NativeTypes]:
                current = start
                while stop is None or current < stop:
                    value = self.sheet.get_cell_value(self.column, current)
                    if value is None and stop is None:
                        # Stop iteration when we hit empty cells in open-ended slice
                        break
                    yield value
                    current += 1
            
            return row_iterator()
        
        else:
            raise TypeError(f"Row index must be int or slice, got {type(row)}")
    
    def __setitem__(self, row: int, value: NativeTypes) -> None:
        """Set value at specific row.
        
        Args:
            row: Row number (1-based)
            value: Value to set
        """
        if row < 1:
            raise ValueError("Row numbers must be 1-based (>= 1)")
        
        self.sheet.set_cell_value(self.column, row, value)
    
    def __iter__(self) -> Iterator[NativeTypes]:
        """Iterate over all values in the column."""
        # Find all rows that have data in this column
        rows_with_data = []
        for (col, row), value in self.sheet._cells.items():
            if col == self.column and value is not None:
                rows_with_data.append(row)
        
        if not rows_with_data:
            return iter([])
        
        # Sort rows and iterate from 1 to max row
        max_row = max(rows_with_data)
        
        for row in range(1, max_row + 1):
            value = self.sheet.get_cell_value(self.column, row)
            if value is not None:
                yield value
    
    def __len__(self) -> int:
        """Get number of non-empty cells in the column."""
        count = 0
        for (col, row), value in self.sheet._cells.items():
            if col == self.column and value is not None:
                count += 1
        return count