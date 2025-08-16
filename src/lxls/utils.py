from __future__ import annotations
import re
from typing import Iterator
from lxml import etree

# Excel column letters: A-Z, AA-ZZ, etc.
COLUMN_PATTERN = re.compile(r'^[A-Z]+$')
CELL_PATTERN = re.compile(r'^([A-Z]+)(\d+)$')

# XLSX XML namespaces
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'pkg': 'http://schemas.openxmlformats.org/package/2006/relationships'
}

# Cell type constants
class CellType:
    """Excel cell type constants."""
    SHARED_STRING = 's'
    BOOLEAN = 'b'
    INLINE_STRING = 'str'
    ERROR = 'e'
    FORMULA = 'f'
    NUMBER = 'n'  # Default when no type specified


# XML attribute constants
class CellAttr:
    """Excel cell XML attribute constants."""
    REFERENCE = 'r'  # Cell reference like "A1", "B2"
    TYPE = 't'       # Cell type
    STYLE = 's'      # Cell style index


class RowAttr:
    """Excel row XML attribute constants."""
    REFERENCE = 'r'  # Row number


def column_to_index(col: str) -> int:
    """Convert Excel column letter(s) to 0-based index.
    
    A=0, B=1, ..., Z=25, AA=26, AB=27, etc.
    """
    if not COLUMN_PATTERN.match(col):
        raise ValueError(f"Invalid column: {col}")
    
    result = 0
    for char in col:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1


def index_to_column(index: int) -> str:
    """Convert 0-based index to Excel column letter(s).
    
    0=A, 1=B, ..., 25=Z, 26=AA, 27=AB, etc.
    """
    if index < 0:
        raise ValueError(f"Column index must be non-negative: {index}")
    
    result = ""
    index += 1  # Convert to 1-based
    while index > 0:
        index -= 1
        result = chr(ord('A') + (index % 26)) + result
        index //= 26
    return result


def parse_cell_ref(cell_ref: str) -> tuple[str, int]:
    """Parse cell reference like 'A1' into column and row.
    
    Returns:
        Tuple of (column, row) where row is 1-based
    """
    match = CELL_PATTERN.match(cell_ref.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    
    return match.group(1), int(match.group(2))


def make_cell_ref(column: str, row: int) -> str:
    """Create cell reference from column and row."""
    return f"{column}{row}"


def column_range(start: str, end: str) -> Iterator[str]:
    """Generate column letters from start to end (inclusive).
    
    Example: column_range('A', 'C') yields 'A', 'B', 'C'
    """
    start_idx = column_to_index(start)
    end_idx = column_to_index(end)
    
    for i in range(start_idx, end_idx + 1):
        yield index_to_column(i)


def get_cell_value_from_element(cell_elem: etree._Element, shared_strings: list[str] | None = None) -> str | int | float | bool | None:
    """Extract cell value from XML element.
    
    Handles different cell types: string, number, boolean, formulas, etc.
    """
    from .types import formula
    
    # Check for formula first
    formula_elem = cell_elem.find('w:f', NAMESPACES)
    if formula_elem is not None and formula_elem.text:
        return formula(formula_elem.text)
    
    # Get cell type
    cell_type = cell_elem.get(CellAttr.TYPE, CellType.NUMBER)
    
    # Find value element
    value_elem = cell_elem.find('w:v', NAMESPACES)
    if value_elem is None:
        return None
    
    value_text = value_elem.text
    if not value_text:
        return None
    
    # Handle different cell types
    if cell_type == CellType.SHARED_STRING:
        if shared_strings is None:
            return f"<shared_string_{value_text}>"  # Placeholder for now
        try:
            return shared_strings[int(value_text)]
        except (IndexError, ValueError):
            return f"<invalid_shared_string_{value_text}>"
    
    elif cell_type == CellType.BOOLEAN:
        return value_text == '1'
    
    elif cell_type == CellType.INLINE_STRING:
        return value_text
    
    else:  # Number (default)
        try:
            # Try integer first
            if '.' not in value_text:
                return int(value_text)
            return float(value_text)
        except ValueError:
            return value_text  # Return as string if parsing fails


def create_cell_element(value: str | int | float | bool, cell_ref: str) -> etree._Element:
    """Create XML element for a cell with the given value."""
    from .types import formula
    
    cell = etree.Element('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
    cell.set(CellAttr.REFERENCE, cell_ref)
    
    if isinstance(value, formula):
        # Create formula element (no type attribute, no value element)
        f_elem = etree.SubElement(cell, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f')
        f_elem.text = str(value)
    elif isinstance(value, bool):
        # Create value element for boolean
        v_elem = etree.SubElement(cell, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
        cell.set(CellAttr.TYPE, CellType.BOOLEAN)
        v_elem.text = '1' if value else '0'
    elif isinstance(value, (int, float)):
        # Create value element for numbers (no type attribute needed)
        v_elem = etree.SubElement(cell, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
        v_elem.text = str(value)
    else:  # String - store as inline string
        v_elem = etree.SubElement(cell, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
        cell.set(CellAttr.TYPE, CellType.INLINE_STRING)
        v_elem.text = str(value)
    
    return cell