from __future__ import annotations
from typing import Iterator, TYPE_CHECKING
from lxml import etree
from .column import Column
from .utils import NAMESPACES, parse_cell_ref, get_cell_value_from_element, CellAttr, RowAttr

if TYPE_CHECKING:
    from .workbook import Workbook


class Sheet:
    """Represents a worksheet with column-based access."""
    
    def __init__(self, name: str, worksheet_path: str, workbook: Workbook) -> None:
        """Initialize a sheet.
        
        Args:
            name: Sheet name
            worksheet_path: Path to worksheet XML within ZIP
            workbook: Parent workbook
        """
        self.name = name
        self.worksheet_path = worksheet_path
        self.workbook = workbook
        self._cells: dict[tuple[str, int], str | int | float | bool | None] = {}
        self._columns: dict[str, Column] = {}
        
        # Load cell data
        self._load_cells()
    
    def _load_cells(self) -> None:
        """Load all cells from the worksheet XML."""
        worksheet_xml = self.workbook.get_zip_data(self.worksheet_path)
        if not worksheet_xml:
            return
        
        root = etree.fromstring(worksheet_xml)
        shared_strings = self.workbook.get_shared_strings()
        
        # Find all cell elements
        for cell_elem in root.xpath('//w:c', namespaces=NAMESPACES):
            cell_ref = cell_elem.get(CellAttr.REFERENCE)
            if not cell_ref:
                continue
            
            try:
                column, row = parse_cell_ref(cell_ref)
                value = get_cell_value_from_element(cell_elem, shared_strings)
                self._cells[(column, row)] = value
            except ValueError:
                # Skip invalid cell references
                continue
    
    def get_cell_value(self, column: str, row: int) -> str | int | float | bool | None:
        """Get value of a specific cell."""
        return self._cells.get((column, row))
    
    def set_cell_value(self, column: str, row: int, value: str | int | float | bool) -> None:
        """Set value of a specific cell."""
        # Store the value
        self._cells[(column, row)] = value
        
        # Update the worksheet XML
        self._update_worksheet_xml(column, row, value)
    
    def _update_worksheet_xml(self, column: str, row: int, value: str | int | float | bool) -> None:
        """Update the worksheet XML with new cell value."""
        from .utils import make_cell_ref, create_cell_element
        
        worksheet_xml = self.workbook.get_zip_data(self.worksheet_path)
        if not worksheet_xml:
            return
        
        root = etree.fromstring(worksheet_xml)
        cell_ref = make_cell_ref(column, row)
        
        # Find existing cell or create location for new one
        existing_cell = root.xpath(f'//w:c[@{CellAttr.REFERENCE}="{cell_ref}"]', namespaces=NAMESPACES)
        
        if existing_cell:
            # Update existing cell
            cell_elem = existing_cell[0]
            parent = cell_elem.getparent()
            if parent is not None:
                parent.remove(cell_elem)
        
        # Create new cell element
        new_cell = create_cell_element(value, cell_ref)
        
        # Find the correct row to insert into
        target_row = None
        for row_elem in root.xpath('//w:row', namespaces=NAMESPACES):
            row_num = int(row_elem.get(RowAttr.REFERENCE, '0'))
            if row_num == row:
                target_row = row_elem
                break
            elif row_num > row:
                # Need to create a new row before this one
                break
        
        if target_row is None:
            # Create new row
            sheet_data = root.xpath('//w:sheetData', namespaces=NAMESPACES)[0]
            target_row = etree.SubElement(sheet_data, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
            target_row.set(RowAttr.REFERENCE, str(row))
        
        # Insert cell in correct position (sorted by column)
        inserted = False
        for existing_cell in target_row.xpath('w:c', namespaces=NAMESPACES):
            existing_ref = existing_cell.get(CellAttr.REFERENCE, '')
            if existing_ref > cell_ref:
                target_row.insert(list(target_row).index(existing_cell), new_cell)
                inserted = True
                break
        
        if not inserted:
            target_row.append(new_cell)
        
        # Update workbook's ZIP data
        updated_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8')
        self.workbook.update_zip_data(self.worksheet_path, updated_xml)
    
    def __getitem__(self, column: str | slice) -> Column | Iterator[Column]:
        """Get a column by letter or range of columns by slice.
        
        Args:
            column: Column letter (A, B, C, etc.) or slice (A:B)
            
        Returns:
            Column object for str, iterator of Column objects for slice
        """
        if isinstance(column, str):
            # Single column
            if column not in self._columns:
                self._columns[column] = Column(self, column)
            return self._columns[column]
        
        elif isinstance(column, slice):
            # Column range (A:C)
            from .utils import column_range
            
            if column.start is None or column.stop is None:
                raise ValueError("Column slice must have both start and stop")
            
            def column_iterator() -> Iterator[Column]:
                for col in column_range(column.start, column.stop):
                    if col not in self._columns:
                        self._columns[col] = Column(self, col)
                    yield self._columns[col]
            
            return column_iterator()
        
        else:
            raise TypeError(f"Column index must be str or slice, got {type(column)}")