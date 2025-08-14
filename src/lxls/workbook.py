from __future__ import annotations
import zipfile
from pathlib import Path
from lxml import etree
from .sheet import Sheet
from .utils import NAMESPACES


class Workbook:
    """Represents an Excel workbook with sheet access and save functionality."""
    
    def __init__(self, filename: Path) -> None:
        """Initialize a workbook.
        
        Args:
            filename: Path to the Excel file
        """
        self.filename = Path(filename)
        self._zip_data: dict[str, bytes] = {}
        self._sheets: dict[str, Sheet] = {}
        self._shared_strings: list[str] = []
        
        # Load the XLSX file into memory
        self._load_workbook()
    
    def _load_workbook(self) -> None:
        """Load XLSX file into memory and parse structure."""
        if not self.filename.exists():
            raise FileNotFoundError(f"Workbook file not found: {self.filename}")
        
        # Load entire ZIP into memory
        with zipfile.ZipFile(self.filename, 'r') as zip_file:
            for name in zip_file.namelist():
                self._zip_data[name] = zip_file.read(name)
        
        # Parse workbook relationships and sheet info
        self._parse_workbook_structure()
        self._load_shared_strings()
    
    def _parse_workbook_structure(self) -> None:
        """Parse workbook.xml to get sheet information."""
        workbook_xml = self._zip_data.get('xl/workbook.xml')
        if not workbook_xml:
            raise ValueError("Invalid XLSX file: missing xl/workbook.xml")
        
        root = etree.fromstring(workbook_xml)
        
        # Find all sheets
        for sheet_elem in root.xpath('//w:sheet', namespaces=NAMESPACES):
            sheet_name = sheet_elem.get('name')
            sheet_id = sheet_elem.get('sheetId')
            rel_id = sheet_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            
            if sheet_name and sheet_id and rel_id:
                # Create sheet object - it will load its own data
                worksheet_path = self._get_worksheet_path(rel_id)
                if worksheet_path:
                    self._sheets[sheet_name] = Sheet(sheet_name, worksheet_path, self)
    
    def _get_worksheet_path(self, rel_id: str) -> str | None:
        """Get worksheet path from relationship ID."""
        rels_xml = self._zip_data.get('xl/_rels/workbook.xml.rels')
        if not rels_xml:
            return None
        
        root = etree.fromstring(rels_xml)
        
        for rel in root.xpath(f'//r:Relationship[@Id="{rel_id}"]', namespaces=NAMESPACES):
            target = rel.get('Target')
            if target:
                return f'xl/{target}'
        
        return None
    
    def _load_shared_strings(self) -> None:
        """Load shared strings table."""
        shared_strings_xml = self._zip_data.get('xl/sharedStrings.xml')
        if not shared_strings_xml:
            return  # No shared strings
        
        root = etree.fromstring(shared_strings_xml)
        
        for si in root.xpath('//w:si', namespaces=NAMESPACES):
            # Handle both simple text and rich text
            t_elem = si.find('w:t', NAMESPACES)
            if t_elem is not None and t_elem.text:
                self._shared_strings.append(t_elem.text)
            else:
                # Handle rich text by concatenating all text elements
                text_parts = []
                for t in si.xpath('.//w:t', namespaces=NAMESPACES):
                    if t.text:
                        text_parts.append(t.text)
                self._shared_strings.append(''.join(text_parts))
    
    def __getitem__(self, sheet_name: str) -> Sheet:
        """Get a sheet by name.
        
        Args:
            sheet_name: Name of the sheet
            
        Returns:
            Sheet object
        """
        if sheet_name not in self._sheets:
            raise KeyError(f"Sheet '{sheet_name}' not found in workbook")
        
        return self._sheets[sheet_name]
    
    def get_zip_data(self, path: str) -> bytes | None:
        """Get data for a specific ZIP member."""
        return self._zip_data.get(path)
    
    def update_zip_data(self, path: str, data: bytes) -> None:
        """Update data for a specific ZIP member."""
        self._zip_data[path] = data
    
    def get_shared_strings(self) -> list[str]:
        """Get the shared strings table."""
        return self._shared_strings
    
    def save(self, filename: str | Path | None = None) -> None:
        """Save the workbook to a file.
        
        Args:
            filename: Path to save the workbook (defaults to original filename)
        """
        target_path = Path(filename) if filename else self.filename
        
        # Create new ZIP with updated data
        with zipfile.ZipFile(target_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for path, data in self._zip_data.items():
                zip_file.writestr(path, data)
