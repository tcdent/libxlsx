"""
Basic functionality tests for libxlsx using the Book1.xlsx fixture.
"""
import pytest
from pathlib import Path
from libxlsx import load_workbook, formula
from libxlsx.const import A, B, C, D


# Get fixture path
FIXTURE_PATH = Path(__file__).parent / "fixtures" / "Book1.xlsx"


@pytest.fixture
def sample_workbook():
    """Load the sample workbook for testing."""
    return load_workbook(FIXTURE_PATH)


@pytest.fixture
def temp_workbook_path(tmp_path):
    """Create a temporary path for saving test workbooks."""
    return tmp_path / "test_output.xlsx"


class TestBasicLoading:
    """Test basic workbook and sheet loading."""
    
    def test_load_workbook(self, sample_workbook):
        """Test that we can load a workbook."""
        assert sample_workbook is not None
        assert hasattr(sample_workbook, 'filename')
        assert sample_workbook.filename.exists()
    
    def test_get_sheet_names(self, sample_workbook):
        """Test that we can access sheets."""
        # Let's discover what sheets exist
        sheet_names = list(sample_workbook._sheets.keys())
        assert len(sheet_names) > 0
        print(f"Available sheets: {sheet_names}")
    
    def test_access_sheet(self, sample_workbook):
        """Test that we can access a sheet."""
        # Try common sheet names
        sheet_names = list(sample_workbook._sheets.keys())
        first_sheet_name = sheet_names[0]
        
        sheet = sample_workbook[first_sheet_name]
        assert sheet is not None
        assert sheet.name == first_sheet_name


class TestCellReading:
    """Test reading cell values from the fixture."""
    
    def test_explore_fixture_content(self, sample_workbook):
        """Explore what's in our fixture to understand the data structure."""
        sheet_names = list(sample_workbook._sheets.keys())
        first_sheet = sample_workbook[sheet_names[0]]
        
        print(f"\nSheet: {first_sheet.name}")
        print("Cell contents:")
        
        # Print first few rows and columns to see what we have
        for row in range(1, 6):  # Rows 1-5
            for col in [A, B, C, D]:  # Columns A-D
                value = first_sheet.get_cell_value(col, row)
                if value is not None:
                    print(f"  {col}{row}: {repr(value)} ({type(value).__name__})")
    
    def test_read_specific_cells(self, sample_workbook):
        """Test reading specific cell values."""
        sheet_names = list(sample_workbook._sheets.keys())
        sheet = sample_workbook[sheet_names[0]]
        
        # Test accessing cells via our API
        value_a1 = sheet[A][1]
        value_b1 = sheet[B][1]
        
        # These might be None if the cells are empty, which is fine
        print(f"A1: {repr(value_a1)}")
        print(f"B1: {repr(value_b1)}")


class TestCellWriting:
    """Test writing cell values."""
    
    def test_write_cell_values(self, sample_workbook):
        """Test writing different types of values to cells."""
        sheet_names = list(sample_workbook._sheets.keys())
        sheet = sample_workbook[sheet_names[0]]
        
        # Test writing different data types
        sheet[A][10] = "Test String"
        sheet[B][10] = 42
        sheet[C][10] = 3.14159
        sheet[D][10] = True
        
        # Verify the values were written
        assert sheet[A][10] == "Test String"
        assert sheet[B][10] == 42
        assert sheet[C][10] == 3.14159
        assert sheet[D][10] is True
    
    def test_write_formula(self, sample_workbook):
        """Test writing a formula."""
        sheet_names = list(sample_workbook._sheets.keys())
        sheet = sample_workbook[sheet_names[0]]
        
        # Write some numbers to sum
        sheet[A][20] = 10
        sheet[A][21] = 20
        
        # Write a formula
        sheet[B][22] = formula("SUM(A20:A21)")
        
        # The formula should be stored as a formula object
        result = sheet[B][22]
        assert isinstance(result, formula)
        assert str(result) == "SUM(A20:A21)"


class TestColumnOperations:
    """Test column-based operations."""
    
    def test_column_iteration(self, sample_workbook):
        """Test iterating over a column."""
        sheet_names = list(sample_workbook._sheets.keys())
        sheet = sample_workbook[sheet_names[0]]
        
        # Add some test data
        test_values = ["First", "Second", "Third"]
        for i, value in enumerate(test_values, start=30):
            sheet[A][i] = value
        
        # Test iteration
        column_a = sheet[A]
        values = list(column_a)
        
        # Should include our test values
        assert "First" in values
        assert "Second" in values
        assert "Third" in values
    
    def test_column_slicing(self, sample_workbook):
        """Test slicing a column."""
        sheet_names = list(sample_workbook._sheets.keys())
        sheet = sample_workbook[sheet_names[0]]
        
        # Add test data
        for i in range(40, 45):
            sheet[B][i] = f"Row{i}"
        
        # Test slicing
        column_b = sheet[B]
        slice_values = list(column_b[40:43])
        
        assert "Row40" in slice_values
        assert "Row41" in slice_values
        assert "Row42" in slice_values
        # Row43 and beyond should not be included (exclusive end)
    
    def test_column_range(self, sample_workbook):
        """Test iterating over multiple columns."""
        sheet_names = list(sample_workbook._sheets.keys())
        sheet = sample_workbook[sheet_names[0]]
        
        # Add test data across columns
        sheet[A][50] = "Col A"
        sheet[B][50] = "Col B"
        sheet[C][50] = "Col C"
        
        # Test column range iteration
        columns = list(sheet[A:C])
        assert len(columns) == 3
        
        # Each should be a Column object
        for col in columns:
            assert hasattr(col, 'column')
            assert hasattr(col, 'sheet')


class TestSaveLoad:
    """Test saving and loading workbooks."""
    
    def test_save_workbook(self, sample_workbook, temp_workbook_path):
        """Test saving a modified workbook."""
        sheet_names = list(sample_workbook._sheets.keys())
        sheet = sample_workbook[sheet_names[0]]
        
        # Make some changes
        sheet[A][100] = "Saved Data"
        sheet[B][100] = 999
        
        # Save the workbook
        sample_workbook.save(temp_workbook_path)
        
        # Verify file was created
        assert temp_workbook_path.exists()
    
    def test_save_and_reload(self, sample_workbook, temp_workbook_path):
        """Test that saved data persists when reloading."""
        sheet_names = list(sample_workbook._sheets.keys())
        sheet = sample_workbook[sheet_names[0]]
        
        # Add unique test data
        test_string = "Persistent Data"
        test_number = 12345
        
        sheet[A][200] = test_string
        sheet[B][200] = test_number
        
        # Save the workbook
        sample_workbook.save(temp_workbook_path)
        
        # Load the saved workbook
        reloaded_workbook = load_workbook(temp_workbook_path)
        reloaded_sheet = reloaded_workbook[sheet_names[0]]
        
        # Verify our data persisted
        assert reloaded_sheet[A][200] == test_string
        assert reloaded_sheet[B][200] == test_number


class TestErrorHandling:
    """Test error conditions and edge cases."""
    
    def test_invalid_sheet_access(self, sample_workbook):
        """Test accessing a non-existent sheet."""
        with pytest.raises(KeyError):
            sample_workbook["NonExistentSheet"]
    
    def test_invalid_row_numbers(self, sample_workbook):
        """Test invalid row numbers."""
        sheet_names = list(sample_workbook._sheets.keys())
        sheet = sample_workbook[sheet_names[0]]
        
        with pytest.raises(ValueError, match="Row numbers must be 1-based"):
            sheet[A][0]  # Row 0 should be invalid
        
        with pytest.raises(ValueError, match="Row numbers must be 1-based"):
            sheet[A][-1]  # Negative rows should be invalid


if __name__ == "__main__":
    # Run tests with verbose output to see the exploration
    pytest.main([__file__, "-v", "-s"])