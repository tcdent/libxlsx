"""
Test writing to an empty workbook to create the same format as our test fixture.
"""
import pytest
from pathlib import Path
from lxls import load_workbook, formula
from lxls.const import A, B, C


# Get fixture paths
EMPTY_FIXTURE_PATH = Path(__file__).parent / "fixtures" / "Book0.xlsx"
ORIGINAL_FIXTURE_PATH = Path(__file__).parent / "fixtures" / "Book1.xlsx"


@pytest.fixture
def empty_workbook():
    """Load the empty workbook for testing."""
    return load_workbook(EMPTY_FIXTURE_PATH)


@pytest.fixture
def original_workbook():
    """Load the original workbook for comparison."""
    return load_workbook(ORIGINAL_FIXTURE_PATH)


@pytest.fixture
def output_path():
    """Create output path in test/out directory."""
    output_dir = Path(__file__).parent / "out"
    output_dir.mkdir(exist_ok=True)
    return output_dir / "recreated_Book1.xlsx"


class TestWriteFromEmpty:
    """Test writing to an empty workbook to recreate our fixture."""
    
    def test_empty_workbook_loads(self, empty_workbook):
        """Test that we can load the empty workbook."""
        assert empty_workbook is not None
        
        # Should have at least one sheet
        sheet_names = list(empty_workbook._sheets.keys())
        assert len(sheet_names) > 0
        print(f"Empty workbook sheets: {sheet_names}")
    
    def test_explore_empty_workbook(self, empty_workbook):
        """Explore the empty workbook structure."""
        sheet_names = list(empty_workbook._sheets.keys())
        first_sheet = empty_workbook[sheet_names[0]]
        
        print(f"\nEmpty sheet: {first_sheet.name}")
        print("Existing cell contents:")
        
        # Check first few rows and columns
        found_any_data = False
        for row in range(1, 6):  # Rows 1-5
            for col in [A, B, C]:  # Columns A-C
                value = first_sheet.get_cell_value(col, row)
                if value is not None:
                    print(f"  {col}{row}: {repr(value)} ({type(value).__name__})")
                    found_any_data = True
        
        if not found_any_data:
            print("  No existing data found (as expected for empty workbook)")
    
    def test_recreate_fixture_data(self, empty_workbook, output_path):
        """Test recreating the same data as in our original fixture."""
        sheet_names = list(empty_workbook._sheets.keys())
        sheet = empty_workbook[sheet_names[0]]
        
        # Write the same values as in our original fixture (Book1.xlsx)
        sheet[A][1] = 123.45                    # Float value
        sheet[A][2] = "Hello world"             # String value  
        sheet[A][3] = formula("SUM(B2:B8)")     # Formula value (as in original)
        
        # Verify the values were written correctly
        assert sheet[A][1] == 123.45
        assert sheet[A][2] == "Hello world" 
        assert isinstance(sheet[A][3], formula)
        assert str(sheet[A][3]) == "SUM(B2:B8)"
        
        # Save the recreated file to test/out directory
        empty_workbook.save(output_path)
        
        # Verify file was created
        assert output_path.exists()
        print(f"\n✅ Created recreated Book1.xlsx at: {output_path}")
        print(f"📁 You can now open this file to verify it matches the original Book1.xlsx fixture")
    
    def test_verify_recreated_vs_original(self, empty_workbook, original_workbook, output_path):
        """Test that our recreated fixture matches the original."""
        # Write to empty workbook
        empty_sheet_names = list(empty_workbook._sheets.keys())
        empty_sheet = empty_workbook[empty_sheet_names[0]]
        
        empty_sheet[A][1] = 123.45
        empty_sheet[A][2] = "Hello world"
        empty_sheet[A][3] = formula("SUM(B2:B8)")
        
        # Save and reload
        empty_workbook.save(output_path)
        recreated_workbook = load_workbook(output_path)
        
        # Compare with original
        original_sheet_names = list(original_workbook._sheets.keys())
        original_sheet = original_workbook[original_sheet_names[0]]
        
        recreated_sheet_names = list(recreated_workbook._sheets.keys())
        recreated_sheet = recreated_workbook[recreated_sheet_names[0]]
        
        # Compare the key cell values
        test_cells = [(A, 1), (A, 2), (A, 3)]
        
        print("\nComparing original vs recreated:")
        for col, row in test_cells:
            original_val = original_sheet.get_cell_value(col, row)
            recreated_val = recreated_sheet.get_cell_value(col, row)
            
            print(f"  {col}{row}: original={repr(original_val)}, recreated={repr(recreated_val)}")
            
            # Values should match
            assert original_val == recreated_val, f"Mismatch at {col}{row}: {original_val} != {recreated_val}"
    
    def test_add_more_complex_data(self, empty_workbook):
        """Test adding more complex data types to validate our writing capabilities."""
        sheet_names = list(empty_workbook._sheets.keys())
        sheet = empty_workbook[sheet_names[0]]
        
        # Add various data types across multiple columns
        sheet[A][10] = "String Test"
        sheet[B][10] = 42
        sheet[C][10] = 3.14159
        sheet[A][11] = True
        sheet[B][11] = False  
        sheet[C][11] = 0
        
        # Add a formula
        sheet[A][12] = formula("=SUM(B10:B11)")
        
        # Add some negative numbers
        sheet[B][12] = -100
        sheet[C][12] = -0.5
        
        # Save and reload to verify persistence  
        output_dir = Path(__file__).parent / "out"
        complex_output_path = output_dir / "complex_data_test.xlsx"
        empty_workbook.save(complex_output_path)
        reloaded_workbook = load_workbook(complex_output_path)
        reloaded_sheet = reloaded_workbook[sheet_names[0]]
        
        # Verify all values persisted correctly
        assert reloaded_sheet[A][10] == "String Test"
        assert reloaded_sheet[B][10] == 42
        assert reloaded_sheet[C][10] == 3.14159
        assert reloaded_sheet[A][11] is True
        assert reloaded_sheet[B][11] is False
        assert reloaded_sheet[C][11] == 0
        # Formula should be preserved as a formula object when reloaded
        formula_value = reloaded_sheet[A][12]
        assert isinstance(formula_value, formula)
        assert str(formula_value) == "=SUM(B10:B11)"
        assert reloaded_sheet[B][12] == -100
        assert reloaded_sheet[C][12] == -0.5
        
        print("All complex data types persisted correctly!")
    
    def test_column_operations_on_empty_workbook(self, empty_workbook):
        """Test that column operations work on initially empty workbook."""
        sheet_names = list(empty_workbook._sheets.keys())
        sheet = empty_workbook[sheet_names[0]]
        
        # Add data to a column
        test_values = ["First", "Second", "Third", "Fourth"]
        for i, value in enumerate(test_values, start=20):
            sheet[B][i] = value
        
        # Test column iteration
        column_b = sheet[B]
        values = list(column_b)
        
        # Should include our test values
        for test_val in test_values:
            assert test_val in values
        
        # Test column slicing  
        slice_values = list(column_b[20:23])
        assert "First" in slice_values
        assert "Second" in slice_values
        assert "Third" in slice_values
        # "Fourth" should not be included (exclusive end)
        
        # Test column length
        assert len(column_b) >= len(test_values)
        
        print(f"Column B operations successful: {len(column_b)} total cells")


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])