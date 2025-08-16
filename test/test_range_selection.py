"""
Comprehensive tests for column and row range selection with mixed data types.
"""
import pytest
from pathlib import Path
from libxlsx import load_workbook, formula
from libxlsx.const import A, B, C, D, E, F


@pytest.fixture
def rich_workbook():
    """Load the rich data fixture for range testing."""
    return load_workbook(Path(__file__).parent / "fixtures" / "Book2.xlsx")


class TestColumnRangeSelection:
    """Test column range selection with mixed data types."""
    
    def test_simple_column_range(self, rich_workbook):
        """Test basic column range A:C."""
        sheet = rich_workbook["Sheet1"]
        
        # Test A:C range
        columns = list(sheet[A:C])
        assert len(columns) == 3
        
        # Verify each column is correct type and has expected data
        col_a, col_b, col_c = columns
        
        assert col_a.column == "A"
        assert col_b.column == "B"  
        assert col_c.column == "C"
        
        # Test that each column has the expected data
        assert col_a[1] == 1.0          # Float
        assert col_b[1] == "Alpha"       # String
        assert col_c[1] == 10           # Integer
    
    def test_larger_column_range(self, rich_workbook):
        """Test larger column range B:F."""
        sheet = rich_workbook["Sheet1"]
        
        columns = list(sheet[B:F])
        assert len(columns) == 5  # B, C, D, E, F
        
        column_letters = [col.column for col in columns]
        assert column_letters == ["B", "C", "D", "E", "F"]
        
        # Test accessing different data types across the range
        for col in columns:
            # Each column should have at least some data
            assert len(col) > 0
    
    def test_column_range_iteration_mixed_data(self, rich_workbook):
        """Test iterating over column ranges with mixed data types."""
        sheet = rich_workbook["Sheet1"]
        
        # Collect all values from columns A:D
        all_values = []
        for col in sheet[A:D]:
            for cell_value in col:
                all_values.append(cell_value)
        
        # Should include various data types
        has_float = any(isinstance(v, float) for v in all_values)
        has_string = any(isinstance(v, str) for v in all_values) 
        has_int = any(isinstance(v, int) for v in all_values)
        has_bool = any(isinstance(v, bool) for v in all_values)
        has_formula = any(isinstance(v, formula) for v in all_values)
        
        assert has_float, "Should have float values"
        assert has_string, "Should have string values" 
        assert has_int, "Should have integer values"
        assert has_bool, "Should have boolean values"
        assert has_formula, "Should have formula values"
        
        print(f"Found {len(all_values)} total values across columns A:D")
        print(f"Data types: float={has_float}, string={has_string}, int={has_int}, bool={has_bool}, formula={has_formula}")
    
    def test_single_column_in_range_context(self, rich_workbook):
        """Test that single column access still works when range is also available."""
        sheet = rich_workbook["Sheet1"]
        
        # Single column access
        col_c = sheet[C]
        assert col_c.column == "C"
        assert col_c[1] == 10
        assert col_c[2] == 20.5
        
        # Range access  
        columns = list(sheet[C:E])
        assert len(columns) == 3  # C, D, E
        
        # Both should work independently
        assert col_c[3] == 30
        assert columns[0][3] == 30  # Same cell via range


class TestRowRangeSelection:
    """Test row range selection within columns."""
    
    def test_row_range_in_single_column(self, rich_workbook):
        """Test row slicing within a single column."""
        sheet = rich_workbook["Sheet1"]
        
        # Test different row ranges in column A
        col_a = sheet[A]
        
        # Test 1:3 range (should get rows 1 and 2, exclusive end)
        rows_1_to_3 = list(col_a[1:3])
        assert len(rows_1_to_3) == 2
        assert rows_1_to_3[0] == 1.0    # A1
        assert rows_1_to_3[1] == 2.5    # A2
        
        # Test 2:5 range  
        rows_2_to_5 = list(col_a[2:5])
        assert len(rows_2_to_5) == 3
        assert rows_2_to_5[0] == 2.5    # A2
        assert rows_2_to_5[1] == 3.0    # A3
        assert rows_2_to_5[2] == 4.7    # A4
    
    def test_row_range_with_formulas(self, rich_workbook):
        """Test row ranges in columns containing formulas."""
        sheet = rich_workbook["Sheet1"]
        
        # Column C has formulas in rows 4 and 5
        col_c = sheet[C]
        
        # Test range that includes formulas
        rows_3_to_6 = list(col_c[3:6])
        assert len(rows_3_to_6) == 3
        assert rows_3_to_6[0] == 30                    # C3 (int)
        assert isinstance(rows_3_to_6[1], formula)     # C4 (formula)
        assert isinstance(rows_3_to_6[2], formula)     # C5 (formula)
        
        # Test the actual formula content
        assert str(rows_3_to_6[1]) == "A4*10"
        assert str(rows_3_to_6[2]) == "SUM(C1:C3)"
    
    def test_row_range_with_strings(self, rich_workbook):
        """Test row ranges in string columns."""
        sheet = rich_workbook["Sheet1"]
        
        # Column B has string data
        col_b = sheet[B]
        
        # Test range of strings
        text_rows = list(col_b[2:5])
        assert len(text_rows) == 3
        assert text_rows[0] == "Beta"     # B2
        assert text_rows[1] == "Gamma"    # B3  
        assert text_rows[2] == "Delta"    # B4
        
        # All should be strings
        for value in text_rows:
            assert isinstance(value, str)
    
    def test_row_range_with_mixed_types(self, rich_workbook):
        """Test row ranges in columns with mixed data types."""
        sheet = rich_workbook["Sheet1"]
        
        # Column D has mixed types: bool, bool, string, float, formula
        col_d = sheet[D]
        
        mixed_range = list(col_d[1:6])
        assert len(mixed_range) == 5
        
        assert mixed_range[0] is True        # D1 (bool)
        assert mixed_range[1] is False       # D2 (bool)
        assert mixed_range[2] == "Mixed"     # D3 (string)
        assert mixed_range[3] == 99.99       # D4 (float)
        assert isinstance(mixed_range[4], formula)  # D5 (formula)
        
        print(f"Mixed range types: {[type(v).__name__ for v in mixed_range]}")
    
    def test_open_ended_row_ranges(self, rich_workbook):
        """Test open-ended row ranges (no stop specified)."""
        sheet = rich_workbook["Sheet1"]
        
        col_a = sheet[A]
        
        # Open-ended range from row 3 onwards
        # Note: Our implementation stops at first None value for open-ended ranges
        open_range = list(col_a[3:])
        
        # Should include rows 3, 4, 5 but stop at the gap (row 6 is None)
        assert 3.0 in open_range      # A3
        assert 4.7 in open_range      # A4  
        assert 5.0 in open_range      # A5
        
        # Should NOT include rows 10-15 because we stop at first gap
        assert "Row10" not in open_range  # Stops before reaching A10
        
        # Length should be 3 (rows 3, 4, 5)
        assert len(open_range) == 3
        
        print(f"Open-ended range found {len(open_range)} values (stops at first gap)")
        
        # Test that we can still access the later data with explicit ranges
        explicit_range = list(col_a[10:16])  # Explicit range to get rows 10-15
        assert "Row10" in explicit_range
        assert "Row15" in explicit_range
        print(f"Explicit range (10:16) found {len(explicit_range)} values")


class TestRangeCombinations:
    """Test combinations of column and row ranges."""
    
    def test_column_range_with_row_slicing(self, rich_workbook):
        """Test column ranges combined with row slicing."""
        sheet = rich_workbook["Sheet1"]
        
        # Get columns B:D and slice rows 1:4 in each
        column_data = {}
        for col in sheet[B:D]:
            column_data[col.column] = list(col[1:4])
        
        # Verify we got the right columns
        assert set(column_data.keys()) == {"B", "C", "D"}
        
        # Verify row slice length
        for col_letter, values in column_data.items():
            assert len(values) == 3  # Rows 1, 2, 3 (exclusive end)
        
        # Check specific values
        assert column_data["B"][0] == "Alpha"    # B1
        assert column_data["B"][1] == "Beta"     # B2
        assert column_data["B"][2] == "Gamma"    # B3
        
        assert column_data["C"][0] == 10         # C1  
        assert column_data["C"][1] == 20.5       # C2
        assert column_data["C"][2] == 30         # C3
        
        print(f"Column B:D rows 1:4 data: {column_data}")
    
    def test_nested_iteration_patterns(self, rich_workbook):
        """Test complex nested iteration patterns."""
        sheet = rich_workbook["Sheet1"]
        
        # Pattern 1: Iterate columns, then rows within each column
        pattern1_data = []
        for col in sheet[A:C]:
            col_values = []
            for row_value in col[1:4]:  # First 3 rows
                col_values.append(row_value)
            pattern1_data.append((col.column, col_values))
        
        assert len(pattern1_data) == 3  # A, B, C
        assert pattern1_data[0][0] == "A"
        assert pattern1_data[0][1] == [1.0, 2.5, 3.0]  # A1:A3
        assert pattern1_data[1][0] == "B"  
        assert pattern1_data[1][1] == ["Alpha", "Beta", "Gamma"]  # B1:B3
        
        # Pattern 2: Collect specific cells across column range
        pattern2_data = []
        for col in sheet[D:F]:
            row2_value = col[2]  # Get row 2 from each column
            pattern2_data.append((col.column, row2_value))
        
        assert len(pattern2_data) == 3  # D, E, F
        assert pattern2_data[0] == ("D", False)           # D2
        assert pattern2_data[1][0] == "E"                 # E2 column
        assert isinstance(pattern2_data[1][1], formula)   # E2 is formula
        assert pattern2_data[2] == ("F", 200)            # F2
        
        print(f"Pattern 1 (rows 1:4 in cols A:C): {pattern1_data}")
        print(f"Pattern 2 (row 2 in cols D:F): {pattern2_data}")
    
    def test_data_type_filtering_across_ranges(self, rich_workbook):
        """Test filtering data types across column and row ranges."""
        sheet = rich_workbook["Sheet1"]
        
        # Collect all values from columns A:F, rows 1:6
        all_values = []
        value_locations = []
        
        for col in sheet[A:F]:
            for i, cell_value in enumerate(col[1:6], start=1):
                all_values.append(cell_value)
                value_locations.append(f"{col.column}{i}")
        
        # Categorize by type
        floats = [(loc, val) for loc, val in zip(value_locations, all_values) if isinstance(val, float)]
        ints = [(loc, val) for loc, val in zip(value_locations, all_values) if isinstance(val, int)]
        strings = [(loc, val) for loc, val in zip(value_locations, all_values) if isinstance(val, str) and not isinstance(val, formula)]
        bools = [(loc, val) for loc, val in zip(value_locations, all_values) if isinstance(val, bool)]
        formulas = [(loc, val) for loc, val in zip(value_locations, all_values) if isinstance(val, formula)]
        
        # Verify we have each type
        assert len(floats) > 0, f"Should have floats, found: {floats}"
        assert len(ints) > 0, f"Should have integers, found: {ints}"
        assert len(strings) > 0, f"Should have strings, found: {strings}"
        assert len(bools) > 0, f"Should have booleans, found: {bools}"
        assert len(formulas) > 0, f"Should have formulas, found: {formulas}"
        
        print(f"Data type distribution across A:F, rows 1:6:")
        print(f"  Floats ({len(floats)}): {floats}")
        print(f"  Integers ({len(ints)}): {ints}")
        print(f"  Strings ({len(strings)}): {strings}")  
        print(f"  Booleans ({len(bools)}): {bools}")
        print(f"  Formulas ({len(formulas)}): {formulas}")
    
    def test_sparse_data_ranges(self, rich_workbook):
        """Test range selection with sparse data (gaps between values)."""
        sheet = rich_workbook["Sheet1"]
        
        # Column A has data in rows 1-5 and 10-15, but gaps in 6-9
        col_a = sheet[A]
        
        # Test range that spans the gap
        range_with_gap = list(col_a[3:12])  # Should include gap
        
        # Should have values from rows 3-5 and 10-11
        assert 3.0 in range_with_gap      # A3
        assert 4.7 in range_with_gap      # A4
        assert 5.0 in range_with_gap      # A5
        assert "Row10" in range_with_gap  # A10
        assert "Row11" in range_with_gap  # A11
        
        # Should also have None values for the gap (rows 6-9)
        none_count = sum(1 for v in range_with_gap if v is None)
        assert none_count >= 4  # At least rows 6, 7, 8, 9
        
        print(f"Range with gap (3:12): {len(range_with_gap)} values, {none_count} None values")


class TestRowRangeEdgeCases:
    """Test edge cases in row range selection."""
    
    def test_row_range_beyond_data(self, rich_workbook):
        """Test row ranges that extend beyond available data."""
        sheet = rich_workbook["Sheet1"]
        
        col_b = sheet[B]
        
        # Range that goes beyond our data
        extended_range = list(col_b[3:20])  
        
        # Should include our known data plus None values
        assert "Gamma" in extended_range   # B3
        assert "Delta" in extended_range   # B4  
        assert "Epsilon" in extended_range # B5
        
        # Should stop naturally when no more data
        # Length should be reasonable (not trying to go to row 19)
        assert len(extended_range) >= 3
        
        print(f"Extended range (3:20) returned {len(extended_range)} values")
    
    def test_row_range_with_formula_dependencies(self, rich_workbook):
        """Test row ranges containing formulas that reference other cells."""
        sheet = rich_workbook["Sheet1"]
        
        # Column E has formulas that reference other columns
        col_e = sheet[E]
        
        formula_range = list(col_e[1:5])
        
        # Should have our formulas
        formulas_found = [v for v in formula_range if isinstance(v, formula)]
        assert len(formulas_found) >= 3  # E1, E2, E3, E4
        
        # Check specific formula content
        e1_formula = col_e[1]  # A1+C1
        assert isinstance(e1_formula, formula)
        assert str(e1_formula) == "A1+C1"
        
        e2_formula = col_e[2]  # LEN(B2) 
        assert isinstance(e2_formula, formula)
        assert str(e2_formula) == "LEN(B2)"
        
        print(f"Formula range (1:5) found {len(formulas_found)} formulas")
        for i, f in enumerate(formulas_found):
            print(f"  Formula {i+1}: {str(f)}")
    
    def test_empty_row_ranges(self, rich_workbook):
        """Test row ranges in areas with no data."""
        sheet = rich_workbook["Sheet1"]
        
        col_a = sheet[A]
        
        # Range in the gap area (rows 6-9 should be empty)
        gap_range = list(col_a[6:10])
        
        # Should be all None or empty
        non_none_values = [v for v in gap_range if v is not None]
        assert len(non_none_values) == 0, f"Gap should be empty, found: {non_none_values}"
        
        print(f"Gap range (6:10): {len(gap_range)} values, all None as expected")


class TestComplexRangeOperations:
    """Test complex combinations of range operations."""
    
    def test_matrix_like_access_pattern(self, rich_workbook):
        """Test matrix-like access patterns across ranges."""
        sheet = rich_workbook["Sheet1"]
        
        # Create a "matrix" view of data from B:E, rows 1:4
        matrix_data = []
        for col in sheet[B:E]:
            row_data = list(col[1:4])  # Rows 1, 2, 3
            matrix_data.append((col.column, row_data))
        
        assert len(matrix_data) == 4  # B, C, D, E columns
        
        # Check "row 1" across all columns (B1, C1, D1, E1)
        row1_across_cols = [data[1][0] for data in matrix_data]  # First row from each column
        
        expected_row1 = ["Alpha", 10, True, "A1+C1"]  # B1, C1, D1, E1 (E1 is formula)
        
        assert row1_across_cols[0] == "Alpha"           # B1
        assert row1_across_cols[1] == 10                # C1
        assert row1_across_cols[2] is True              # D1
        assert isinstance(row1_across_cols[3], formula) # E1
        assert str(row1_across_cols[3]) == "A1+C1"
        
        print(f"Matrix view (B:E × 1:4):")
        for col_letter, row_data in matrix_data:
            print(f"  Column {col_letter}: {[type(v).__name__ for v in row_data]}")
    
    def test_range_operations_preserve_types(self, rich_workbook):
        """Test that range operations preserve original data types correctly."""
        sheet = rich_workbook["Sheet1"]
        
        # Test that slicing preserves types exactly
        original_values = []
        sliced_values = []
        
        # Get individual cell values
        for row in range(1, 4):
            original_values.append((f"C{row}", sheet[C][row]))
        
        # Get same values via slicing
        col_c = sheet[C]
        for i, value in enumerate(col_c[1:4], start=1):
            sliced_values.append((f"C{i}", value))
        
        # Should be identical
        assert original_values == sliced_values
        
        # Types should be preserved exactly
        for (orig_ref, orig_val), (slice_ref, slice_val) in zip(original_values, sliced_values):
            assert type(orig_val) == type(slice_val)
            assert orig_val == slice_val
            
        print(f"Type preservation test passed:")
        for ref, val in original_values:
            print(f"  {ref}: {repr(val)} ({type(val).__name__})")


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])