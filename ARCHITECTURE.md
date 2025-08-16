# libxlsx Architecture

## Overview

libxlsx is a surgical XLSX editor designed for lossless, minimal modifications to existing Excel workbooks. Rather than attempting to recreate Excel's full feature set, we surgically edit only what we must while preserving all other file contents byte-for-byte.

## Design Philosophy

### Surgical Editing Principle
- **Goal**: Make targeted edits to specific cells while leaving everything else untouched
- **Fidelity**: Only re-serialize XML parts we actually modify
- **Compatibility**: Preserve complex Excel features (charts, pivots, macros, formatting) exactly as created

### Why This Approach?
Excel workbooks contain incredibly complex internal structures. Attempting to parse, understand, and reimplement all of Excel's features would be rebuilding Excel itself—an enormous, never-ending project. Instead, we:

1. Load the entire XLSX file into memory
2. Parse only what we need for targeted operations
3. Make surgical modifications to the in-memory representation
4. Save only the changed parts back to disk

This ensures the original workbook remains fully functional when reopened in Excel.

## Technical Architecture

### Core Technologies
- **Backend**: `zipfile` + `lxml.etree`
  - `zipfile`: Handle XLSX as ZIP archive format
  - `lxml.etree`: Fast, precise XML parsing and modification
- **Format**: XLSX only (Open XML)
- **Compatibility**: Python 3.11+

### Data Flow

```
┌─────────────┐    ┌──────────────┐    ┌─────────────┐
│ Load XLSX   │ ──▶│ Build Cell   │ ──▶│ Provide API │
│ into Memory │    │ Index        │    │ Access      │
└─────────────┘    └──────────────┘    └─────────────┘
                                              │
┌─────────────┐    ┌──────────────┐    ┌─────────────┐
│ Save        │ ◀──│ Modify       │ ◀──│ Edit Cell   │
│ Changes     │    │ XML in       │    │ Values      │
└─────────────┘    │ Memory       │    └─────────────┘
                   └──────────────┘
```

1. **Load**: Parse XLSX structure and build complete cell coordinate index
2. **Access**: Use string indexing (`sheet[A][1]`) to read values
3. **Edit**: Modify XML nodes in-memory, re-parse affected parts
4. **Save**: Write only modified ZIP members back to disk

### API Design

```python
from libxlsx import load_workbook, formula
from libxlsx.const import *

# Load existing workbook
workbook = load_workbook("existing.xlsx")
sheet = workbook["Sheet1"]

# Read/write individual cells
value = sheet[A][1]              # Read cell A1
sheet[A][1] = "Hello World"      # Write cell A1
sheet[B][2] = formula("=A1*2")   # Write formula

# Iterate over columns and ranges
for cell in sheet[A]:            # All values in column A
    print(cell)

for cell in sheet[A][1:5]:       # Cells A1:A5
    print(cell)

for col in sheet[A:C]:           # Columns A through C
    for cell in col:
        print(cell)

# Save with minimal changes
workbook.save("modified.xlsx")
```

## Implementation Scope

### What We Support
- **Cell value reading/writing** (strings, numbers, booleans)
- **Formula editing** (as string values)
- **Column/row iteration and slicing**
- **Existing workbook modification**

### What We Don't Support (Yet)
- **Creating new workbooks** from scratch (`NotImplementedError`)
- **Structural changes** (adding sheets, rows, columns)
- **Formatting modifications** (colors, fonts, borders)
- **Advanced features** (charts, pivots, macros modification)

### What Excel Handles
- **Formula recalculation** (happens when file reopened in Excel)
- **Dependency tracking** (Excel manages the computational graph)
- **Data validation** (Excel enforces rules on reopened files)

## Key Design Decisions

### Memory Strategy
- **Load everything upfront**: RAM is cheap, simplicity is valuable
- **Complete cell indexing**: Resolve all cell coordinates on initial load
- **In-memory editing**: Modify ZIP contents in memory before saving

### Error Handling
- **Fail fast**: If Excel created it, we should be able to read it
- **Parse errors become bug reports**: Any failure to read valid XLSX is a library bug to fix

### Future Optimizations
- **Shared strings table**: Handle string deduplication properly
- **Selective re-parsing**: Only re-parse modified XML sections
- **Dirty state tracking**: Track which ZIP members need saving
- **Performance profiling**: Optimize bottlenecks as they emerge

## File Structure

```
src/libxlsx/
├── __init__.py          # Public API exports
├── workbook.py          # Workbook class
├── sheet.py             # Sheet class
├── column.py            # Column class
├── const.py             # A-Z column constants
└── ARCHITECTURE.md      # This file
```

The clean separation allows each class to focus on its specific responsibilities while maintaining the surgical editing philosophy throughout the codebase.