
# libxlsx - Surgical XLSX Editor

**libxlsx** is a Python library for making surgical edits to existing Excel (.xlsx) files while preserving all original formatting, charts, pivot tables, macros, and advanced features.

## Why libxlsx?

Traditional Excel libraries try to parse and recreate entire workbooks, often losing complex formatting and advanced features in the process. This approach requires implementing Excel's vast feature set—essentially rebuilding Excel itself.

**libxlsx takes a different approach**: it loads the entire XLSX file into memory and makes only targeted, surgical modifications to specific cells while leaving everything else byte-for-byte identical. This ensures complete fidelity with the original workbook.

## Quick Start

```bash
pip install libxlsx
```

### Basic Usage

```python
from libxlsx import load_workbook, formula
from libxlsx.const import *

# Load existing workbook (preserves everything)
workbook = load_workbook("report.xlsx")
sheet = workbook["Sheet1"]

# Read values
revenue = sheet[A][1]  # Get cell A1
print(f"Current revenue: {revenue}")

# Write values (only these cells are modified)
sheet[A][1] = 125000.0        # Update revenue
sheet[B][1] = "Updated"        # Add status
sheet[C][1] = formula("=A1*1.1")  # Add 10% projection

# Save with minimal changes
workbook.save("report_updated.xlsx")
```

### Real-World Example

```python
from libxlsx import load_workbook, formula
from libxlsx.const import *

# Update a financial model without breaking formulas/charts
model = load_workbook("financial_model.xlsx")
assumptions = model["Assumptions"]

# Update key assumptions (formulas elsewhere will recalculate)
assumptions[B][5] = 0.15    # Growth rate
assumptions[B][6] = 0.25    # Tax rate
assumptions[B][7] = 1000000 # Base revenue

# Add validation formula
assumptions[B][10] = formula("=IF(B5>0.5,\"Check growth rate\",\"OK\")")

# All charts, pivot tables, and complex formulas remain intact
model.save("financial_model_updated.xlsx")
```

## API Design

The following code examples demonstrate the intended API design and functionality:

```python

from libxlsx import load_workbook, formula
from libxlsx.const import *

workbook = load_workbook("Workbook.xlsx")

sheet = workbook["Sheet1"]


# column + row access
sheet[A][1] = 123.45
sheet[A][2] = "Hello world"
sheet[B][1] = formula("=SUM(B2:B8)")


x: str = sheet[A][2]

for row in sheet[A]:
    print(row)
# -> 123.45
# -> "Hello world"

len(sheet[A])
# -> 2


for row in sheet[A][1:2]:
    print(row)
# -> 123.45
# -> "Hello world"


for col in sheet[A:B]:
    for cell in col:
        print(cell)
# -> 123.45
# -> "Hello world"
# -> =SUM(B2:B8)


workbook.save("Workbook (modified).xlsx")
```

