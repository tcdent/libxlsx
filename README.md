

```python

from lxls import load_workbook, formula
from lxls.const import *

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