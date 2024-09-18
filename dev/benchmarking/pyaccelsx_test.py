import sys
from pyaccelsx import ExcelWorkbook

# Default to 1 sheet with 1000 rows x 50 cols
row_max = 4000
col_max = 50
sheets = 1

for i, arg in enumerate(sys.argv):
    if i == 1:
        row_max = int(sys.argv[1])
    elif i == 2:
        col_max = int(sys.argv[2])
    elif i == 3:
        sheets = int(sys.argv[3])

workbook = ExcelWorkbook()

for r in range(sheets):
    workbook.add_worksheet()
    for row in range(0, row_max):
        for col in range(0, col_max):
            if col % 2:
                workbook.write(row, col, "Foo")
            else:
                workbook.write(row, col, 12345)

workbook.save("pyaccelsx_perf_test.xlsx")