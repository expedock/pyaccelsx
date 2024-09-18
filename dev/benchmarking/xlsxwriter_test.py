import sys
import xlsxwriter

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

workbook = xlsxwriter.Workbook('py_xlsxwriter_perf_test.xlsx')

for r in range(sheets):
    worksheet = workbook.add_worksheet()
    for row in range(0, row_max):
        for col in range(0, col_max):
            if col % 2:
                worksheet.write_string(row, col, "Foo")
            else:
                worksheet.write_number(row, col, 12345)

workbook.close()