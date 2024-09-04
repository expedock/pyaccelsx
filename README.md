# Pyaccelsx

Simple Python-bindings for rust_xlsxwriter for writing large worksheets, fast.

### Simple Example
```
from pyaccelsx import ExcelWorkbook

# Create a new workbook and add a worksheet
workbook = ExcelWorkbook()
workbook.add_worksheet("Sheet 1")

# Write some data to the worksheet
workbook.write_string(0, 0, "Hello")
workbook.write_string(0, 1, "World!")

# Save the workbook
workbook.save("example.xlsx")
```

### Writing with Formatting
```
from pyaccelsx import ExcelWorkbook, ExcelFormat

# Create a new workbook and add a worksheet
workbook = ExcelWorkbook()
workbook.add_worksheet("Sheet 1")

# Write some formats to be applied to cells
bold_format = ExcelFormat(
    bold=True,
)
numeric_format = ExcelFormat(
    num_format="#,##0.00",
)
right_aligned_format = ExcelFormat(
    align="right",
)
border_format = ExcelFormat(
    border_right=True,
    border_bottom=True,
)
color_format = ExcelFormat(
    font_color="FF0000",
)
merge_format = ExcelFormat(
    border=True,
    bold=True,
    align="center",
)

# Write some data to the worksheet
workbook.write_string(0, 0, "Hello", bold_format)
workbook.write_number(0, 1, 44123.456, numeric_format)
workbook.write_string(0, 2, "Right", right_aligned_format)
workbook.write_string(0, 3, "Color", color_format)
workbook.write_string_and_merge_range(1, 0, 1, 3, "Merge", merge_format)
workbook.write_number_and_merge_range(2, 0, 2, 3, 123456, merge_format)
workbook.write_string(3, 1, "border", border_format)

# Save the workbook
workbook.save("example.xlsx")
```

## Contributing

This library uses [pre-commit](https://pre-commit.com/). Please ensure it's installed before submitting PRs.
