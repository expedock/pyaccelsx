# Pyaccelsx

Simple Python-bindings for rust_xlsxwriter for writing large worksheets, fast.

## Examples

### Writing a Simple Workbook

```python
from pyaccelsx import ExcelWorkbook

# Create a new workbook and add a worksheet
workbook = ExcelWorkbook()
workbook.add_worksheet("Sheet 1")

# Write some data to the worksheet
workbook.write(0, 0, "Hello")
workbook.write(0, 1, "World!")

# Save the workbook
workbook.save("example.xlsx")
```

### Writing with Formatting

```python
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
workbook.write(0, 0, "Hello", format_option=bold_format)
workbook.write(0, 1, 44123.456, format_option=numeric_format)
workbook.write(0, 2, "Right", format_option=right_aligned_format)
workbook.write(0, 3, "Color", format_option=color_format)
workbook.write_and_merge_range(1, 0, 1, 3, "Merge", format_option=merge_format)
workbook.write_and_merge_range(2, 0, 2, 3, 123456, format_option=merge_format)
workbook.write(3, 1, "border", format_option=border_format)

# Save the workbook
workbook.save("example.xlsx")
```

## Performance

We evaluate `pyaccelsx` performance on writing **4,000 rows**, **50 columns**, and **1 sheet**, and used [hyperfine](https://lib.rs/crates/hyperfine) to compare the performance with `rust_xlsxwriter`, `XlsxWriter`, and `openpyxl`.

As of testing, we used the following versions:

- `rust_xlsxwriter == 0.77.0`
- `XlsxWriter == 3.1.9`
- `openpyxl == 3.0.9`

### `XlsxWriter`, `openpyxl`, and `pyaccelsx`

```bash
$ hyperfine 'python3 xlsxwriter_test.py' 'python3 openpyxl_test.py' 'python3 pyaccelsx_test.py' --warmup 5 --runs 20

Benchmark 1: python3 xlsxwriter_test.py
  Time (mean ± σ):     727.7 ms ±  19.0 ms    [User: 706.0 ms, System: 20.4 ms]
  Range (min … max):   704.6 ms … 781.3 ms    20 runs
 
Benchmark 2: python3 openpyxl_test.py
  Time (mean ± σ):      1.860 s ±  0.061 s    [User: 2.133 s, System: 1.075 s]
  Range (min … max):    1.765 s …  2.003 s    20 runs
 
Benchmark 3: python3 pyaccelsx_test.py
  Time (mean ± σ):     330.3 ms ±  12.2 ms    [User: 305.2 ms, System: 19.1 ms]
  Range (min … max):   314.1 ms … 372.6 ms    20 runs
 
Summary
  'python3 pyaccelsx_test.py' ran
    2.20 ± 0.10 times faster than 'python3 xlsxwriter_test.py'
    5.63 ± 0.28 times faster than 'python3 openpyxl_test.py'
```

### `rust_xlsxwriter`, `XlsxWriter`, `openpyxl`, and `pyaccelsx`

```bash
$ hyperfine './target/release/rust_test' 'python3 xlsxwriter_test.py' 'python3 openpyxl_test.py' 'python3 pyaccelsx_test.py' --warmup 5 --runs 20

Benchmark 1: ./target/release/rust_test
  Time (mean ± σ):     166.1 ms ±   3.5 ms    [User: 148.6 ms, System: 10.5 ms]
  Range (min … max):   160.5 ms … 176.8 ms    20 runs
 
Benchmark 2: python3 xlsxwriter_test.py
  Time (mean ± σ):     765.4 ms ±  37.5 ms    [User: 734.7 ms, System: 29.4 ms]
  Range (min … max):   713.0 ms … 862.0 ms    20 runs
 
Benchmark 3: python3 openpyxl_test.py
  Time (mean ± σ):      1.846 s ±  0.064 s    [User: 2.116 s, System: 1.071 s]
  Range (min … max):    1.747 s …  1.979 s    20 runs
 
Benchmark 4: python3 pyaccelsx_test.py
  Time (mean ± σ):     324.8 ms ±  10.9 ms    [User: 302.1 ms, System: 16.7 ms]
  Range (min … max):   312.8 ms … 360.9 ms    20 runs
 
Summary
  './target/release/rust_test' ran
    1.95 ± 0.08 times faster than 'python3 pyaccelsx_test.py'
    4.61 ± 0.25 times faster than 'python3 xlsxwriter_test.py'
   11.11 ± 0.45 times faster than 'python3 openpyxl_test.py'
```

## Contributing

This library uses [pre-commit](https://pre-commit.com/). Please ensure it's installed before submitting PRs.
