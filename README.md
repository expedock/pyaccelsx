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

We evaluate `pyaccelsx` performance on writing **4,000 rows**, **50 columns**, and **1 sheet**, and used [hyperfine](https://lib.rs/crates/hyperfine) to compare the performance with `rust_xlsxwriter` and `xlsxwriter`.

### `xlsxwriter` and `pyaccelsx`

```bash
$ hyperfine 'python3 py_xlsxwriter_test.py' 'python3 pyaccelsx_test.py' --warmup 5 --runs 20

Benchmark 1: python3 py_xlsxwriter_test.py
  Time (mean ± σ):     736.0 ms ±  25.1 ms    [User: 708.7 ms, System: 24.9 ms]
  Range (min … max):   704.4 ms … 800.0 ms    20 runs
 
Benchmark 2: python3 pyaccelsx_test.py
  Time (mean ± σ):     338.9 ms ±  13.6 ms    [User: 313.8 ms, System: 19.7 ms]
  Range (min … max):   324.5 ms … 374.5 ms    20 runs
 
Summary
  'python3 pyaccelsx_test.py' ran
    2.17 ± 0.11 times faster than 'python3 py_xlsxwriter_test.py'
```

### `rust_xlsxwriter`, `xlsxwriter`, and `pyaccelsx`

```bash
$ hyperfine './target/release/rust_test' 'python3 py_xlsxwriter_test.py' 'python3 pyaccelsx_test.py' --warmup 5 --runs 20

Benchmark 1: ./target/release/rust_test
  Time (mean ± σ):     166.8 ms ±   4.2 ms    [User: 149.0 ms, System: 10.2 ms]
  Range (min … max):   162.7 ms … 182.0 ms    20 runs
 
Benchmark 2: python3 py_xlsxwriter_test.py
  Time (mean ± σ):     742.2 ms ±  24.6 ms    [User: 711.8 ms, System: 30.0 ms]
  Range (min … max):   715.5 ms … 819.2 ms    20 runs
 
Benchmark 3: python3 pyaccelsx_test.py
  Time (mean ± σ):     330.1 ms ±   7.3 ms    [User: 307.3 ms, System: 15.8 ms]
  Range (min … max):   320.5 ms … 343.0 ms    20 runs
 
Summary
  './target/release/rust_test' ran
    1.98 ± 0.07 times faster than 'python3 pyaccelsx_test.py'
    4.45 ± 0.19 times faster than 'python3 py_xlsxwriter_test.py'
```

## Contributing

This library uses [pre-commit](https://pre-commit.com/). Please ensure it's installed before submitting PRs.
