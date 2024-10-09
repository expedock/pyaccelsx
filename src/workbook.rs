use super::format::{self, ExcelFormat};
use pyo3::prelude::*;
use rust_xlsxwriter::{ColNum, Format, RowNum, Workbook};

use crate::util::ValueType;
use crate::writer;

#[pyclass]
/// The `ExcelWorkbook` struct represents an Excel workbook.
/// This contains the workbook object and the active worksheet index.
/// Worksheet methods are directly implemented under this class,
/// as they are mutable references in which the ownership cannot be transferred.
pub struct ExcelWorkbook {
    workbook: Workbook,
    active_worksheet_index: usize,
}

#[pymethods]
impl ExcelWorkbook {
    #[new]
    /// Create a new workbook.
    /// ## Examples
    /// The following example demonstrates creating a simple workbook, with one unused worksheet.
    /// ```
    /// from pyaccelsx import ExcelWorkbook
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///     workbook.add_worksheet()
    ///     workbook.save("example.xlsx")
    /// ```
    #[pyo3(signature = (use_zip64=false))]
    pub fn new(use_zip64: bool) -> ExcelWorkbook {
        let mut workbook = Workbook::new();
        if use_zip64 {
            workbook.use_zip_large_file(true);
        }
        ExcelWorkbook {
            workbook,
            active_worksheet_index: 0,
        }
    }

    /// Add a new worksheet to the workbook and update the active worksheet index.
    #[pyo3(signature = (name=None))]
    /// Adds a new worksheet to the workbook with the given sheet name.
    /// If no name is given, the standard names (`Sheet1`, `Sheet2`, etc.) will be used.
    ///
    /// Pyaccelsx used `active_worksheet_index` to keep track of the active worksheet.
    /// Adding a new worksheet into the workbook will automatically increment `active_worksheet_index`.
    ///
    /// ## Parameters
    /// - `name`: The name of the new worksheet _(optional)_
    ///
    /// ## Examples
    /// The following example demonstrates adding worksheets to a workbook.
    /// ```
    /// from pyaccelsx import ExcelWorkbook
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///
    ///     workbook.add_worksheet()    // Sheet1
    ///     workbook.add_worksheet("My Sheet")
    ///     workbook.add_worksheet()    // Sheet3
    ///
    ///     // This is written in Sheet3
    ///     workbook..write(0, 0, "Hello")
    ///     
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn add_worksheet(&mut self, name: Option<String>) -> PyResult<()> {
        if name.is_none() {
            self.workbook.add_worksheet();
        } else {
            self.workbook
                .add_worksheet()
                .set_name(name.unwrap())
                .unwrap();
        }
        self.active_worksheet_index = self.workbook.worksheets().len() - 1;
        Ok(())
    }

    /// Set the active worksheet index.
    ///
    /// ## Parameters
    /// - `index`: The index of the worksheet
    ///
    /// ## Examples
    /// The following example demonstrates setting the active worksheet.
    /// ```
    /// from pyaccelsx import ExcelWorkbook
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///     workbook.add_worksheet("Sheet 1")
    ///     workbook.add_worksheet("Sheet 2")
    ///     workbook.set_active_worksheet(0)
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn set_active_worksheet(&mut self, index: usize) -> PyResult<()> {
        self.active_worksheet_index = index;
        Ok(())
    }

    /// Save the workbook into the specified path.
    ///
    /// ## Parameters
    /// - `path`: The path to save the workbook
    ///
    /// ## Examples
    /// The following example demonstrates saving a workbook.
    /// ```
    /// from pyaccelsx import ExcelWorkbook
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///     workbook.add_worksheet()
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn save(&mut self, path: String) -> PyResult<()> {
        self.workbook.save(path).unwrap();
        Ok(())
    }

    #[pyo3(signature = (row, column, value=None, override_true_value=None, override_false_value=None, override_value=None, format_option=None))]
    /// Worksheet handler for writing a value to a cell.
    ///
    /// ## Parameters
    /// - `row`: The row number of the cell
    /// - `column`: The column number of the cell
    /// - `value`: The value to write
    /// - `override_true_value`: The value to write if the value is `True`
    /// - `override_false_value`: The value to write if the value is `False`
    /// - `override_value`: The value to write if the value is `None`
    /// - `format_option`: The format to apply to the cell
    ///
    /// ## Examples
    /// The following example demonstrates writing a value to a cell in a workbook.
    /// ```
    /// from pyaccelsx import ExcelWorkbook, ExcelFormat
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///     workbook.add_worksheet()
    ///     
    ///     // Write a string
    ///     workbook.write(0, 0, "Hello")
    ///     // Write a boolean
    ///     workbook.write(1, 0, True, "Yes", "No")
    ///     // Write an integer
    ///     workbook.write(2, 0, 42, format_option=ExcelFormat(num_format="#,##0"))
    ///     // Write a float
    ///     workbook.write(3, 0, 3.14, format_option=ExcelFormat(num_format="#,##0.00"))
    ///     // Write None
    ///     workbook.write(4, 0, None, override_value="Empty")
    ///     
    ///     workbook.save("example.xlsx")
    pub fn write(
        &mut self,
        row: RowNum,
        column: ColNum,
        value: Option<ValueType>,
        override_true_value: Option<String>,
        override_false_value: Option<String>,
        override_value: Option<String>,
        format_option: Option<ExcelFormat>,
    ) -> PyResult<()> {
        let worksheet = self
            .workbook
            .worksheet_from_index(self.active_worksheet_index)
            .unwrap();

        if let Some(value) = value {
            match value {
                ValueType::String(value) => {
                    writer::write_string(worksheet, row, column, value, format_option)
                }
                ValueType::Bool(value) => writer::write_boolean(
                    worksheet,
                    row,
                    column,
                    value,
                    override_true_value,
                    override_false_value,
                    format_option,
                ),
                ValueType::Int(value) => {
                    writer::write_number(worksheet, row, column, value, format_option)
                }
                ValueType::Float(value) => {
                    writer::write_number(worksheet, row, column, value, format_option)
                }
            }
            .unwrap();
        } else {
            writer::write_null(worksheet, row, column, override_value, format_option).unwrap();
        }

        Ok(())
    }

    #[pyo3(signature = (row, column, format_option=None))]
    /// Worksheet handler for writing a "blank" cell.
    /// This function will only perform write if `format_option` is specified.
    /// If there is no format option specified, the corresponding cell will be an "empty" cell.
    ///
    /// See this [documentation](https://docs.rs/rust_xlsxwriter/0.75.0/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.write_blank) for difference between "blank cell" and "empty cell".
    ///
    /// ## Parameters
    /// - `row`: The row index of the cell
    /// - `column`: The column index of the cell
    /// - `format_option`: The format of the cell _(optional)_
    ///
    /// ## Examples
    /// The following example demonstrates writing a "blank" cell to a worksheet.
    /// ```
    /// from pyaccelsx import ExcelWorkbook, ExcelFormat
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///     workbook.add_worksheet()
    ///     
    ///     format_option = ExcelFormat(border=True)
    ///     workbook.write_blank(0, 0, format_option)
    ///     // This will not perform any write
    ///     workbook.write_blank(0, 1)
    ///
    ///     workbook.save("example.xlsx")
    /// ```
    ///
    pub fn write_blank(
        &mut self,
        row: RowNum,
        column: ColNum,
        format_option: Option<ExcelFormat>,
    ) -> PyResult<()> {
        if let Some(format_option) = format_option {
            let worksheet = self
                .workbook
                .worksheet_from_index(self.active_worksheet_index)
                .unwrap();
            let format = format::create_format(format_option);
            worksheet.write_blank(row, column, &format).unwrap();
        }
        Ok(())
    }

    #[pyo3(signature = (start_row, start_column, end_row, end_column, format_option=None))]
    /// Worksheet handler for merging a range of cells. This will not do any
    /// writing to the cell values. To write values, use `write_and_merge_range`.
    ///
    /// ## Parameters
    /// - `start_row`: The start row index of the range
    /// - `start_column`: The start column index of the range
    /// - `end_row`: The end row index of the range
    /// - `end_column`: The end column index of the range
    /// - `format_option`: The format of the cell _(optional)_
    ///
    /// ## Examples
    /// The following example demonstrates merging cells in a worksheet.
    /// ```
    /// from pyaccelsx import ExcelWorkbook, ExcelFormat
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///     workbook.add_worksheet()
    ///     
    ///     format_option = ExcelFormat(align="center", border=True)
    ///     workbook.merge_range(0, 0, 0, 2, format_option)
    ///
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn merge_range(
        &mut self,
        start_row: RowNum,
        start_column: ColNum,
        end_row: RowNum,
        end_column: ColNum,
        format_option: Option<ExcelFormat>,
    ) -> PyResult<()> {
        let worksheet = self
            .workbook
            .worksheet_from_index(self.active_worksheet_index)
            .unwrap();
        if format_option.is_none() {
            worksheet
                .merge_range(
                    start_row,
                    start_column,
                    end_row,
                    end_column,
                    "",
                    &Format::new(),
                )
                .unwrap();
        } else {
            let format = format::create_format(format_option.unwrap());
            worksheet
                .merge_range(start_row, start_column, end_row, end_column, "", &format)
                .unwrap();
        }
        Ok(())
    }

    #[pyo3(signature = (start_row, start_column, end_row, end_column, value=None, override_true_value=None, override_false_value=None, override_value=None, format_option=None))]
    /// Worksheet handler for merging a range of cells and writing string value into the merged cells.
    ///
    /// ## Parameters
    /// - `start_row`: The start row index of the range
    /// - `start_column`: The start column index of the range
    /// - `end_row`: The end row index of the range
    /// - `end_column`: The end column index of the range
    /// - `value`: The string value to write
    /// - `override_true_value`: The string value to write if the cell value is `True` _(optional)_
    /// - `override_false_value`: The string value to write if the cell value is `False` _(optional)_
    /// - `override_value`: The string value to write if the cell value is not `True` or `False` _(optional)_
    /// - `format_option`: The format of the cell _(optional)_
    ///
    /// ## Examples
    /// The following example demonstrates merging cells and writing string value into the merged cells in a worksheet.
    /// ```
    /// from pyaccelsx import ExcelWorkbook, ExcelFormat
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///     workbook.add_worksheet()
    ///     
    ///     format_option = ExcelFormat(align="center", border=True)
    ///     workbook.write_and_merge_range(0, 0, 0, 2, "Hello World!", format_option)
    ///
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn write_and_merge_range(
        &mut self,
        start_row: RowNum,
        start_column: ColNum,
        end_row: RowNum,
        end_column: ColNum,
        value: Option<ValueType>,
        override_true_value: Option<String>,
        override_false_value: Option<String>,
        override_value: Option<String>,
        format_option: Option<ExcelFormat>,
    ) -> PyResult<()> {
        if let Some(value) = value {
            // Prevent using moved value
            let cloned_format_option = format_option.clone();
            self.merge_range(start_row, start_column, end_row, end_column, format_option)
                .unwrap();
            if cloned_format_option.is_none() {
                self.write(
                    start_row,
                    start_column,
                    Some(value),
                    override_true_value,
                    override_false_value,
                    override_value,
                    cloned_format_option,
                )
                .unwrap();
            } else {
                self.write(
                    start_row,
                    start_column,
                    Some(value),
                    override_true_value,
                    override_false_value,
                    override_value,
                    cloned_format_option,
                )
                .unwrap();
            }
        } else {
            self.merge_range(start_row, start_column, end_row, end_column, format_option)
                .unwrap();
        }
        Ok(())
    }

    /// Worksheet handler for setting column width.
    ///
    /// ## Parameters
    /// - `column`: The column index of the cell
    /// - `width`: The width of the column
    ///
    /// ## Examples
    /// The following example demonstrates setting column width in a worksheet.
    /// ```
    /// from pyaccelsx import ExcelWorkbook
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///     workbook.add_worksheet()
    ///
    ///     workbook.write(0, 0, "Hello World!")
    ///     workbook.set_column_width(0, 20)
    ///
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn set_column_width(&mut self, column: ColNum, width: f64) -> PyResult<()> {
        let worksheet = self
            .workbook
            .worksheet_from_index(self.active_worksheet_index)
            .unwrap();
        worksheet.set_column_width(column, width).unwrap();
        Ok(())
    }

    /// Worksheet handler for freezing panes.
    ///
    /// ## Parameters
    /// - `row`: The row index of the cell which will be frozen
    /// - `column`: The column index of the cell which will be frozen
    ///
    /// ## Examples
    /// The following example demonstrates freezing panes in a worksheet.
    /// ```
    /// from pyaccelsx import ExcelWorkbook
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///     workbook.add_worksheet()
    ///
    ///     workbook.write(0, 0, "Hello World!")
    ///     // This freezes the first row and first column
    ///     workbook.freeze_panes(0, 0)
    ///
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn freeze_panes(&mut self, row: RowNum, column: ColNum) -> PyResult<()> {
        let worksheet = self
            .workbook
            .worksheet_from_index(self.active_worksheet_index)
            .unwrap();
        worksheet.set_freeze_panes(row, column).unwrap();
        Ok(())
    }
}

impl Default for ExcelWorkbook {
    fn default() -> Self {
        Self::new(false)
    }
}
