use super::format::{self, ExcelFormat};
use pyo3::prelude::*;
use rust_xlsxwriter::{Format, Workbook};

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
    pub fn new() -> ExcelWorkbook {
        let workbook = Workbook::new();
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
    ///     workbook..write_string(0, 0, "Hello")
    ///     
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn add_worksheet(&mut self, name: Option<&str>) -> PyResult<()> {
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
    pub fn save(&mut self, path: &str) -> PyResult<()> {
        self.workbook.save(path).unwrap();
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
        row: u32,
        column: u16,
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

    #[pyo3(signature = (row, column, override_value=None, format_option=None))]
    /// Worksheet handler for writing `None` values.
    ///
    /// ## Parameters
    /// - `row`: The row index of the cell
    /// - `column`: The column index of the cell
    /// - `override_value`: The string value to write if cell value is None _(optional)_
    /// - `format_option`: The format of the cell _(optional)_
    ///
    /// ## Examples
    /// The following example demonstrates writing a `None` value to a worksheet.
    /// ```
    /// from pyaccelsx import ExcelWorkbook, ExcelFormat
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///     workbook.add_worksheet()
    ///
    ///     // This will write "N/A" for `None` values
    ///     workbook.write_null(0, 0, "N/A")
    ///     // This will not perform any write    
    ///     workbook.write_null(0, 1)
    ///     // This will perform `write_blank`
    ///     format_option = ExcelFormat(align="right")
    ///     workbook.write_null(row=0, column=2, format_option=format_option)
    ///
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn write_null(
        &mut self,
        row: u32,
        column: u16,
        override_value: Option<&str>,
        format_option: Option<ExcelFormat>,
    ) -> PyResult<()> {
        let worksheet = self
            .workbook
            .worksheet_from_index(self.active_worksheet_index)
            .unwrap();
        if format_option.is_none() {
            if let Some(override_value) = override_value {
                worksheet.write_string(row, column, override_value).unwrap();
            }
        } else {
            let format = format::create_format(format_option.unwrap());
            match override_value {
                Some(override_value) => worksheet
                    .write_string_with_format(row, column, override_value, &format)
                    .unwrap(),
                None => worksheet.write_blank(row, column, &format).unwrap(),
            };
        }
        Ok(())
    }

    #[pyo3(signature = (row, column, value, format_option=None))]
    /// Worksheet handler for writing string values.
    ///
    /// ## Parameters
    /// - `row`: The row index of the cell
    /// - `column`: The column index of the cell
    /// - `value`: The string value to write
    /// - `format_option`: The format of the cell _(optional)_
    ///
    /// ## Examples
    /// The following example demonstrates writing a string to a worksheet.
    /// ```
    /// from pyaccelsx import ExcelWorkbook, ExcelFormat
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///     workbook.add_worksheet()
    ///
    ///     workbook.write_string(0, 0, "Hello")
    ///     format_option = ExcelFormat(bold=True)
    ///     workbook.write_string(0, 1, "World!", format_option)
    ///
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn write_string(
        &mut self,
        row: u32,
        column: u16,
        value: &str,
        format_option: Option<ExcelFormat>,
    ) -> PyResult<()> {
        let worksheet = self
            .workbook
            .worksheet_from_index(self.active_worksheet_index)
            .unwrap();
        if format_option.is_none() {
            worksheet.write_string(row, column, value).unwrap();
        } else {
            let format = format::create_format(format_option.unwrap());
            worksheet
                .write_string_with_format(row, column, value, &format)
                .unwrap();
        }
        Ok(())
    }

    #[pyo3(signature = (row, column, value, format_option=None))]
    /// Worksheet handler for writing numeric values. By default, values
    /// are passed as a float (`f64`) type. To specify whether to write intenger
    /// or float, use the `format_option` and set the `num_format` field.
    ///
    /// ## Parameters
    /// - `row`: The row index of the cell
    /// - `column`: The column index of the cell
    /// - `value`: The numeric value to write
    /// - `format_option`: The format of the cell _(optional)_
    ///
    /// ## Examples
    /// The following example demonstrates writing numeric values to a worksheet.
    /// ```
    /// from pyaccelsx import ExcelWorkbook, ExcelFormat
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///     workbook.add_worksheet()
    ///     
    ///     // This will be written to cell as float (123445.21)
    ///     format_option = ExcelFormat(num_format="#,##0.00")
    ///     workbook.write_number(0, 0, 123445.21, format_option)
    ///     // This will be written to cell as integer (123,456)
    ///     format_option = ExcelFormat(num_format="#,##0")
    ///     workbook.write_number(0, 1, 123456, format_option)
    ///
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn write_number(
        &mut self,
        row: u32,
        column: u16,
        value: f64,
        format_option: Option<ExcelFormat>,
    ) -> PyResult<()> {
        let worksheet = self
            .workbook
            .worksheet_from_index(self.active_worksheet_index)
            .unwrap();
        if format_option.is_none() {
            worksheet.write_number(row, column, value).unwrap();
        } else {
            let format = format::create_format(format_option.unwrap());
            worksheet
                .write_number_with_format(row, column, value, &format)
                .unwrap();
        }
        Ok(())
    }

    #[pyo3(signature = (row, column, value, override_true=None, override_false=None, format_option=None))]
    /// Worksheet handler for writing boolean values. By default, values
    /// are written as `True` or `False`. To specify an override string value
    /// to replace the boolean values, use the `override_value` field.
    ///
    /// ## Parameters
    /// - `row`: The row index of the cell
    /// - `column`: The column index of the cell
    /// - `value`: The boolean value to write
    /// - `override_true`: The string value to override `True` value _(optional)_
    /// - `override_false`: The string value to override `False` value _(optional)_
    /// - `format_option`: The format of the cell _(optional)_
    ///
    /// ## Examples
    /// The following example demonstrates writing boolean values to a worksheet.
    /// ```
    /// from pyaccelsx import ExcelWorkbook
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///     workbook.add_worksheet()
    ///     
    ///     // This will be written to cell as "TRUE"
    ///     workbook.write_boolean(0, 0, True)
    ///     // This will be written to cell as "FALSE"
    ///     workbook.write_boolean(0, 1, False)
    ///     // This will be written to cell as "No"
    ///     workbook.write_boolean(0, 2, False, "Yes", "No")
    ///
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn write_boolean(
        &mut self,
        row: u32,
        column: u16,
        value: bool,
        override_true: Option<&str>,
        override_false: Option<&str>,
        format_option: Option<ExcelFormat>,
    ) -> PyResult<()> {
        let worksheet = self
            .workbook
            .worksheet_from_index(self.active_worksheet_index)
            .unwrap();
        let override_value = if value { override_true } else { override_false };
        if format_option.is_none() {
            match override_value {
                Some(override_value) => {
                    worksheet.write_string(row, column, override_value).unwrap()
                }
                None => worksheet.write_boolean(row, column, value).unwrap(),
            };
        } else {
            let format = format::create_format(format_option.unwrap());
            match override_value {
                Some(override_value) => worksheet
                    .write_string_with_format(row, column, override_value, &format)
                    .unwrap(),
                None => worksheet
                    .write_boolean_with_format(row, column, value, &format)
                    .unwrap(),
            };
        }
        Ok(())
    }

    #[pyo3(signature = (start_row, start_column, end_row, end_column, format_option=None))]
    /// Worksheet handler for merging a range of cells. This will not do any
    /// writing to the cell values. To write values, use either
    /// `write_string_and_merge_range` or `write_number_and_merge_range`.
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
        start_row: u32,
        start_column: u16,
        end_row: u32,
        end_column: u16,
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

    #[pyo3(signature = (start_row, start_column, end_row, end_column, value, format_option=None))]
    /// Worksheet handler for merging a range of cells and writing string value into the merged cells.
    ///
    /// ## Parameters
    /// - `start_row`: The start row index of the range
    /// - `start_column`: The start column index of the range
    /// - `end_row`: The end row index of the range
    /// - `end_column`: The end column index of the range
    /// - `value`: The string value to write
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
    ///     workbook.write_string_and_merge_range(0, 0, 0, 2, "Hello World!", format_option)
    ///
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn write_string_and_merge_range(
        &mut self,
        start_row: u32,
        start_column: u16,
        end_row: u32,
        end_column: u16,
        value: &str,
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
                    value,
                    &Format::new(),
                )
                .unwrap();
        } else {
            let format = format::create_format(format_option.unwrap());
            worksheet
                .merge_range(start_row, start_column, end_row, end_column, value, &format)
                .unwrap();
        }
        Ok(())
    }

    #[pyo3(signature = (start_row, start_column, end_row, end_column, value, format_option=None))]
    /// Worksheet handler for merging a range of cells and writing numeric values into the merged cells.
    /// By default, values are passed as a float (`f64`) type. To specify whether to write intenger
    /// or float, use the `format_option` and set the `num_format` field.
    ///
    /// ## Parameters
    /// - `start_row`: The start row index of the range
    /// - `start_column`: The start column index of the range
    /// - `end_row`: The end row index of the range
    /// - `end_column`: The end column index of the range
    /// - `value`: The string value to write
    /// - `format_option`: The format of the cell _(optional)_
    ///
    /// ## Examples
    /// The following example demonstrates merging cells and writing numeric value into the merged cells in a worksheet.
    /// ```
    /// from pyaccelsx import ExcelWorkbook, ExcelFormat
    ///
    /// def main():
    ///     workbook = ExcelWorkbook()
    ///     workbook.add_worksheet()
    ///     
    ///     format_option = ExcelFormat(align="center", border=True, num_format="#,##0.00")
    ///     workbook.write_number_and_merge_range(0, 0, 0, 2, 12.45, format_option)
    ///
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn write_number_and_merge_range(
        &mut self,
        start_row: u32,
        start_column: u16,
        end_row: u32,
        end_column: u16,
        value: f64,
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
            worksheet
                .write_number_with_format(start_row, start_column, value, &Format::new())
                .unwrap();
        } else {
            let format = format::create_format(format_option.unwrap());
            worksheet
                .merge_range(start_row, start_column, end_row, end_column, "", &format)
                .unwrap();
            worksheet
                .write_number_with_format(start_row, start_column, value, &format)
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
    ///     workbook.write_string(0, 0, "Hello World!")
    ///     workbook.set_column_width(0, 20)
    ///
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn set_column_width(&mut self, column: u16, width: u16) -> PyResult<()> {
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
    ///     workbook.write_string(0, 0, "Hello World!")
    ///     // This freezes the first row and first column
    ///     workbook.freeze_panes(0, 0)
    ///
    ///     workbook.save("example.xlsx")
    /// ```
    pub fn freeze_panes(&mut self, row: u32, column: u16) -> PyResult<()> {
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
        Self::new()
    }
}
