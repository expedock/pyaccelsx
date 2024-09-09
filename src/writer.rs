use pyo3::prelude::*;
use rust_xlsxwriter::{ColNum, RowNum, Worksheet};

use crate::format::{self, ExcelFormat};

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
    worksheet: &mut Worksheet,
    row: RowNum,
    column: ColNum,
    value: &str,
    format_option: Option<ExcelFormat>,
) -> PyResult<()> {
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
    worksheet: &mut Worksheet,
    row: RowNum,
    column: ColNum,
    value: f64,
    format_option: Option<ExcelFormat>,
) -> PyResult<()> {
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

/// Worksheet handler for writing boolean values. By default, values
/// are written as `True` or `False`. To specify an override string value
/// to replace the boolean values, use the `override_value` field.
///
/// ## Parameters
/// - `row`: The row index of the cell
/// - `column`: The column index of the cell
/// - `value`: The boolean value to write
/// - `override_true_value`: The string value to override `True` value _(optional)_
/// - `override_false_value`: The string value to override `False` value _(optional)_
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
    worksheet: &mut Worksheet,
    row: RowNum,
    column: ColNum,
    value: bool,
    override_true_value: Option<String>,
    override_false_value: Option<String>,
    format_option: Option<ExcelFormat>,
) -> PyResult<()> {
    let override_value = if value {
        override_true_value
    } else {
        override_false_value
    };
    if format_option.is_none() {
        match override_value {
            Some(override_value) => worksheet.write_string(row, column, override_value).unwrap(),
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
    worksheet: &mut Worksheet,
    row: u32,
    column: u16,
    override_value: Option<String>,
    format_option: Option<ExcelFormat>,
) -> PyResult<()> {
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
