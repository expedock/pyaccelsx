use pyo3::prelude::*;
use rust_xlsxwriter::{ColNum, RowNum, Worksheet};

use crate::format::{self, ExcelFormat};

const MAX_LENGTH: usize = 32767;

pub fn write_string(
    worksheet: &mut Worksheet,
    row: RowNum,
    column: ColNum,
    value: String,
    format_option: Option<ExcelFormat>,
) -> PyResult<()> {
    let mut value = value;
    if value.len() > MAX_LENGTH {
        // Truncate the string
        value.truncate(MAX_LENGTH);
    }
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

pub fn write_null(
    worksheet: &mut Worksheet,
    row: RowNum,
    column: ColNum,
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
