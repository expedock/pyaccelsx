use super::format::{self, ExcelFormat};
use pyo3::prelude::*;
use rust_xlsxwriter::{Format, Workbook};

const DEFAULT_SHEETNAME: &str = "Sheet1";

#[pyclass]
pub struct ExcelWorkbook {
    workbook: Workbook,
    active_worksheet_name: String,
}

#[pymethods]
impl ExcelWorkbook {
    #[new]
    pub fn new() -> ExcelWorkbook {
        let workbook = Workbook::new();
        let active_worksheet_name = DEFAULT_SHEETNAME.to_string();
        ExcelWorkbook {
            workbook,
            active_worksheet_name,
        }
    }

    /// Add a new worksheet to the workbook.
    pub fn add_worksheet(&mut self, name: &str) {
        self.workbook.add_worksheet().set_name(name).unwrap();
        self.active_worksheet_name = name.to_string();
    }

    /// Save the workbook into the specified path.
    pub fn save(&mut self, path: &str) {
        self.workbook.save(path).unwrap();
    }

    #[pyo3(signature = (row, column, format_option=None))]
    pub fn write_blank(&mut self, row: u32, column: u16, format_option: Option<ExcelFormat>) {
        let worksheet = self
            .workbook
            .worksheet_from_name(&self.active_worksheet_name)
            .unwrap();
        if format_option.is_none() {
            worksheet.write(row, column, "").unwrap();
        } else {
            let format = format::create_format(format_option.unwrap());
            worksheet
                .write_with_format(row, column, "", &format)
                .unwrap();
        }
    }

    #[pyo3(signature = (row, column, value, format_option=None))]
    pub fn write_string(
        &mut self,
        row: u32,
        column: u16,
        value: &str,
        format_option: Option<ExcelFormat>,
    ) {
        let worksheet = self
            .workbook
            .worksheet_from_name(&self.active_worksheet_name)
            .unwrap();
        if format_option.is_none() {
            worksheet.write(row, column, value).unwrap();
        } else {
            let format = format::create_format(format_option.unwrap());
            worksheet
                .write_with_format(row, column, value, &format)
                .unwrap();
        }
    }

    #[pyo3(signature = (row, column, value, format_option=None))]
    pub fn write_number(
        &mut self,
        row: u32,
        column: u16,
        value: f64,
        format_option: Option<ExcelFormat>,
    ) {
        let worksheet = self
            .workbook
            .worksheet_from_name(&self.active_worksheet_name)
            .unwrap();
        let format = format::create_format(format_option.unwrap());
        worksheet
            .write_number_with_format(row, column, value, &format)
            .unwrap();
    }

    #[pyo3(signature = (start_row, start_column, end_row, end_column, format_option=None))]
    pub fn merge_range(
        &mut self,
        start_row: u32,
        start_column: u16,
        end_row: u32,
        end_column: u16,
        format_option: Option<ExcelFormat>,
    ) {
        let worksheet = self
            .workbook
            .worksheet_from_name(&self.active_worksheet_name)
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
    }

    #[pyo3(signature = (start_row, start_column, end_row, end_column, value, format_option=None))]
    pub fn write_string_and_merge_range(
        &mut self,
        start_row: u32,
        start_column: u16,
        end_row: u32,
        end_column: u16,
        value: &str,
        format_option: Option<ExcelFormat>,
    ) {
        let worksheet = self
            .workbook
            .worksheet_from_name(&self.active_worksheet_name)
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
    }

    #[pyo3(signature = (start_row, start_column, end_row, end_column, value, format_option=None))]
    pub fn write_number_and_merge_range(
        &mut self,
        start_row: u32,
        start_column: u16,
        end_row: u32,
        end_column: u16,
        value: f64,
        format_option: Option<ExcelFormat>,
    ) {
        let worksheet = self
            .workbook
            .worksheet_from_name(&self.active_worksheet_name)
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
    }

    pub fn set_column_width(&mut self, column: u16, width: u16) {
        let worksheet = self
            .workbook
            .worksheet_from_name(&self.active_worksheet_name)
            .unwrap();
        worksheet.set_column_width(column, width).unwrap();
    }

    pub fn freeze_panes(&mut self, row: u32, column: u16) {
        let worksheet = self
            .workbook
            .worksheet_from_name(&self.active_worksheet_name)
            .unwrap();
        worksheet.set_freeze_panes(row, column).unwrap();
    }
}

impl Default for ExcelWorkbook {
    fn default() -> Self {
        Self::new()
    }
}
