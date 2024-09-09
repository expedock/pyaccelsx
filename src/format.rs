/// This module contains the formatting for the Excel workbook.
use pyo3::prelude::*;
use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, FormatUnderline};

/// The `ExcelFormat` contains the format options passed from Python
/// to Rust, and used to create a custom `Format` object depending on
/// the configured format options.
///
/// ## Examples
/// The following example demonstrates creating a `rust_xlsxwriter::Format` object.
/// ```
/// from pyaccelsx import ExcelWorkbook, ExcelFormat
///
/// def main():
///     workbook = ExcelWorkbook()
///     workbook.add_worksheet()
///
///     format_option = ExcelFormat(
///         align="center",
///         bg_color="FFFF00",
///         bold=True,
///     )
///
///     workbook.write_string(0, 0, "Hello World!", format_option)
///
///     workbook.save("example.xlsx")
/// ```
#[pyclass(get_all, set_all)]
#[derive(Clone)]
pub struct ExcelFormat {
    align: Option<String>,
    bg_color: Option<String>,
    bold: Option<bool>,
    border: Option<bool>,
    border_top: Option<bool>,
    border_bottom: Option<bool>,
    border_left: Option<bool>,
    border_right: Option<bool>,
    font_color: Option<String>,
    num_format: Option<String>,
    underline: Option<String>,
}

#[pymethods]
impl ExcelFormat {
    #[new]
    #[pyo3(signature = (
        align=None,
        bg_color=None,
        bold=None,
        border=None,
        border_top=None,
        border_bottom=None,
        border_left=None,
        border_right=None,
        font_color=None,
        num_format=None,
        underline=None,
    ))]
    pub fn new(
        align: Option<String>,
        bg_color: Option<String>,
        bold: Option<bool>,
        border: Option<bool>,
        border_top: Option<bool>,
        border_bottom: Option<bool>,
        border_left: Option<bool>,
        border_right: Option<bool>,
        font_color: Option<String>,
        num_format: Option<String>,
        underline: Option<String>,
    ) -> ExcelFormat {
        ExcelFormat {
            align,
            bg_color,
            bold,
            border,
            border_top,
            border_bottom,
            border_left,
            border_right,
            font_color,
            num_format,
            underline,
        }
    }
}

/// Creates a `rust_xlsxwriter::Format` object from the `ExcelFormat`
/// options passed from Python.
///
/// ## Parameters
/// - `format_option`: The format options passed from Python
///
/// ## Returns
/// - A `rust_xlsxwriter::Format` object
pub fn create_format(format_option: ExcelFormat) -> Format {
    let mut format = Format::new();

    if let Some(align) = format_option.align {
        format = format.set_align(match align.as_str() {
            "left" => FormatAlign::Left,
            "center" => FormatAlign::Center,
            "right" => FormatAlign::Right,
            "fill" => FormatAlign::Fill,
            "justify" => FormatAlign::Justify,
            "center_across" => FormatAlign::CenterAcross,
            "distributed" => FormatAlign::Distributed,
            "top" => FormatAlign::Top,
            "bottom" => FormatAlign::Bottom,
            "vertical_center" => FormatAlign::VerticalCenter,
            "vertical_distributed" => FormatAlign::VerticalDistributed,
            "vertical_justify" => FormatAlign::VerticalJustify,
            _ => FormatAlign::General,
        })
    }

    if let Some(bg_color) = format_option.bg_color {
        format = format.set_background_color(bg_color.as_str());
    }

    if format_option.bold.unwrap_or(false) {
        format = format.set_bold();
    }

    if format_option.border.unwrap_or(false) {
        format = format.set_border(FormatBorder::Thin);
    } else {
        if format_option.border_top.unwrap_or(false) {
            format = format.set_border_top(FormatBorder::Thin);
        }

        if format_option.border_bottom.unwrap_or(false) {
            format = format.set_border_bottom(FormatBorder::Thin);
        }

        if format_option.border_left.unwrap_or(false) {
            format = format.set_border_left(FormatBorder::Thin);
        }

        if format_option.border_right.unwrap_or(false) {
            format = format.set_border_right(FormatBorder::Thin);
        }
    }

    if let Some(color) = format_option.font_color {
        format = format.set_font_color(color.as_str());
    }

    if let Some(num_format) = format_option.num_format {
        format = format.set_num_format(num_format.as_str());
    }

    if let Some(underline) = format_option.underline {
        format = format.set_underline(match underline.as_str() {
            "single" => FormatUnderline::Single,
            "double" => FormatUnderline::Double,
            "single_accounting" => FormatUnderline::SingleAccounting,
            "double_accounting" => FormatUnderline::DoubleAccounting,
            _ => FormatUnderline::Single,
        });
    }

    return format;
}
