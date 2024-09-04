/// This module contains the formatting for the Excel workbook.
use pyo3::prelude::*;
use rust_xlsxwriter::{Format, FormatAlign, FormatBorder};

const AGGREGATE_LABEL_COLOR: &str = "666666";

/// Format option passed from Python
#[pyclass(get_all, set_all)]
#[derive(FromPyObject)]
pub struct FormatOption {
    align: Option<String>,
    bold: Option<bool>,
    borders: Option<Vec<bool>>,
    color_override: Option<String>,
    is_float: Option<bool>,
    is_integer: Option<bool>,
}

#[pymethods]
impl FormatOption {
    #[new]
    #[pyo3(signature = (align=None, bold=None, borders=None, color_override=None, is_float=None, is_integer=None))]
    pub fn new(
        align: Option<String>,
        bold: Option<bool>,
        borders: Option<Vec<bool>>,
        color_override: Option<String>,
        is_float: Option<bool>,
        is_integer: Option<bool>,
    ) -> FormatOption {
        FormatOption {
            align,
            bold,
            borders,
            color_override,
            is_float,
            is_integer,
        }
    }
}

pub fn custom_format(format_option: FormatOption) -> Format {
    let mut format = Format::new();

    match format_option.align {
        Some(align) => {
            format = format.set_align(match align.as_str() {
                "left" => FormatAlign::Left,
                "center" => FormatAlign::Center,
                "right" => FormatAlign::Right,
                _ => FormatAlign::Left,
            });
        }
        None => {}
    }

    match format_option.bold {
        Some(true) => {
            format = format.set_bold();
        }
        _ => {}
    }

    match format_option.is_integer {
        Some(true) => {
            format = format.set_num_format("#,##0");
        }
        _ => {}
    }

    match format_option.is_float {
        Some(true) => {
            format = format.set_num_format("#,##0.00");
        }
        _ => {}
    }

    match format_option.color_override {
        Some(color) => {
            format = format.set_font_color(color.as_str());
        }
        None => {}
    }

    if format_option.borders.is_some() {
        // Borders are in the order of top, bottom, left, right
        let vec = format_option.borders;
        let vec = vec.unwrap();
        let values = vec.iter();
        for (index, value) in values.enumerate() {
            match value {
                true => {
                    format = match index {
                        0 => format.set_border_top(FormatBorder::Thin),
                        1 => format.set_border_bottom(FormatBorder::Thin),
                        2 => format.set_border_left(FormatBorder::Thin),
                        3 => format.set_border_right(FormatBorder::Thin),
                        _ => format,
                    };
                }
                false => {}
            }
        }
    }

    return format;
}

/// Expedock-specific code
/// TODO: generalize
pub fn aggregate_label(row_position: &str) -> Format {
    let mut format = Format::new()
        .set_font_color(AGGREGATE_LABEL_COLOR)
        .set_border_left(FormatBorder::Thin);
    match row_position {
        "top" => format = format.set_border_top(FormatBorder::Thin),
        "bottom" => format = format.set_border_bottom(FormatBorder::Thin),
        _ => {}
    }
    return format;
}

/// Expedock-specific code
/// TODO: generalize
pub fn aggregate_value(row_position: &str, is_float: Option<bool>) -> Format {
    let mut format = Format::new()
        .set_align(FormatAlign::Right)
        .set_border_right(FormatBorder::Thin);

    match row_position {
        "top" => format = format.set_border_top(FormatBorder::Thin),
        "bottom" => format = format.set_border_bottom(FormatBorder::Thin),
        _ => {}
    }

    match is_float {
        Some(true) => format = format.set_bold().set_num_format("#,##0.00"),
        Some(false) => format = format.set_bold().set_num_format("#,##0"),
        _ => {}
    }

    return format;
}
