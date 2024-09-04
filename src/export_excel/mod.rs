pub mod format;
pub mod workbook;

pub use format::FormatOption;
use pyo3::prelude::*;
pub use workbook::ExcelWorkbook;

/// A Python module for writing into Excel using rust_xlsxwriter.
#[pymodule]
pub fn export_excel(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<ExcelWorkbook>()?;
    m.add_class::<FormatOption>()?;
    Ok(())
}
