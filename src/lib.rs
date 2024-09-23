mod format;
mod util;
mod workbook;
mod writer;

pub use crate::format::ExcelFormat;
pub use crate::workbook::ExcelWorkbook;

use pyo3::prelude::*;

/// A Python module implemented in Rust.
#[pymodule]
fn pyaccelsx(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<ExcelWorkbook>()?;
    m.add_class::<ExcelFormat>()?;
    Ok(())
}
