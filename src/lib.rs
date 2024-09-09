pub mod format;
pub mod util;
pub mod workbook;
pub mod writer;

use format::ExcelFormat;
use pyo3::prelude::*;
use workbook::ExcelWorkbook;

/// A Python module implemented in Rust.
#[pymodule]
fn pyaccelsx(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<ExcelWorkbook>()?;
    m.add_class::<ExcelFormat>()?;
    Ok(())
}
