pub mod export_excel;
use pyo3::prelude::*;

fn register_child_module(parent_module: &Bound<'_, PyModule>) -> PyResult<()> {
    let export_excel_module = PyModule::new_bound(parent_module.py(), "export_excel")?;
    export_excel::export_excel(&export_excel_module)?;
    parent_module.add_submodule(&export_excel_module)
}

/// A Python module implemented in Rust.
#[pymodule]
fn pyaccelsx(m: &Bound<'_, PyModule>) -> PyResult<()> {
    register_child_module(m)?;
    Ok(())
}
