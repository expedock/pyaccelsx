use pyo3::prelude::*;

#[derive(FromPyObject)]
pub enum ValueType {
    #[pyo3(transparent, annotation = "str")]
    String(String),
    #[pyo3(transparent, annotation = "bool")]
    Bool(bool),
    #[pyo3(transparent, annotation = "int")]
    Int(f64),
    #[pyo3(transparent, annotation = "float")]
    Float(f64),
}
