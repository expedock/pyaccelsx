pub enum ValueType<T> {
    String(T),
    Float(T),
    Int(T),
    Bool(T),
    None,
}