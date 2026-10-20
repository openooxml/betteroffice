use crate::Color;

/// A half-open UTF-8 byte range into the cell's existing plain string.
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq)]
pub struct RichRun {
    pub start: u32,
    pub end: u32,
    pub style: RunStyle,
}

/// Run overrides; absent properties inherit the cell font.
#[non_exhaustive]
#[derive(Debug, Clone, Default, PartialEq)]
pub struct RunStyle {
    pub name: Option<String>,
    pub size_pt: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub strike: Option<bool>,
    pub underline: Option<RunUnderline>,
    pub color: Option<Color>,
    pub baseline: Option<RunBaseline>,
}

impl RichRun {
    pub fn new(start: u32, end: u32, style: RunStyle) -> Self {
        Self { start, end, style }
    }
}

#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RunUnderline {
    None,
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
}

#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RunBaseline {
    Baseline,
    Superscript,
    Subscript,
}
