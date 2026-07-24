//! streaming spreadsheetml parser + serializer over `xlsx_model`. parse treats
//! every byte as attacker-controlled: depth/count caps, no file-sized allocation.

mod read;
mod styles;
mod write;
mod xml;

pub use read::parse_workbook;
pub use write::serialize_workbook;

/// hard nesting limit for xml elements; deeper input is rejected as hostile.
pub const MAX_DEPTH: usize = 64;

/// upper bound on cells parsed across a single worksheet stream.
pub const MAX_CELLS: u64 = 10_000_000;

/// upper bound on entries in the shared string table.
pub const MAX_SHARED_STRINGS: usize = 10_000_000;

/// upper bound on workbook- and sheet-scoped defined names.
pub const MAX_DEFINED_NAMES: usize = 65_536;

/// upper bound on hyperlinks in one worksheet.
pub const MAX_HYPERLINKS: usize = 65_536;

/// upper bound on entries in any single style pool (fonts, fills, borders,
/// cellXfs, numFmts).
pub const MAX_STYLE_ENTRIES: usize = 65_536;

/// everything that can go wrong turning bytes into a workbook (or back).
#[derive(Debug, Clone, PartialEq)]
pub enum ParseError {
    /// a required opc part was absent.
    MissingPart(String),
    /// well-formedness or decoding failure reported by quick-xml.
    Xml(String),
    /// structurally valid xml that violates the spreadsheetml shape.
    Malformed(String),
    /// element nesting exceeded [`MAX_DEPTH`].
    DepthExceeded,
    /// a worksheet declared more cells than [`MAX_CELLS`].
    TooManyCells,
    /// the shared string table exceeded [`MAX_SHARED_STRINGS`].
    TooManyStrings,
    /// the defined-name table exceeded [`MAX_DEFINED_NAMES`].
    TooManyDefinedNames,
    /// a worksheet exceeded [`MAX_HYPERLINKS`].
    TooManyHyperlinks,
    /// a style pool exceeded [`MAX_STYLE_ENTRIES`].
    TooManyStyles,
}

impl core::fmt::Display for ParseError {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        match self {
            ParseError::MissingPart(p) => write!(f, "missing part: {p}"),
            ParseError::Xml(e) => write!(f, "xml error: {e}"),
            ParseError::Malformed(m) => write!(f, "malformed spreadsheetml: {m}"),
            ParseError::DepthExceeded => write!(f, "xml nesting exceeded depth cap"),
            ParseError::TooManyCells => write!(f, "worksheet cell count exceeded cap"),
            ParseError::TooManyStrings => write!(f, "shared string count exceeded cap"),
            ParseError::TooManyDefinedNames => write!(f, "defined name count exceeded cap"),
            ParseError::TooManyHyperlinks => write!(f, "worksheet hyperlink count exceeded cap"),
            ParseError::TooManyStyles => write!(f, "style pool count exceeded cap"),
        }
    }
}

impl std::error::Error for ParseError {}

#[cfg(test)]
mod tests;
