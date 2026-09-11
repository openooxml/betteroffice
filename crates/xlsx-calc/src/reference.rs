//! Whole-column formula references, distinct from finite cell rectangles.

use xlsx_model::addr::{AddrError, MAX_ROWS, col_to_letters};
use xlsx_model::{CellRange, CellRef};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ColumnRange {
    pub start: u32,
    pub end: u32,
    pub abs_start: bool,
    pub abs_end: bool,
}

impl ColumnRange {
    pub fn parse_a1(source: &str) -> Result<Self, AddrError> {
        let (start, end) = source.split_once(':').ok_or(AddrError::Malformed)?;
        let mut start = column(start)?;
        let mut end = column(end)?;
        if start.col > end.col {
            std::mem::swap(&mut start, &mut end);
        }
        Ok(Self {
            start: start.col,
            end: end.col,
            abs_start: start.abs_col,
            abs_end: end.abs_col,
        })
    }

    pub fn cell_range(&self) -> CellRange {
        CellRange {
            start: CellRef {
                abs_col: self.abs_start,
                ..CellRef::new(0, self.start)
            },
            end: CellRef {
                abs_col: self.abs_end,
                ..CellRef::new(MAX_ROWS - 1, self.end)
            },
        }
    }

    pub fn to_a1(&self) -> String {
        format!(
            "{}{}:{}{}",
            if self.abs_start { "$" } else { "" },
            col_to_letters(self.start),
            if self.abs_end { "$" } else { "" },
            col_to_letters(self.end),
        )
    }
}

pub(crate) fn column(source: &str) -> Result<CellRef, AddrError> {
    let letters = source.strip_prefix('$').unwrap_or(source);
    if letters.is_empty() || !letters.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return Err(AddrError::Malformed);
    }
    CellRef::parse_a1(&format!("{}1", source.to_ascii_uppercase()))
}
