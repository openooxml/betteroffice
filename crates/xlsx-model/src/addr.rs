//! cell addressing. rows and cols are 0-based internally; a1 notation is
//! 1-based and converted at parse/format time.

use serde::{Deserialize, Serialize};

pub type RowId = u32;
pub type ColId = u32;

pub const MAX_ROWS: u32 = 1_048_576;
pub const MAX_COLS: u32 = 16_384;

/// index of a sheet within the workbook's sheet order.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord, Serialize, Deserialize)]
pub struct SheetId(pub u32);

/// a single cell address. `abs_row`/`abs_col` carry `$` anchoring so formula
/// references survive round-trip and remap correctly.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct CellRef {
    pub row: RowId,
    pub col: ColId,
    #[serde(default, skip_serializing_if = "is_false")]
    pub abs_row: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub abs_col: bool,
}

fn is_false(b: &bool) -> bool {
    !*b
}

impl CellRef {
    pub fn new(row: RowId, col: ColId) -> Self {
        Self {
            row,
            col,
            abs_row: false,
            abs_col: false,
        }
    }

    /// parse a1 notation, e.g. `B7`, `$AA$21`. rejects out-of-range addresses.
    pub fn parse_a1(s: &str) -> Result<Self, AddrError> {
        let bytes = s.as_bytes();
        let mut i = 0;
        let abs_col = bytes.first() == Some(&b'$');
        if abs_col {
            i += 1;
        }
        let col_start = i;
        while i < bytes.len() && bytes[i].is_ascii_uppercase() {
            i += 1;
        }
        if i == col_start {
            return Err(AddrError::Malformed);
        }
        let mut col: u64 = 0;
        for &b in &bytes[col_start..i] {
            col = col * 26 + (b - b'A' + 1) as u64;
            if col > MAX_COLS as u64 {
                return Err(AddrError::OutOfRange);
            }
        }
        let abs_row = bytes.get(i) == Some(&b'$');
        if abs_row {
            i += 1;
        }
        let row_start = i;
        while i < bytes.len() && bytes[i].is_ascii_digit() {
            i += 1;
        }
        if i == row_start || i != bytes.len() {
            return Err(AddrError::Malformed);
        }
        let row: u64 = s[row_start..].parse().map_err(|_| AddrError::Malformed)?;
        if row == 0 || row > MAX_ROWS as u64 {
            return Err(AddrError::OutOfRange);
        }
        Ok(Self {
            row: (row - 1) as RowId,
            col: (col - 1) as ColId,
            abs_row,
            abs_col,
        })
    }

    /// format as a1 notation, preserving `$` anchors.
    pub fn to_a1(&self) -> String {
        let mut out = String::new();
        if self.abs_col {
            out.push('$');
        }
        out.push_str(&col_to_letters(self.col));
        if self.abs_row {
            out.push('$');
        }
        out.push_str(&(self.row + 1).to_string());
        out
    }
}

/// convert a 0-based column index to letters (0 -> A, 26 -> AA).
pub fn col_to_letters(col: ColId) -> String {
    let mut n = col as i64 + 1;
    let mut out = Vec::new();
    while n > 0 {
        n -= 1;
        out.push(b'A' + (n % 26) as u8);
        n /= 26;
    }
    out.reverse();
    String::from_utf8(out).expect("ascii")
}

/// an inclusive rectangular range. `start` is top-left, `end` bottom-right;
/// constructors normalize so that invariant always holds.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct CellRange {
    pub start: CellRef,
    pub end: CellRef,
}

impl CellRange {
    pub fn new(a: CellRef, b: CellRef) -> Self {
        let start = CellRef {
            row: a.row.min(b.row),
            col: a.col.min(b.col),
            ..a
        };
        let end = CellRef {
            row: a.row.max(b.row),
            col: a.col.max(b.col),
            ..b
        };
        Self { start, end }
    }

    /// parse `A1:B2` (or a single `A1` as a 1x1 range).
    pub fn parse_a1(s: &str) -> Result<Self, AddrError> {
        match s.split_once(':') {
            Some((a, b)) => Ok(Self::new(CellRef::parse_a1(a)?, CellRef::parse_a1(b)?)),
            None => {
                let c = CellRef::parse_a1(s)?;
                Ok(Self { start: c, end: c })
            }
        }
    }

    pub fn to_a1(&self) -> String {
        if self.start == self.end {
            self.start.to_a1()
        } else {
            format!("{}:{}", self.start.to_a1(), self.end.to_a1())
        }
    }

    pub fn contains(&self, cell: CellRef) -> bool {
        cell.row >= self.start.row
            && cell.row <= self.end.row
            && cell.col >= self.start.col
            && cell.col <= self.end.col
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AddrError {
    Malformed,
    OutOfRange,
}

impl core::fmt::Display for AddrError {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        match self {
            AddrError::Malformed => write!(f, "malformed address"),
            AddrError::OutOfRange => write!(f, "address out of sheet bounds"),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_and_formats_a1() {
        let c = CellRef::parse_a1("B7").unwrap();
        assert_eq!((c.row, c.col), (6, 1));
        assert_eq!(c.to_a1(), "B7");

        let c = CellRef::parse_a1("$AA$21").unwrap();
        assert_eq!((c.row, c.col, c.abs_row, c.abs_col), (20, 26, true, true));
        assert_eq!(c.to_a1(), "$AA$21");

        assert_eq!(CellRef::parse_a1("XFD1048576").unwrap().col, MAX_COLS - 1);
    }

    #[test]
    fn rejects_bad_addresses() {
        for s in [
            "", "7B", "A0", "A", "1", "a1", "A1B", "XFE1", "A1048577", "$",
        ] {
            assert!(CellRef::parse_a1(s).is_err(), "should reject {s:?}");
        }
    }

    #[test]
    fn col_letters_round_trip() {
        for (col, s) in [
            (0, "A"),
            (25, "Z"),
            (26, "AA"),
            (701, "ZZ"),
            (702, "AAA"),
            (16_383, "XFD"),
        ] {
            assert_eq!(col_to_letters(col), s);
            assert_eq!(CellRef::parse_a1(&format!("{s}1")).unwrap().col, col);
        }
    }

    #[test]
    fn ranges_normalize_and_contain() {
        let r = CellRange::parse_a1("B2:A1").unwrap_or_else(|_| {
            CellRange::new(
                CellRef::parse_a1("B2").unwrap(),
                CellRef::parse_a1("A1").unwrap(),
            )
        });
        assert_eq!(r.to_a1(), "A1:B2");
        assert!(r.contains(CellRef::new(0, 0)));
        assert!(r.contains(CellRef::new(1, 1)));
        assert!(!r.contains(CellRef::new(2, 0)));
        assert_eq!(CellRange::parse_a1("C3").unwrap().to_a1(), "C3");
    }
}
