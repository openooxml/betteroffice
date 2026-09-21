//! Whole-column formula references, distinct from finite cell rectangles, plus
//! the rectangle arithmetic OFFSET is defined by.

use std::cmp::Ordering;

use xlsx_model::addr::{AddrError, MAX_COLS, MAX_ROWS, col_to_letters};
use xlsx_model::{CellRange, CellRef, ErrorValue, Table};

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

/// The band of a table a structured reference selects. An empty selector list
/// means [`TableBand::Data`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableBand {
    All,
    Headers,
    Data,
    Totals,
    ThisRow,
}

impl TableBand {
    /// `#Headers` and friends as the file spells them, case-insensitively.
    pub fn parse(source: &str) -> Option<Self> {
        Some(match source.to_ascii_lowercase().as_str() {
            "#all" => TableBand::All,
            "#headers" => TableBand::Headers,
            "#data" => TableBand::Data,
            "#totals" => TableBand::Totals,
            "#this row" => TableBand::ThisRow,
            _ => return None,
        })
    }

    pub fn keyword(self) -> &'static str {
        match self {
            TableBand::All => "#All",
            TableBand::Headers => "#Headers",
            TableBand::Data => "#Data",
            TableBand::Totals => "#Totals",
            TableBand::ThisRow => "#This Row",
        }
    }
}

/// The bracketed body of a structured reference: which bands, and which column
/// or column span.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TableSpec {
    pub bands: Vec<TableBand>,
    pub first_column: Option<String>,
    pub last_column: Option<String>,
}

/// Bracket items one structured reference may hold.
pub(crate) const MAX_TABLE_SPEC_ITEMS: usize = 16;

/// The rectangle a structured reference designates. `cell` is the formula's own
/// cell, which `#This Row` needs; a missing or outside-the-table cell is
/// #VALUE!, and a named band or column the table does not have is #REF!.
pub(crate) fn table_rect(
    table: &Table,
    spec: &TableSpec,
    cell: Option<CellRef>,
) -> Result<CellRange, ErrorValue> {
    const DEFAULT: [TableBand; 1] = [TableBand::Data];
    let bands = if spec.bands.is_empty() {
        &DEFAULT[..]
    } else {
        &spec.bands[..]
    };
    let mut top = u32::MAX;
    let mut bottom = 0;
    for band in bands {
        let span = match band {
            TableBand::All => Some((table.range.start.row, table.range.end.row)),
            TableBand::Headers => table.header_range(),
            TableBand::Data => table.data_rows(),
            TableBand::Totals => table.totals_range(),
            TableBand::ThisRow => {
                let row = cell.ok_or(ErrorValue::Value)?.row;
                if !(table.range.start.row..=table.range.end.row).contains(&row) {
                    return Err(ErrorValue::Value);
                }
                Some((row, row))
            }
        };
        let (band_top, band_bottom) = span.ok_or(ErrorValue::Ref)?;
        top = top.min(band_top);
        bottom = bottom.max(band_bottom);
    }
    if top > bottom {
        return Err(ErrorValue::Ref);
    }
    let width = table.range.end.col - table.range.start.col;
    let first = match &spec.first_column {
        Some(name) => table.column_index(name).ok_or(ErrorValue::Ref)?,
        None => 0,
    };
    let last = match &spec.last_column {
        Some(name) => table.column_index(name).ok_or(ErrorValue::Ref)?,
        None if spec.first_column.is_some() => first,
        None => width,
    };
    let (first, last) = (first.min(last), first.max(last));
    Ok(CellRange::new(
        CellRef::new(top, table.range.start.col + first),
        CellRef::new(bottom, table.range.start.col + last),
    ))
}

/// the rectangle `OFFSET(anchor, rows, cols, height, width)` designates: the
/// anchor's top-left shifted by `rows`/`cols`, sized `height` x `width` (each
/// defaulting to the anchor's own extent). a negative size extends back from
/// the shifted corner. `None` is #REF!: a zero size, or a rectangle that leaves
/// the sheet.
pub(crate) fn offset_rect(
    anchor: CellRange,
    rows: i64,
    cols: i64,
    height: Option<i64>,
    width: Option<i64>,
) -> Option<CellRange> {
    let height = height.unwrap_or_else(|| i64::from(anchor.end.row - anchor.start.row) + 1);
    let width = width.unwrap_or_else(|| i64::from(anchor.end.col - anchor.start.col) + 1);
    let (top, bottom) = offset_span(anchor.start.row.into(), rows, height, MAX_ROWS.into())?;
    let (left, right) = offset_span(anchor.start.col.into(), cols, width, MAX_COLS.into())?;
    Some(CellRange::new(
        CellRef::new(top, left),
        CellRef::new(bottom, right),
    ))
}

/// one axis of an OFFSET rectangle, as inclusive 0-based bounds.
fn offset_span(origin: i64, delta: i64, size: i64, limit: i64) -> Option<(u32, u32)> {
    let shifted = origin.checked_add(delta)?;
    let (start, end) = match size.cmp(&0) {
        Ordering::Greater => (shifted, shifted.checked_add(size - 1)?),
        Ordering::Less => (shifted.checked_add(size + 1)?, shifted),
        Ordering::Equal => return None,
    };
    (start >= 0 && end < limit).then_some((start as u32, end as u32))
}

pub(crate) fn column(source: &str) -> Result<CellRef, AddrError> {
    let letters = source.strip_prefix('$').unwrap_or(source);
    if letters.is_empty() || !letters.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return Err(AddrError::Malformed);
    }
    CellRef::parse_a1(&format!("{}1", source.to_ascii_uppercase()))
}
