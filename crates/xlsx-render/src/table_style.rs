//! The table style element formats a styled table gives each of its cells.

use xlsx_model::TableStyleElementType::{self, *};
use xlsx_model::{CellRef, DifferentialFormat, RowId, SheetId, Table};

/// The styled tables of one sheet.
pub(crate) struct SheetTables<'a> {
    tables: Vec<Bands<'a>>,
}

/// A table with its row bands and stripe sizes worked out once.
struct Bands<'a> {
    table: &'a Table,
    header: Option<(RowId, RowId)>,
    totals: Option<(RowId, RowId)>,
    data_top: Option<RowId>,
    rows: Stripe,
    cols: Stripe,
}

/// Alternating stripe bands: the first `first` of every `period` take the
/// first element.
struct Stripe {
    first: u32,
    period: u32,
    bits: (u16, u16),
}

impl Stripe {
    fn new(table: &Table, first: TableStyleElementType, second: TableStyleElementType) -> Self {
        let size = |kind| {
            table
                .style
                .iter()
                .find(|element| element.kind == kind)
                .map_or(1, |element| element.size.max(1))
        };
        let first_size = size(first);
        Self {
            first: first_size,
            period: first_size.saturating_add(size(second)),
            bits: (bit(first), bit(second)),
        }
    }

    fn at(&self, offset: u32) -> u16 {
        if offset % self.period < self.first {
            self.bits.0
        } else {
            self.bits.1
        }
    }
}

impl<'a> SheetTables<'a> {
    pub(crate) fn new(tables: &'a [Table], sheet: SheetId) -> Self {
        Self {
            tables: tables
                .iter()
                .filter(|table| table.sheet == sheet && !table.style.is_empty())
                .map(|table| Bands {
                    table,
                    header: table.header_range(),
                    totals: table.totals_range(),
                    data_top: table.data_rows().map(|(top, _)| top),
                    rows: Stripe::new(table, FirstRowStripe, SecondRowStripe),
                    cols: Stripe::new(table, FirstColumnStripe, SecondColumnStripe),
                })
                .collect(),
        }
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.tables.is_empty()
    }

    /// The table covering `at` and a bit per element type that applies there.
    pub(crate) fn locate(&self, at: CellRef) -> Option<(usize, u16)> {
        let (index, bands) = self
            .tables
            .iter()
            .enumerate()
            .find(|(_, bands)| bands.table.range.contains(at))?;
        let range = bands.table.range;
        let within = |band: Option<(RowId, RowId)>| {
            band.is_some_and(|(top, bottom)| (top..=bottom).contains(&at.row))
        };
        let first = at.col == range.start.col;
        let last = at.col == range.end.col;
        let mut mask = bit(WholeTable) | when(first, FirstColumn) | when(last, LastColumn);
        if within(bands.header) {
            mask |= bit(HeaderRow) | when(first, FirstHeaderCell) | when(last, LastHeaderCell);
        } else if within(bands.totals) {
            mask |= bit(TotalRow) | when(first, FirstTotalCell) | when(last, LastTotalCell);
        } else if let Some(top) = bands.data_top {
            mask |= bands.rows.at(at.row - top) | bands.cols.at(at.col - range.start.col);
        }
        Some((index, mask))
    }

    /// What the elements in `mask` of table `index` merge to; for each facet the
    /// highest element that sets it wins.
    pub(crate) fn format(&self, index: usize, mask: u16) -> DifferentialFormat {
        let mut merged = DifferentialFormat::default();
        for element in &self.tables[index].table.style {
            if mask & bit(element.kind) == 0 {
                continue;
            }
            let format = &element.format;
            merged.font_color = format.font_color.clone().or(merged.font_color);
            merged.bold = format.bold.or(merged.bold);
            merged.italic = format.italic.or(merged.italic);
            merged.fill = format.fill.clone().or(merged.fill);
        }
        merged
    }
}

fn bit(kind: TableStyleElementType) -> u16 {
    1 << kind as u16
}

fn when(on: bool, kind: TableStyleElementType) -> u16 {
    if on { bit(kind) } else { 0 }
}
