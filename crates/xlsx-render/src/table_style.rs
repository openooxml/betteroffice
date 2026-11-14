//! The table style element formats a styled table gives each of its cells.

use xlsx_model::TableStyleElementType::{self, *};
use xlsx_model::{CellRef, DifferentialFormat, SheetId, Table};

/// The styled tables of one sheet.
pub(crate) struct SheetTables<'a> {
    tables: Vec<&'a Table>,
}

impl<'a> SheetTables<'a> {
    pub(crate) fn new(tables: &'a [Table], sheet: SheetId) -> Self {
        Self {
            tables: tables
                .iter()
                .filter(|table| table.sheet == sheet && !table.style.is_empty())
                .collect(),
        }
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.tables.is_empty()
    }

    /// The table covering `at` and a bit per element type that applies there.
    pub(crate) fn locate(&self, at: CellRef) -> Option<(usize, u16)> {
        let (index, table) = self
            .tables
            .iter()
            .enumerate()
            .find(|(_, table)| table.range.contains(at))?;
        let within = |band: Option<(u32, u32)>| {
            band.is_some_and(|(top, bottom)| (top..=bottom).contains(&at.row))
        };
        let first = at.col == table.range.start.col;
        let last = at.col == table.range.end.col;
        let mut mask = bit(WholeTable) | when(first, FirstColumn) | when(last, LastColumn);
        if within(table.header_range()) {
            mask |= bit(HeaderRow) | when(first, FirstHeaderCell) | when(last, LastHeaderCell);
        } else if within(table.totals_range()) {
            mask |= bit(TotalRow) | when(first, FirstTotalCell) | when(last, LastTotalCell);
        } else if let Some((top, _)) = table.data_rows() {
            mask |= stripe(table, at.row - top, FirstRowStripe, SecondRowStripe);
            mask |= stripe(
                table,
                at.col - table.range.start.col,
                FirstColumnStripe,
                SecondColumnStripe,
            );
        }
        Some((index, mask))
    }

    /// What the elements in `mask` of table `index` merge to; for each facet the
    /// highest element that sets it wins.
    pub(crate) fn format(&self, index: usize, mask: u16) -> DifferentialFormat {
        let mut merged = DifferentialFormat::default();
        for element in &self.tables[index].style {
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

/// The stripe element of band `offset`, alternating `first` and `second` bands.
fn stripe(
    table: &Table,
    offset: u32,
    first: TableStyleElementType,
    second: TableStyleElementType,
) -> u16 {
    let size = |kind| {
        table
            .style
            .iter()
            .find(|element| element.kind == kind)
            .map_or(1, |element| element.size.max(1))
    };
    let first_size = size(first);
    if offset % first_size.saturating_add(size(second)) < first_size {
        bit(first)
    } else {
        bit(second)
    }
}
