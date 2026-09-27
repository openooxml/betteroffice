//! sparse workbook containers and the calc-facing cell-access trait.

use std::borrow::Cow;
use std::collections::BTreeMap;
use std::ops::{Range, RangeInclusive};

use serde::{Deserialize, Serialize};

use crate::addr::{CellRange, CellRef, ColId, RowId, SheetId};
use crate::chart::SheetChart;
use crate::date::DateSystem;
use crate::styles::Stylesheet;
use crate::value::CellValue;

/// upper bound on the cells one array formula may fill. a malformed or hostile
/// `ref` must not be able to ask for a sheet's worth of cells.
pub const MAX_SPILL_CELLS: usize = 262_144;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct FreezePane {
    pub rows: RowId,
    pub cols: ColId,
    pub top_left: CellRef,
}

impl FreezePane {
    pub fn new(rows: RowId, cols: ColId, top_left: CellRef) -> Self {
        Self {
            rows,
            cols,
            top_left,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct DefinedName {
    pub name: String,
    pub formula: String,
    pub local_sheet: Option<SheetId>,
    pub hidden: bool,
}

/// One `xl/tables/tableN.xml` definition: the rectangle a structured reference
/// resolves against, with the header and totals bands split out.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Table {
    pub name: String,
    pub sheet: SheetId,
    pub range: CellRange,
    pub header_rows: u32,
    pub totals_rows: u32,
    pub columns: Vec<String>,
}

impl Table {
    /// Inclusive row band holding data, `None` when the table has no data rows.
    pub fn data_rows(&self) -> Option<(RowId, RowId)> {
        let top = self.range.start.row.checked_add(self.header_rows)?;
        let bottom = self.range.end.row.checked_sub(self.totals_rows)?;
        (top <= bottom).then_some((top, bottom))
    }

    /// Inclusive header band, `None` when the table has no header row.
    pub fn header_range(&self) -> Option<(RowId, RowId)> {
        if self.header_rows == 0 {
            return None;
        }
        let bottom = self.range.start.row.checked_add(self.header_rows - 1)?;
        Some((self.range.start.row, bottom.min(self.range.end.row)))
    }

    /// Inclusive totals band, `None` when the table has no totals row.
    pub fn totals_range(&self) -> Option<(RowId, RowId)> {
        if self.totals_rows == 0 {
            return None;
        }
        let top = self.range.end.row.checked_sub(self.totals_rows - 1)?;
        Some((top, self.range.end.row))
    }

    /// 0-based index of `column` within the table, matched case-insensitively.
    pub fn column_index(&self, column: &str) -> Option<u32> {
        let index = self
            .columns
            .iter()
            .position(|name| name.eq_ignore_ascii_case(column))?;
        let index = u32::try_from(index).ok()?;
        let col = self.range.start.col.checked_add(index)?;
        (col <= self.range.end.col).then_some(index)
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Hyperlink {
    pub range: CellRange,
    pub external_target: Option<String>,
    pub location: Option<String>,
    pub tooltip: Option<String>,
    pub display: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Cell {
    pub value: CellValue,
    /// original formula text without the leading `=`, if any.
    pub formula: Option<String>,
    /// index into the workbook style table (cellXfs).
    pub style: Option<u32>,
}

/// cells to relocate: each source address with the address it moves to, or
/// `None` when the cell is refused.
type CellMoves = Vec<((RowId, ColId), Option<(RowId, ColId)>)>;

/// `sheetFormatPr` sizing defaults. `custom_height` is the author's claim that
/// every unsized row is pinned at `default_row_height_pt`; without it an
/// unsized row takes the height of its tallest content.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct SheetFormat {
    pub default_row_height_pt: Option<f64>,
    pub custom_height: bool,
    /// `zeroHeight`: every row without its own `ht` is hidden.
    pub zero_height: bool,
}

/// a `<col>` run's style: the `cellXfs` index its columns give a cell that names none.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct ColStyle {
    pub first: ColId,
    pub last: ColId,
    pub xf: u32,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Sheet {
    pub name: String,
    cells: BTreeMap<(RowId, ColId), Cell>,
    pub freeze_pane: Option<FreezePane>,
    pub hyperlinks: Vec<Hyperlink>,
    pub merges: Vec<CellRange>,
    pub col_widths: BTreeMap<ColId, f64>,
    pub row_heights: BTreeMap<RowId, f64>,
    /// parsed from `sheetFormatPr`; read by the renderer, never by the writer.
    pub format: SheetFormat,
    /// `<col>` style runs in source order; read by the renderer, never the writer.
    pub col_styles: Vec<ColStyle>,
    pub charts: Vec<SheetChart>,
    /// anchors of `t="array"` formulas mapped to the rectangle their result
    /// occupies. authored from the file, then kept current by recalc.
    array_formulas: BTreeMap<(RowId, ColId), CellRange>,
}

impl Sheet {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            ..Self::default()
        }
    }

    /// the rectangle the array formula anchored at `at` currently fills.
    pub fn array_formula(&self, at: CellRef) -> Option<CellRange> {
        self.array_formulas.get(&(at.row, at.col)).copied()
    }

    pub fn set_array_formula(&mut self, at: CellRef, spill: CellRange) {
        self.array_formulas.insert((at.row, at.col), spill);
    }

    pub fn clear_array_formula(&mut self, at: CellRef) {
        self.array_formulas.remove(&(at.row, at.col));
    }

    /// every array-formula anchor with its rectangle, in address order.
    pub fn array_formulas(&self) -> impl Iterator<Item = (CellRef, CellRange)> + '_ {
        self.array_formulas
            .iter()
            .map(|(&(row, col), &spill)| (CellRef::new(row, col), spill))
    }

    pub fn cell(&self, at: CellRef) -> Option<&Cell> {
        self.cells.get(&(at.row, at.col))
    }

    /// the style a `<col>` run gives a cell that names none; later runs win.
    /// only the fill and the row fit consult it; other facets read `Cell::style`.
    pub fn col_style(&self, col: ColId) -> Option<u32> {
        self.col_styles
            .iter()
            .rev()
            .find(|run| (run.first..=run.last).contains(&col))
            .map(|run| run.xf)
    }

    pub fn cell_mut(&mut self, at: CellRef) -> Option<&mut Cell> {
        self.cells.get_mut(&(at.row, at.col))
    }

    pub fn set_cell(&mut self, at: CellRef, cell: Cell) {
        if cell == Cell::default() {
            self.cells.remove(&(at.row, at.col));
        } else {
            self.cells.insert((at.row, at.col), cell);
        }
    }

    /// moves every occupied cell to the address `remap` names; the cells it
    /// refuses come back in address order. cells already at their target are
    /// left untouched, and a target several cells reach keeps the one from the
    /// largest source address.
    pub fn remap_cells(
        &mut self,
        remap: impl Fn(CellRef) -> Option<CellRef>,
    ) -> Vec<(CellRef, Cell)> {
        self.remap_array_formulas(&remap);
        let mut dropped = Vec::new();
        self.cells.retain(|_, cell| *cell != Cell::default());
        let mut plan: CellMoves = Vec::new();
        let mut split: Option<(RowId, ColId)> = None;
        let mut suffix = true;
        for &at in self.cells.keys() {
            let target = match remap(CellRef::new(at.0, at.1)) {
                Some(to) if (to.row, to.col) != at => Some((to.row, to.col)),
                Some(_) => {
                    if split.is_some() {
                        suffix = false;
                    }
                    continue;
                }
                None => None,
            };
            if split.is_none() {
                split = Some(at);
            }
            plan.push((at, target));
        }
        let Some(split) = split else {
            return dropped;
        };
        if suffix && plan.iter().all(|(_, to)| to.is_none_or(|to| to >= split)) {
            let mut rebuilt = BTreeMap::new();
            let mut tail = self.cells.split_off(&split).into_iter();
            for ((row, col), target) in plan {
                let (_, cell) = tail.next().expect("plan covers every split-off cell");
                match target {
                    Some(to) => {
                        rebuilt.insert(to, cell);
                    }
                    None => dropped.push((CellRef::new(row, col), cell)),
                }
            }
            self.cells.append(&mut rebuilt);
        } else {
            let old = std::mem::take(&mut self.cells);
            for ((row, col), cell) in old {
                let Some(to) = remap(CellRef::new(row, col)) else {
                    dropped.push((CellRef::new(row, col), cell));
                    continue;
                };
                self.cells.insert((to.row, to.col), cell);
            }
        }
        dropped
    }

    /// move each array anchor with its cell, translating its rectangle by the
    /// same delta; anchors the remap refuses lose their array identity.
    fn remap_array_formulas(&mut self, remap: &impl Fn(CellRef) -> Option<CellRef>) {
        if self.array_formulas.is_empty() {
            return;
        }
        let mut moved = BTreeMap::new();
        for (&(row, col), &spill) in &self.array_formulas {
            let Some(to) = remap(CellRef::new(row, col)) else {
                continue;
            };
            let rows = spill.end.row.saturating_sub(spill.start.row);
            let cols = spill.end.col.saturating_sub(spill.start.col);
            let end = CellRef::new(
                to.row.saturating_add(rows).min(crate::addr::MAX_ROWS - 1),
                to.col.saturating_add(cols).min(crate::addr::MAX_COLS - 1),
            );
            moved.insert((to.row, to.col), CellRange::new(to, end));
        }
        self.array_formulas = moved;
    }

    /// ordered iteration over occupied cells (row-major).
    pub fn iter_cells(&self) -> impl Iterator<Item = (CellRef, &Cell)> {
        self.cells
            .iter()
            .map(|(&(row, col), cell)| (CellRef::new(row, col), cell))
    }

    pub fn iter_cells_in_rect(
        &self,
        rows: Range<RowId>,
        cols: Range<ColId>,
    ) -> impl Iterator<Item = (CellRef, &Cell)> {
        let start_col = cols.start;
        let end_col = cols.end.max(start_col);
        rows.flat_map(move |row| {
            self.cells
                .range((row, start_col)..(row, end_col))
                .map(|(&(row, col), cell)| (CellRef::new(row, col), cell))
        })
    }

    /// stored cells inside `range`, row-major, without enumerating its empty
    /// positions: each step seeks the next stored cell.
    pub fn cells_in_range(&self, range: CellRange) -> SparseCells<'_> {
        SparseCells {
            cells: &self.cells,
            range,
            next: Some((range.start.row, range.start.col)),
            seeks: 0,
        }
    }

    /// whether `row` has no height: hidden, authored at zero, or unsized under
    /// `zeroHeight`. rendering and export share this reading.
    pub fn row_hidden(&self, row: RowId) -> bool {
        match self.row_heights.get(&row) {
            Some(&height) => no_extent(height),
            None => self.format.zero_height,
        }
    }

    /// whether `col` has no width: hidden or authored at zero.
    pub fn col_hidden(&self, col: ColId) -> bool {
        self.col_widths
            .get(&col)
            .is_some_and(|&width| no_extent(width))
    }

    /// maximal runs of rows inside `rows` that [`Sheet::row_hidden`] hides, as
    /// inclusive pairs in order; costs one step per sized row, not per row.
    pub fn hidden_row_spans(
        &self,
        rows: RangeInclusive<RowId>,
    ) -> impl Iterator<Item = (RowId, RowId)> + '_ {
        HiddenSpans::new(
            self.row_heights
                .range(rows.clone())
                .map(|(&row, &height)| (row, no_extent(height))),
            rows,
            self.format.zero_height,
        )
    }

    /// maximal runs of columns inside `cols` that [`Sheet::col_hidden`] hides.
    pub fn hidden_col_spans(
        &self,
        cols: RangeInclusive<ColId>,
    ) -> impl Iterator<Item = (ColId, ColId)> + '_ {
        HiddenSpans::new(
            self.col_widths
                .range(cols.clone())
                .map(|(&col, &width)| (col, no_extent(width))),
            cols,
            false,
        )
    }

    pub fn hyperlink_at(&self, at: CellRef) -> Option<&Hyperlink> {
        self.hyperlinks.iter().find(|link| link.range.contains(at))
    }

    pub fn used_range(&self) -> Option<CellRange> {
        let mut bounds = self
            .cells
            .keys()
            .map(|&(row, col)| CellRange::new(CellRef::new(row, col), CellRef::new(row, col)))
            .chain(self.hyperlinks.iter().map(|link| link.range));
        let first = bounds.next()?;
        let (mut min_r, mut max_r, mut min_c, mut max_c) = (
            first.start.row,
            first.end.row,
            first.start.col,
            first.end.col,
        );
        for range in bounds {
            let r = range.start.row;
            let c = range.start.col;
            min_r = min_r.min(r);
            max_r = max_r.max(range.end.row);
            min_c = min_c.min(c);
            max_c = max_c.max(range.end.col);
        }
        Some(CellRange::new(
            CellRef::new(min_r, min_c),
            CellRef::new(max_r, max_c),
        ))
    }
}

fn no_extent(extent: f64) -> bool {
    extent <= 0.0
}

/// runs of hidden indices from the sized entries of one axis; an index without
/// an entry is hidden when `unsized_hidden`.
struct HiddenSpans<I> {
    sized: I,
    next: Option<u32>,
    last: u32,
    unsized_hidden: bool,
    run: Option<(u32, u32)>,
}

impl<I: Iterator<Item = (u32, bool)>> HiddenSpans<I> {
    fn new(sized: I, span: RangeInclusive<u32>, unsized_hidden: bool) -> Self {
        let (first, last) = span.into_inner();
        Self {
            sized,
            next: (first <= last).then_some(first),
            last,
            unsized_hidden,
            run: None,
        }
    }

    /// classifies `start..=end`, returning the run a visible index closes.
    fn mark(&mut self, start: u32, end: u32, hidden: bool) -> Option<(u32, u32)> {
        if start > end {
            return None;
        }
        if hidden {
            self.run = Some((self.run.map_or(start, |(first, _)| first), end));
            None
        } else {
            self.run.take()
        }
    }
}

impl<I: Iterator<Item = (u32, bool)>> Iterator for HiddenSpans<I> {
    type Item = (u32, u32);

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            let Some(next) = self.next else {
                return self.run.take();
            };
            let closed = match self.sized.next() {
                Some((index, hidden)) => {
                    let gap = match index.checked_sub(1) {
                        Some(before) if index > next => {
                            self.mark(next, before, self.unsized_hidden)
                        }
                        _ => None,
                    };
                    self.next = index.checked_add(1).filter(|&after| after <= self.last);
                    gap.or(self.mark(index, index, hidden))
                }
                None => {
                    self.next = None;
                    self.mark(next, self.last, self.unsized_hidden)
                }
            };
            if closed.is_some() {
                return closed;
            }
        }
    }
}

/// cursor over the stored cells of one rectangle; see [`Sheet::cells_in_range`].
pub struct SparseCells<'a> {
    cells: &'a BTreeMap<(RowId, ColId), Cell>,
    range: CellRange,
    next: Option<(RowId, ColId)>,
    seeks: u64,
}

impl SparseCells<'_> {
    /// map lookups made so far; rows whose stored cells all sit outside the
    /// column band cost one each.
    pub fn seeks(&self) -> u64 {
        self.seeks
    }

    /// continue from the first column of the row after `row`.
    pub fn skip_row(&mut self, row: RowId) {
        self.next = (row < self.range.end.row).then(|| (row + 1, self.range.start.col));
    }
}

impl<'a> Iterator for SparseCells<'a> {
    type Item = (CellRef, &'a Cell);

    fn next(&mut self) -> Option<Self::Item> {
        let end = (self.range.end.row, self.range.end.col);
        loop {
            let from = self.next.filter(|from| *from <= end)?;
            self.seeks += 1;
            let Some((&(row, col), cell)) = self.cells.range(from..=end).next() else {
                self.next = None;
                return None;
            };
            if col < self.range.start.col {
                self.next = Some((row, self.range.start.col));
            } else if col > self.range.end.col {
                self.skip_row(row);
            } else {
                self.next = if col < self.range.end.col {
                    Some((row, col + 1))
                } else {
                    (row < self.range.end.row).then(|| (row + 1, self.range.start.col))
                };
                return Some((CellRef::new(row, col), cell));
            }
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
    pub date_system: DateSystem,
    pub defined_names: Vec<DefinedName>,
    /// shared string table as parsed; kept for round-trip fidelity.
    pub shared_strings: Vec<String>,
    /// parsed style tables + theme; a cell's `style` indexes `styles.cell_xfs`.
    pub styles: Stylesheet,
    /// table parts, in package order; structured references resolve through them.
    pub tables: Vec<Table>,
}

impl Workbook {
    pub fn sheet(&self, id: SheetId) -> Option<&Sheet> {
        self.sheets.get(id.0 as usize)
    }

    pub fn sheet_mut(&mut self, id: SheetId) -> Option<&mut Sheet> {
        self.sheets.get_mut(id.0 as usize)
    }

    pub fn sheet_by_name(&self, name: &str) -> Option<(SheetId, &Sheet)> {
        let name = name.to_lowercase();
        self.sheets
            .iter()
            .enumerate()
            .find(|(_, sheet)| sheet.name.to_lowercase() == name)
            .map(|(i, s)| (SheetId(i as u32), s))
    }

    pub fn table(&self, name: &str) -> Option<&Table> {
        self.tables
            .iter()
            .find(|table| table.name.eq_ignore_ascii_case(name))
    }

    pub fn defined_name(&self, sheet: SheetId, name: &str) -> Option<&DefinedName> {
        self.defined_names
            .iter()
            .find(|defined| {
                defined.local_sheet == Some(sheet) && defined.name.eq_ignore_ascii_case(name)
            })
            .or_else(|| {
                self.defined_names.iter().find(|defined| {
                    defined.local_sheet.is_none() && defined.name.eq_ignore_ascii_case(name)
                })
            })
    }
}

/// read access the calc engine evaluates through.
pub trait CellProvider {
    fn value(&self, sheet: SheetId, at: CellRef) -> CellValue;
    /// Borrowing variant of `value`; absent cells read as `CellValue::Empty`.
    fn value_cow(&self, sheet: SheetId, at: CellRef) -> Cow<'_, CellValue> {
        Cow::Owned(self.value(sheet, at))
    }
    fn formula(&self, sheet: SheetId, at: CellRef) -> Option<&str>;
    fn sheet_id(&self, name: &str) -> Option<SheetId>;
    fn defined_name(&self, _sheet: SheetId, _name: &str) -> Option<&DefinedName> {
        None
    }

    /// the table a structured reference names, matched case-insensitively.
    fn table(&self, _name: &str) -> Option<&Table> {
        None
    }

    /// rows worth materializing for a whole-column reference; `0` means the
    /// sheet is empty. bounds array evaluation to authored data.
    fn used_rows(&self, _sheet: SheetId) -> RowId {
        0
    }

    /// columns worth materializing for a whole-row reference; `0` means the
    /// sheet is empty. bounds array evaluation to authored data.
    fn used_cols(&self, _sheet: SheetId) -> ColId {
        0
    }

    /// the rectangle the array formula anchored at `at` fills, if any.
    fn spill_range(&self, _sheet: SheetId, _at: CellRef) -> Option<CellRange> {
        None
    }
}

impl CellProvider for Workbook {
    fn value(&self, sheet: SheetId, at: CellRef) -> CellValue {
        self.value_cow(sheet, at).into_owned()
    }

    fn value_cow(&self, sheet: SheetId, at: CellRef) -> Cow<'_, CellValue> {
        match self.sheet(sheet).and_then(|s| s.cell(at)) {
            Some(cell) => Cow::Borrowed(&cell.value),
            None => Cow::Owned(CellValue::Empty),
        }
    }

    fn formula(&self, sheet: SheetId, at: CellRef) -> Option<&str> {
        self.sheet(sheet)?.cell(at)?.formula.as_deref()
    }

    fn sheet_id(&self, name: &str) -> Option<SheetId> {
        self.sheet_by_name(name).map(|(id, _)| id)
    }

    fn defined_name(&self, sheet: SheetId, name: &str) -> Option<&DefinedName> {
        self.defined_name(sheet, name)
    }

    fn table(&self, name: &str) -> Option<&Table> {
        self.table(name)
    }

    fn used_rows(&self, sheet: SheetId) -> RowId {
        self.sheet(sheet)
            .and_then(Sheet::used_range)
            .map_or(0, |range| range.end.row.saturating_add(1))
    }

    fn used_cols(&self, sheet: SheetId) -> ColId {
        self.sheet(sheet)
            .and_then(Sheet::used_range)
            .map_or(0, |range| range.end.col.saturating_add(1))
    }

    fn spill_range(&self, sheet: SheetId, at: CellRef) -> Option<CellRange> {
        self.sheet(sheet)?.array_formula(at)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sparse_set_get_and_used_range() {
        let mut sheet = Sheet::new("Sheet1");
        assert!(sheet.used_range().is_none());

        let b2 = CellRef::parse_a1("B2").unwrap();
        let d7 = CellRef::parse_a1("D7").unwrap();
        sheet.set_cell(
            b2,
            Cell {
                value: CellValue::Number { value: 1.0 },
                ..Cell::default()
            },
        );
        sheet.set_cell(
            d7,
            Cell {
                value: CellValue::Text { value: "x".into() },
                ..Cell::default()
            },
        );

        assert_eq!(sheet.used_range().unwrap().to_a1(), "B2:D7");
        assert_eq!(sheet.iter_cells().count(), 2);

        sheet.set_cell(b2, Cell::default());
        assert_eq!(sheet.used_range().unwrap().to_a1(), "D7");
    }

    #[test]
    fn cells_in_range_seeks_only_stored_rows() {
        let mut sheet = Sheet::new("Sheet1");
        for a1 in ["A1", "C1", "B2", "D2", "C1000000", "B1048576"] {
            sheet.set_cell(
                CellRef::parse_a1(a1).unwrap(),
                Cell {
                    value: CellValue::Number { value: 1.0 },
                    ..Cell::default()
                },
            );
        }
        let range = CellRange::parse_a1("B1:C1048576").unwrap();
        let mut cells = sheet.cells_in_range(range);
        let found = cells.by_ref().map(|(at, _)| at.to_a1()).collect::<Vec<_>>();
        assert_eq!(found, ["C1", "B2", "C1000000", "B1048576"]);
        assert!(cells.seeks() <= 8, "{} seeks", cells.seeks());

        let mut skipping = sheet.cells_in_range(range);
        assert_eq!(skipping.next().unwrap().0.to_a1(), "C1");
        skipping.skip_row(1);
        assert_eq!(skipping.next().unwrap().0.to_a1(), "C1000000");
        assert_eq!(
            sheet
                .cells_in_range(CellRange::parse_a1("E1:F9").unwrap())
                .count(),
            0
        );
    }

    #[test]
    fn hidden_spans_agree_with_the_per_index_predicates() {
        let mut sheet = Sheet::new("Sheet1");
        for (row, height) in [
            (2, 0.0),
            (3, 0.0),
            (5, 12.0),
            (6, 0.0),
            (9, 15.0),
            (12, 0.0),
        ] {
            sheet.row_heights.insert(row, height);
        }
        for (col, width) in [(0, 0.0), (1, 8.0), (2, 0.0), (3, 0.0), (7, 0.0)] {
            sheet.col_widths.insert(col, width);
        }
        let brute = |hidden: &dyn Fn(u32) -> bool, first: u32, last: u32| {
            let mut spans: Vec<(u32, u32)> = Vec::new();
            for index in first..=last {
                if !hidden(index) {
                    continue;
                }
                match spans.last_mut() {
                    Some(span) if span.1 + 1 == index => span.1 = index,
                    _ => spans.push((index, index)),
                }
            }
            spans
        };
        for zero_height in [false, true] {
            sheet.format.zero_height = zero_height;
            for (first, last) in [(0, 14), (3, 9), (5, 5), (10, 11), (6, 6)] {
                assert_eq!(
                    sheet.hidden_row_spans(first..=last).collect::<Vec<_>>(),
                    brute(&|row| sheet.row_hidden(row), first, last),
                    "rows {first}..={last} zero_height={zero_height}"
                );
            }
        }
        for (first, last) in [(0, 9), (2, 3), (4, 6)] {
            assert_eq!(
                sheet.hidden_col_spans(first..=last).collect::<Vec<_>>(),
                brute(&|col| sheet.col_hidden(col), first, last)
            );
        }
        sheet.format.zero_height = true;
        assert_eq!(
            sheet.hidden_row_spans(0..=crate::addr::MAX_ROWS - 1).last(),
            Some((10, crate::addr::MAX_ROWS - 1))
        );
    }

    #[test]
    fn hidden_rows_and_columns_have_no_extent() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.row_heights.insert(1, 0.0);
        sheet.row_heights.insert(2, 15.0);
        sheet.col_widths.insert(3, 0.0);
        assert!(sheet.row_hidden(1));
        assert!(!sheet.row_hidden(2) && !sheet.row_hidden(3));
        assert!(sheet.col_hidden(3) && !sheet.col_hidden(4));
        sheet.format.zero_height = true;
        assert!(sheet.row_hidden(3) && !sheet.row_hidden(2));
    }

    #[test]
    fn workbook_cell_provider() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Data"));
        wb.defined_names.push(DefinedName {
            name: "Answer".into(),
            formula: "A1".into(),
            local_sheet: None,
            hidden: false,
        });
        let a1 = CellRef::parse_a1("A1").unwrap();
        wb.sheet_mut(SheetId(0)).unwrap().set_cell(
            a1,
            Cell {
                value: CellValue::Number { value: 42.0 },
                formula: Some("40+2".into()),
                style: None,
            },
        );

        let id = wb.sheet_id("Data").unwrap();
        assert_eq!(wb.sheet_id("data"), Some(id));
        assert_eq!(wb.value(id, a1), CellValue::Number { value: 42.0 });
        assert_eq!(wb.formula(id, a1), Some("40+2"));
        assert_eq!(
            wb.value(id, CellRef::parse_a1("Z9").unwrap()),
            CellValue::Empty
        );
        assert!(wb.sheet_id("Nope").is_none());
        assert_eq!(
            CellProvider::defined_name(&wb, id, "answer").map(|defined| defined.formula.as_str()),
            Some("A1")
        );
    }

    #[test]
    fn local_defined_name_shadows_workbook_name() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Data"));
        wb.defined_names.extend([
            DefinedName {
                name: "Rate".into(),
                formula: "1".into(),
                local_sheet: None,
                hidden: false,
            },
            DefinedName {
                name: "rate".into(),
                formula: "2".into(),
                local_sheet: Some(SheetId(0)),
                hidden: false,
            },
        ]);

        assert_eq!(
            wb.defined_name(SheetId(0), "RATE")
                .map(|defined| defined.formula.as_str()),
            Some("2")
        );
        assert_eq!(
            wb.defined_name(SheetId(1), "RATE")
                .map(|defined| defined.formula.as_str()),
            Some("1")
        );
    }

    #[test]
    fn iterates_only_cells_in_rectangle() {
        let mut sheet = Sheet::new("Data");
        for address in ["A1", "B2", "C3", "Z100"] {
            sheet.set_cell(
                CellRef::parse_a1(address).unwrap(),
                Cell {
                    value: CellValue::Number { value: 1.0 },
                    ..Cell::default()
                },
            );
        }
        let cells: Vec<_> = sheet
            .iter_cells_in_rect(0..3, 0..2)
            .map(|(cell, _)| cell.to_a1())
            .collect();
        assert_eq!(cells, vec!["A1", "B2"]);
        let mut reversed = 1..2;
        std::mem::swap(&mut reversed.start, &mut reversed.end);
        assert_eq!(sheet.iter_cells_in_rect(0..3, reversed).count(), 0);
    }

    #[test]
    fn hyperlinks_are_addressable_and_extend_the_used_range() {
        let mut sheet = Sheet::new("Data");
        sheet.hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("C4:D5").unwrap(),
            external_target: Some("https://example.com".into()),
            location: None,
            tooltip: None,
            display: Some("Example".into()),
        });

        assert_eq!(sheet.used_range().unwrap().to_a1(), "C4:D5");
        assert_eq!(
            sheet
                .hyperlink_at(CellRef::parse_a1("D5").unwrap())
                .and_then(|link| link.display.as_deref()),
            Some("Example")
        );
        assert!(
            sheet
                .hyperlink_at(CellRef::parse_a1("A1").unwrap())
                .is_none()
        );
    }
}
