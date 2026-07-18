//! sparse workbook containers and the calc-facing cell-access trait.

use std::collections::BTreeMap;
use std::ops::Range;

use crate::addr::{CellRange, CellRef, ColId, RowId, SheetId};
use crate::date::DateSystem;
use crate::styles::Stylesheet;
use crate::value::CellValue;

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Cell {
    pub value: CellValue,
    /// original formula text without the leading `=`, if any.
    pub formula: Option<String>,
    /// index into the workbook style table (cellXfs).
    pub style: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Sheet {
    pub name: String,
    cells: BTreeMap<(RowId, ColId), Cell>,
    pub merges: Vec<CellRange>,
    pub col_widths: BTreeMap<ColId, f64>,
    pub row_heights: BTreeMap<RowId, f64>,
}

impl Sheet {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            ..Self::default()
        }
    }

    pub fn cell(&self, at: CellRef) -> Option<&Cell> {
        self.cells.get(&(at.row, at.col))
    }

    pub fn set_cell(&mut self, at: CellRef, cell: Cell) {
        if cell == Cell::default() {
            self.cells.remove(&(at.row, at.col));
        } else {
            self.cells.insert((at.row, at.col), cell);
        }
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

    pub fn used_range(&self) -> Option<CellRange> {
        let mut it = self.cells.keys();
        let &(r0, c0) = it.next()?;
        let (mut min_r, mut max_r, mut min_c, mut max_c) = (r0, r0, c0, c0);
        for &(r, c) in it {
            min_r = min_r.min(r);
            max_r = max_r.max(r);
            min_c = min_c.min(c);
            max_c = max_c.max(c);
        }
        Some(CellRange::new(
            CellRef::new(min_r, min_c),
            CellRef::new(max_r, max_c),
        ))
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
    pub date_system: DateSystem,
    /// shared string table as parsed; kept for round-trip fidelity.
    pub shared_strings: Vec<String>,
    /// parsed style tables + theme; a cell's `style` indexes `styles.cell_xfs`.
    pub styles: Stylesheet,
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
}

/// read access the calc engine evaluates through.
pub trait CellProvider {
    fn value(&self, sheet: SheetId, at: CellRef) -> CellValue;
    fn formula(&self, sheet: SheetId, at: CellRef) -> Option<&str>;
    fn sheet_id(&self, name: &str) -> Option<SheetId>;
}

impl CellProvider for Workbook {
    fn value(&self, sheet: SheetId, at: CellRef) -> CellValue {
        self.sheet(sheet)
            .and_then(|s| s.cell(at))
            .map(|c| c.value.clone())
            .unwrap_or_default()
    }

    fn formula(&self, sheet: SheetId, at: CellRef) -> Option<&str> {
        self.sheet(sheet)?.cell(at)?.formula.as_deref()
    }

    fn sheet_id(&self, name: &str) -> Option<SheetId> {
        self.sheet_by_name(name).map(|(id, _)| id)
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
    fn workbook_cell_provider() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Data"));
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
}
