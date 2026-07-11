//! range -> viewport helpers: turn a cell range (or a sheet's used range) into
//! the content-pixel rectangle the display list is built for.

use xlsx_model::workbook::Sheet;
use xlsx_model::{CellRange, CellRef};

use crate::Viewport;
use crate::geometry::GridGeometry;

/// content-pixel rectangle spanning an inclusive cell range, with the viewport
/// origin at the range's top-left.
pub fn viewport_for_range(sheet: &Sheet, range: CellRange) -> Viewport {
    let geom = GridGeometry::new(sheet);
    let x = geom.col_x(range.start.col);
    let y = geom.row_y(range.start.row);
    let right = geom.col_x(range.end.col + 1);
    let bottom = geom.row_y(range.end.row + 1);
    Viewport {
        x,
        y,
        width: right - x,
        height: bottom - y,
    }
}

/// content-pixel rectangle spanning the sheet's whole used range; an empty
/// sheet falls back to a1:z50, matching `xlsx-wasm`'s `sheet_info` default extent.
pub fn viewport_for_used_range(sheet: &Sheet) -> Viewport {
    let range = sheet
        .used_range()
        .unwrap_or_else(|| CellRange::new(CellRef::new(0, 0), CellRef::new(49, 25)));
    viewport_for_range(sheet, range)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::geometry::{
        DEFAULT_COL_WIDTH_CHARS, DEFAULT_ROW_HEIGHT_PT, col_chars_to_px, row_pt_to_px,
    };
    use xlsx_model::CellValue;
    use xlsx_model::workbook::Cell;

    fn text_cell(s: &str) -> Cell {
        Cell {
            value: CellValue::Text { value: s.into() },
            ..Cell::default()
        }
    }

    #[test]
    fn range_viewport_spans_all_default_cells() {
        let sheet = Sheet::new("S");
        let dc = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        let dr = row_pt_to_px(DEFAULT_ROW_HEIGHT_PT);

        let vp = viewport_for_range(&sheet, CellRange::parse_a1("B2:C4").unwrap());
        assert_eq!(vp.x, dc);
        assert_eq!(vp.y, dr);
        assert!((vp.width - dc * 2.0).abs() < 0.01);
        assert!((vp.height - dr * 3.0).abs() < 0.01);
    }

    #[test]
    fn single_cell_range_is_one_cell() {
        let sheet = Sheet::new("S");
        let dc = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        let dr = row_pt_to_px(DEFAULT_ROW_HEIGHT_PT);
        let vp = viewport_for_range(&sheet, CellRange::parse_a1("A1").unwrap());
        assert_eq!((vp.x, vp.y), (0.0, 0.0));
        assert!((vp.width - dc).abs() < 0.01);
        assert!((vp.height - dr).abs() < 0.01);
    }

    #[test]
    fn used_range_tracks_populated_cells() {
        let mut sheet = Sheet::new("S");
        sheet.set_cell(CellRef::parse_a1("B2").unwrap(), text_cell("a"));
        sheet.set_cell(CellRef::parse_a1("D5").unwrap(), text_cell("b"));
        let dc = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        let dr = row_pt_to_px(DEFAULT_ROW_HEIGHT_PT);

        let vp = viewport_for_used_range(&sheet);
        assert_eq!(vp.x, dc);
        assert_eq!(vp.y, dr);
        assert!((vp.width - dc * 3.0).abs() < 0.01);
        assert!((vp.height - dr * 4.0).abs() < 0.01);
    }

    #[test]
    fn empty_sheet_falls_back_to_a1_z50() {
        let sheet = Sheet::new("S");
        let dc = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        let dr = row_pt_to_px(DEFAULT_ROW_HEIGHT_PT);
        let vp = viewport_for_used_range(&sheet);
        assert_eq!((vp.x, vp.y), (0.0, 0.0));
        assert!((vp.width - dc * 26.0).abs() < 0.01);
        assert!((vp.height - dr * 50.0).abs() < 0.01);
    }
}
