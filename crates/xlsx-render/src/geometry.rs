//! grid geometry: cumulative pixel offsets for columns and rows. tables cover
//! only the explicitly-sized prefix; past it, default sizes are extrapolated analytically.

use std::ops::Range;

use xlsx_model::workbook::Sheet;
use xlsx_model::{ColId, RowId};

use crate::Viewport;

/// default column width in characters of max-digit-width (excel default).
pub const DEFAULT_COL_WIDTH_CHARS: f64 = 8.43;
/// default row height in points (excel default for 11pt calibri).
pub const DEFAULT_ROW_HEIGHT_PT: f64 = 15.0;
/// max-digit-width of the default font (calibri 11) in pixels at 96dpi.
pub const MAX_DIGIT_WIDTH_PX: f64 = 7.0;
/// css/screen pixels per point at 96dpi (96/72).
pub const PX_PER_PT: f64 = 96.0 / 72.0;

/// column width in characters -> pixels per ecma-376 §18.3.1.13: the stored
/// width folds in 5px of padding; reverse it and snap to excel's 1/256 grid.
pub fn col_chars_to_px(chars: f64) -> f32 {
    let mdw = MAX_DIGIT_WIDTH_PX;
    (((chars * mdw + 5.0) / mdw * 256.0).round() / 256.0 * mdw) as f32
}

/// row height in points -> pixels at 96dpi.
pub fn row_pt_to_px(pt: f64) -> f32 {
    (pt * PX_PER_PT) as f32
}

/// cumulative left-edge x offsets for columns and top-edge y offsets for rows,
/// both in pixels from the sheet origin.
#[derive(Debug, Clone)]
pub struct GridGeometry {
    /// `col_x[i]` is the left edge of column `i`; len is `n_cols + 1`.
    col_x: Vec<f32>,
    row_y: Vec<f32>,
    n_cols: ColId,
    n_rows: RowId,
    default_col_px: f32,
    default_row_px: f32,
}

impl GridGeometry {
    /// build cumulative offset tables from a sheet's custom widths/heights.
    pub fn new(sheet: &Sheet) -> Self {
        let default_col_px = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        let default_row_px = row_pt_to_px(DEFAULT_ROW_HEIGHT_PT);

        let n_cols = sheet
            .col_widths
            .keys()
            .next_back()
            .map(|&c| c + 1)
            .unwrap_or(0);
        let n_rows = sheet
            .row_heights
            .keys()
            .next_back()
            .map(|&r| r + 1)
            .unwrap_or(0);

        let mut col_x = Vec::with_capacity(n_cols as usize + 1);
        col_x.push(0.0);
        for c in 0..n_cols {
            let w = sheet
                .col_widths
                .get(&c)
                .map(|&w| col_chars_to_px(w))
                .unwrap_or(default_col_px);
            col_x.push(col_x[c as usize] + w);
        }

        let mut row_y = Vec::with_capacity(n_rows as usize + 1);
        row_y.push(0.0);
        for r in 0..n_rows {
            let h = sheet
                .row_heights
                .get(&r)
                .map(|&h| row_pt_to_px(h))
                .unwrap_or(default_row_px);
            row_y.push(row_y[r as usize] + h);
        }

        Self {
            col_x,
            row_y,
            n_cols,
            n_rows,
            default_col_px,
            default_row_px,
        }
    }

    /// left edge of a column in pixels; extrapolates past the sized prefix.
    pub fn col_x(&self, col: ColId) -> f32 {
        if (col as usize) < self.col_x.len() {
            self.col_x[col as usize]
        } else {
            let last = *self.col_x.last().unwrap();
            last + (col - self.n_cols) as f32 * self.default_col_px
        }
    }

    /// top edge of a row in pixels; extrapolates past the sized prefix.
    pub fn row_y(&self, row: RowId) -> f32 {
        if (row as usize) < self.row_y.len() {
            self.row_y[row as usize]
        } else {
            let last = *self.row_y.last().unwrap();
            last + (row - self.n_rows) as f32 * self.default_row_px
        }
    }

    /// column whose span contains `x`; clamps negatives to column 0.
    pub fn col_at_x(&self, x: f32) -> ColId {
        let last_edge = *self.col_x.last().unwrap();
        if x < last_edge {
            let idx = self.col_x.partition_point(|&edge| edge <= x);
            idx.saturating_sub(1) as ColId
        } else {
            let extra = ((x - last_edge) / self.default_col_px).max(0.0) as ColId;
            self.n_cols + extra
        }
    }

    /// row whose span contains `y`; clamps negatives to row 0.
    pub fn row_at_y(&self, y: f32) -> RowId {
        let last_edge = *self.row_y.last().unwrap();
        if y < last_edge {
            let idx = self.row_y.partition_point(|&edge| edge <= y);
            idx.saturating_sub(1) as RowId
        } else {
            let extra = ((y - last_edge) / self.default_row_px).max(0.0) as RowId;
            self.n_rows + extra
        }
    }

    /// half-open (row, col) ranges of cells intersecting the viewport.
    pub fn viewport_range(&self, vp: &Viewport) -> (Range<RowId>, Range<ColId>) {
        let r0 = self.row_at_y(vp.y);
        let r1 = self.row_at_y(vp.y + vp.height);
        let c0 = self.col_at_x(vp.x);
        let c1 = self.col_at_x(vp.x + vp.width);
        (r0..r1 + 1, c0..c1 + 1)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::BTreeMap;

    fn sheet_with(cols: &[(ColId, f64)], rows: &[(RowId, f64)]) -> Sheet {
        let mut s = Sheet::new("S");
        s.col_widths = cols.iter().copied().collect::<BTreeMap<_, _>>();
        s.row_heights = rows.iter().copied().collect::<BTreeMap<_, _>>();
        s
    }

    #[test]
    fn unit_conversion_anchors() {
        let px = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        assert!((px - 64.0).abs() < 0.1, "default col width was {px}");
        assert_eq!(row_pt_to_px(DEFAULT_ROW_HEIGHT_PT), 20.0);
        assert!(col_chars_to_px(0.0) > 0.0);
    }

    #[test]
    fn all_default_grid_extrapolates() {
        let g = GridGeometry::new(&Sheet::new("S"));
        let dc = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        let dr = row_pt_to_px(DEFAULT_ROW_HEIGHT_PT);
        assert_eq!(g.col_x(0), 0.0);
        assert_eq!(g.col_x(3), 3.0 * dc);
        assert_eq!(g.row_y(10), 10.0 * dr);
        assert_eq!(g.col_at_x(dc * 2.5), 2);
        assert_eq!(g.row_at_y(dr * 4.0), 4);
        assert_eq!(g.col_at_x(-5.0), 0);
    }

    #[test]
    fn custom_widths_and_binary_search_edges() {
        let g = GridGeometry::new(&sheet_with(&[(0, 20.0), (2, 4.0)], &[(0, 30.0)]));
        let w0 = col_chars_to_px(20.0);
        let w1 = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        let w2 = col_chars_to_px(4.0);
        assert_eq!(g.col_x(0), 0.0);
        assert_eq!(g.col_x(1), w0);
        assert_eq!(g.col_x(2), w0 + w1);
        assert_eq!(g.col_x(3), w0 + w1 + w2);

        assert_eq!(g.col_at_x(0.0), 0);
        assert_eq!(g.col_at_x(w0), 1);
        assert_eq!(g.col_at_x(w0 + w1), 2);
        assert_eq!(g.col_at_x(w0 - 0.01), 0);
        assert_eq!(g.col_at_x(w0 + w1 + w2), 3);

        assert_eq!(g.row_y(1), row_pt_to_px(30.0));
    }

    #[test]
    fn viewport_range_covers_visible_cells() {
        let g = GridGeometry::new(&Sheet::new("S"));
        let dc = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        let dr = row_pt_to_px(DEFAULT_ROW_HEIGHT_PT);
        let vp = Viewport {
            x: dc * 1.5,
            y: dr * 2.5,
            width: dc * 2.0,
            height: dr * 1.0,
        };
        let (rows, cols) = g.viewport_range(&vp);
        assert_eq!(cols, 1..4);
        assert_eq!(rows, 2..4);
    }
}
