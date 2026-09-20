//! grid geometry: cumulative pixel offsets for columns and rows. tables cover
//! only the explicitly-sized prefix; past it, default sizes are extrapolated analytically.

use std::collections::{BTreeMap, HashSet};
use std::ops::Range;

use xlsx_model::styles::Stylesheet;
use xlsx_model::workbook::Sheet;
use xlsx_model::{ColId, RowId};

use crate::Viewport;

/// default column width in characters of max-digit-width (excel default).
pub const DEFAULT_COL_WIDTH_CHARS: f64 = 8.43;
/// default row height in points (excel default for 11pt calibri).
pub const DEFAULT_ROW_HEIGHT_PT: f64 = 15.0;
/// point size assumed for a cell whose style names none.
const DEFAULT_FONT_SIZE_PT: f64 = 11.0;
/// max-digit-width of the default font (calibri 11) in pixels at 96dpi.
pub const MAX_DIGIT_WIDTH_PX: f64 = 7.0;
/// css/screen pixels per point at 96dpi (96/72).
pub const PX_PER_PT: f64 = 96.0 / 72.0;
pub const EMU_PER_PT: f64 = 12_700.0;

/// column width in characters -> pixels per ecma-376 §18.3.1.13: the stored
/// width folds in 5px of padding; reverse it and snap to excel's 1/256 grid.
pub fn col_chars_to_px(chars: f64) -> f32 {
    if chars == 0.0 {
        return 0.0;
    }
    let mdw = MAX_DIGIT_WIDTH_PX;
    (((chars * mdw + 5.0) / mdw * 256.0).round() / 256.0 * mdw) as f32
}

/// row height in points -> pixels at 96dpi.
pub fn row_pt_to_px(pt: f64) -> f32 {
    (pt * PX_PER_PT) as f32
}

/// height in points excel gives a row auto-fitted to `size_pt` text, on
/// excel's quarter-point grid. the constants are fitted to the row bands in
/// excel's own renders of the fidelity corpus.
pub fn autofit_row_height_pt(size_pt: f64) -> f64 {
    (((size_pt * AUTOFIT_LINE_RATIO - AUTOFIT_LEADING_PT) * 4.0).round() / 4.0).max(0.0)
}

const AUTOFIT_LINE_RATIO: f64 = 1.4615;
const AUTOFIT_LEADING_PT: f64 = 2.0769;

/// floors a rescaled height to whole points, absorbing the float error a ratio
/// of quarter-point heights leaves just under an exact integer (75.0 x 11/15
/// evaluates to 54.999...).
fn floor_pt(pt: f64) -> f64 {
    (pt + 1e-9).floor()
}

/// a stored `ht` is expressed against the sheet's `defaultRowHeight`; when that
/// default is a cached hint excel recomputes, every stored height is rescaled
/// with it and floored to whole points. the ratio needs excel's own fitted
/// height for the normal font, and [`autofit_row_height_pt`] is measured on
/// calibri, so a sheet whose normal font names another face keeps its heights.
fn stored_height_scale(sheet: &Sheet, styles: &Stylesheet) -> Option<f64> {
    if sheet.format.custom_height {
        return None;
    }
    let declared = sheet.format.default_row_height_pt.filter(|&pt| pt > 0.0)?;
    let normal = styles.fonts.first();
    if !normal
        .and_then(|font| font.name.as_deref())
        .is_none_or(|name| name.eq_ignore_ascii_case("calibri"))
    {
        return None;
    }
    let fitted = autofit_row_height_pt(
        normal
            .and_then(|font| font.size_pt)
            .unwrap_or(DEFAULT_FONT_SIZE_PT),
    );
    (fitted > 0.0 && fitted != declared).then_some(fitted / declared)
}

/// rows excel auto-fits, with the height each takes: every row that carries no
/// `ht` and is not pinned by `sheetFormatPr/@customHeight`. a declared
/// `defaultRowHeight` is a cached hint excel recomputes, so a row with content
/// takes its fitted height outright; without one the sheet default is already
/// the normal font's fitted height and only taller content grows a row.
fn autofit_rows(sheet: &Sheet, styles: &Stylesheet, default_pt: f64) -> BTreeMap<RowId, f64> {
    let mut fitted = BTreeMap::new();
    if sheet.format.custom_height {
        return fitted;
    }
    let cached_default = sheet.format.default_row_height_pt.is_some();
    let spanned: HashSet<(RowId, ColId)> = sheet
        .merges
        .iter()
        .filter(|range| range.end.row > range.start.row)
        .map(|range| (range.start.row, range.start.col))
        .collect();
    for (at, cell) in sheet.iter_cells() {
        if sheet.row_heights.contains_key(&at.row) || spanned.contains(&(at.row, at.col)) {
            continue;
        }
        let size = cell
            .style
            .and_then(|style| styles.font_for(style))
            .and_then(|font| font.size_pt)
            .unwrap_or(DEFAULT_FONT_SIZE_PT);
        let height = autofit_row_height_pt(size);
        if !cached_default && height <= default_pt {
            continue;
        }
        let entry = fitted.entry(at.row).or_insert(height);
        if height > *entry {
            *entry = height;
        }
    }
    fitted
}

/// EMU -> pixels in the renderer's 96dpi coordinate space.
pub fn emu_to_px(emu: i64) -> f64 {
    emu as f64 / EMU_PER_PT * PX_PER_PT
}

#[derive(Debug, Clone, serde::Deserialize, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PrintMetrics {
    pub dpi: f32,
    pub max_digit_width: f32,
    pub default_row_height_pt: f32,
    pub default_column_width: Option<f64>,
    pub font_size_pt: f32,
    pub font_family: String,
    pub font_ascent: f32,
    pub font_descent: f32,
}

impl PrintMetrics {
    pub fn is_valid(&self) -> bool {
        !self.font_family.is_empty()
            && self.font_family.len() <= 1024
            && (36.0..=600.0).contains(&self.dpi)
            && (1.0..=256.0).contains(&self.max_digit_width)
            && (1.0..=409.0).contains(&self.default_row_height_pt)
            && self
                .default_column_width
                .is_none_or(|w| (0.0..=255.0).contains(&w))
            && (1.0..=409.0).contains(&self.font_size_pt)
            && (0.0..=4096.0).contains(&self.font_ascent)
            && (0.0..=4096.0).contains(&self.font_descent)
            && self.font_ascent + self.font_descent > 0.0
    }

    pub fn column_pixels(&self, width: f64) -> f32 {
        let digit = self.max_digit_width as f64;
        (((256.0 * width + (128.0 / digit).floor()) / 256.0 * digit).floor() * 96.0
            / self.dpi as f64) as f32
    }
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
    pub fn new(sheet: &Sheet, styles: &Stylesheet) -> Self {
        Self::with_sizes(
            sheet,
            styles,
            col_chars_to_px(DEFAULT_COL_WIDTH_CHARS),
            sheet
                .format
                .default_row_height_pt
                .unwrap_or(DEFAULT_ROW_HEIGHT_PT),
            col_chars_to_px,
        )
    }

    pub fn for_print(sheet: &Sheet, styles: &Stylesheet, metrics: &PrintMetrics) -> Self {
        Self::with_sizes(
            sheet,
            styles,
            metrics
                .default_column_width
                .map_or(col_chars_to_px(DEFAULT_COL_WIDTH_CHARS), |width| {
                    metrics.column_pixels(width)
                }),
            metrics.default_row_height_pt as f64,
            |width| metrics.column_pixels(width),
        )
    }

    fn with_sizes(
        sheet: &Sheet,
        styles: &Stylesheet,
        default_col_px: f32,
        default_row_pt: f64,
        column_pixels: impl Fn(f64) -> f32,
    ) -> Self {
        let default_row_px = row_pt_to_px(default_row_pt);
        let fitted = autofit_rows(sheet, styles, default_row_pt);
        let scale = stored_height_scale(sheet, styles);
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
            .into_iter()
            .chain(fitted.keys().next_back())
            .max()
            .map(|&r| r + 1)
            .unwrap_or(0);

        let mut col_x = Vec::with_capacity(n_cols as usize + 1);
        col_x.push(0.0);
        for c in 0..n_cols {
            let w = sheet
                .col_widths
                .get(&c)
                .map(|&w| column_pixels(w))
                .unwrap_or(default_col_px);
            let start = col_x.last().copied().unwrap_or(0.0);
            col_x.push(start + w);
        }

        let mut row_y = Vec::with_capacity(n_rows as usize + 1);
        row_y.push(0.0);
        for r in 0..n_rows {
            let h = sheet
                .row_heights
                .get(&r)
                .map(|&h| row_pt_to_px(scale.map_or(h, |s| floor_pt(h * s))))
                .or_else(|| fitted.get(&r).map(|&h| row_pt_to_px(h)))
                .unwrap_or(default_row_px);
            let start = row_y.last().copied().unwrap_or(0.0);
            row_y.push(start + h);
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
        self.col_x.get(col as usize).copied().unwrap_or_else(|| {
            let last = self.col_x.last().copied().unwrap_or(0.0);
            last + col.saturating_sub(self.n_cols) as f32 * self.default_col_px
        })
    }

    /// top edge of a row in pixels; extrapolates past the sized prefix.
    pub fn row_y(&self, row: RowId) -> f32 {
        self.row_y.get(row as usize).copied().unwrap_or_else(|| {
            let last = self.row_y.last().copied().unwrap_or(0.0);
            last + row.saturating_sub(self.n_rows) as f32 * self.default_row_px
        })
    }

    /// column whose span contains `x`; clamps negatives to column 0.
    pub fn col_at_x(&self, x: f32) -> ColId {
        let last_edge = self.col_x.last().copied().unwrap_or(0.0);
        if x < last_edge {
            let idx = self.col_x.partition_point(|&edge| edge <= x);
            idx.saturating_sub(1) as ColId
        } else if self.default_col_px == 0.0 {
            self.n_cols
        } else {
            let extra = ((x - last_edge) / self.default_col_px).max(0.0) as ColId;
            self.n_cols + extra
        }
    }

    /// row whose span contains `y`; clamps negatives to row 0.
    pub fn row_at_y(&self, y: f32) -> RowId {
        let last_edge = self.row_y.last().copied().unwrap_or(0.0);
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
    use xlsx_model::CellRef;
    use xlsx_model::styles::{Font, Xf};

    fn styles() -> Stylesheet {
        Stylesheet::default()
    }

    fn sheet_with(cols: &[(ColId, f64)], rows: &[(RowId, f64)]) -> Sheet {
        let mut s = Sheet::new("S");
        s.col_widths = cols.iter().copied().collect::<BTreeMap<_, _>>();
        s.row_heights = rows.iter().copied().collect::<BTreeMap<_, _>>();
        s
    }

    #[test]
    fn autofit_matches_the_row_bands_excel_renders() {
        assert_eq!(autofit_row_height_pt(11.0), 14.0);
        assert_eq!(autofit_row_height_pt(12.0), 15.5);
        assert_eq!(autofit_row_height_pt(24.0), 33.0);
        assert_eq!(autofit_row_height_pt(30.0), 41.75);
        assert_eq!(autofit_row_height_pt(0.0), 0.0);
    }

    #[test]
    fn an_unsized_row_takes_its_tallest_font_unless_the_sheet_pins_it() {
        let mut styles = Stylesheet::default();
        styles.fonts.push(Font {
            size_pt: Some(30.0),
            ..Font::default()
        });
        styles.cell_xfs.push(Xf {
            font: Some(0),
            ..Xf::default()
        });
        let mut sheet = Sheet::new("S");
        sheet.set_cell(
            CellRef::new(1, 0),
            xlsx_model::workbook::Cell {
                style: Some(0),
                ..Default::default()
            },
        );
        let default_px = row_pt_to_px(DEFAULT_ROW_HEIGHT_PT);
        let grown = GridGeometry::new(&sheet, &styles);
        assert_eq!(grown.row_y(1), default_px);
        assert!(
            (grown.row_y(2) - grown.row_y(1) - row_pt_to_px(autofit_row_height_pt(30.0))).abs()
                < 0.001
        );

        sheet.format.custom_height = true;
        let pinned = GridGeometry::new(&sheet, &styles);
        assert_eq!(pinned.row_y(2), default_px * 2.0);

        sheet.format.default_row_height_pt = Some(DEFAULT_ROW_HEIGHT_PT);
        let cached = GridGeometry::new(&sheet, &styles);
        assert_eq!(cached.row_y(2), default_px * 2.0);

        sheet.format.custom_height = false;
        let recomputed = GridGeometry::new(&sheet, &styles);
        assert!(
            (recomputed.row_y(2) - recomputed.row_y(1) - row_pt_to_px(autofit_row_height_pt(30.0)))
                .abs()
                < 0.001
        );

        sheet.format.default_row_height_pt = None;
        sheet.row_heights.insert(1, 9.0);
        let authored = GridGeometry::new(&sheet, &styles);
        assert!((authored.row_y(2) - authored.row_y(1) - row_pt_to_px(9.0)).abs() < 0.001);
    }

    #[test]
    fn stored_heights_rescale_with_a_recomputed_cached_default() {
        let mut styles = Stylesheet::default();
        styles.fonts.push(Font {
            name: Some("Calibri".into()),
            size_pt: Some(11.0),
            ..Font::default()
        });
        let mut sheet = sheet_with(&[], &[(0, 15.0), (1, 24.0), (2, 37.5)]);
        sheet.format.default_row_height_pt = Some(15.0);
        let rescaled = GridGeometry::new(&sheet, &styles);
        for (row, pt) in [(0, 14.0), (1, 22.0), (2, 35.0)] {
            assert!(
                (rescaled.row_y(row + 1) - rescaled.row_y(row) - row_pt_to_px(pt)).abs() < 0.001
            );
        }

        sheet.format.custom_height = true;
        let pinned = GridGeometry::new(&sheet, &styles);
        assert!((pinned.row_y(1) - row_pt_to_px(15.0)).abs() < 0.001);

        sheet.format.custom_height = false;
        styles.fonts[0].name = Some("Arial".into());
        let unmeasured = GridGeometry::new(&sheet, &styles);
        assert!((unmeasured.row_y(1) - row_pt_to_px(15.0)).abs() < 0.001);
    }

    #[test]
    fn a_rescaled_height_landing_on_an_integer_keeps_it() {
        let mut styles = Stylesheet::default();
        styles.fonts.push(Font {
            name: Some("Calibri".into()),
            size_pt: Some(11.0),
            ..Font::default()
        });
        let mut sheet = sheet_with(&[], &[(0, 90.0)]);
        sheet.format.default_row_height_pt = Some(20.0);
        let grid = GridGeometry::new(&sheet, &styles);
        assert!((grid.row_y(1) - row_pt_to_px(63.0)).abs() < 0.001);
    }

    #[test]
    fn unit_conversion_anchors() {
        let px = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        assert!((px - 64.0).abs() < 0.1, "default col width was {px}");
        assert_eq!(row_pt_to_px(DEFAULT_ROW_HEIGHT_PT), 20.0);
        assert_eq!(col_chars_to_px(0.0), 0.0);
        assert_eq!(emu_to_px(9_525), 1.0);
    }

    #[test]
    fn all_default_grid_extrapolates() {
        let g = GridGeometry::new(&Sheet::new("S"), &styles());
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
        let g = GridGeometry::new(&sheet_with(&[(0, 20.0), (2, 4.0)], &[(0, 30.0)]), &styles());
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
        let g = GridGeometry::new(&Sheet::new("S"), &styles());
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

    #[test]
    fn hidden_default_columns_preserve_visible_spans_and_end_at_the_sized_prefix() {
        let g = GridGeometry::with_sizes(
            &sheet_with(&[(1, 12.0), (3, 10.0), (4, 0.0)], &[]),
            &styles(),
            0.0,
            DEFAULT_ROW_HEIGHT_PT,
            col_chars_to_px,
        );
        let first_width = col_chars_to_px(12.0);
        let total_width = first_width + col_chars_to_px(10.0);
        assert_eq!(g.col_x(1), 0.0);
        assert_eq!(g.col_x(2), first_width);
        assert_eq!(g.col_x(3), first_width);
        assert_eq!(g.col_x(16_384), total_width);
        assert_eq!(g.col_at_x(-1.0), 0);
        assert_eq!(g.col_at_x(0.0), 1);
        assert_eq!(g.col_at_x(first_width), 3);
        assert_eq!(g.col_at_x(total_width - 0.01), 3);
        assert_eq!(g.col_at_x(total_width), 5);
        assert_eq!(g.col_at_x(total_width + 100.0), 5);
        assert_eq!(
            g.viewport_range(&Viewport {
                x: 0.0,
                y: 0.0,
                width: total_width + 100.0,
                height: 20.0,
            })
            .1,
            1..6
        );
    }
}
