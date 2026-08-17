//! range -> viewport helpers: turn a cell range (or a sheet's used range) into
//! the content-pixel rectangle the display list is built for.

use ooxml_drawingml::chart::PlotRect;
use xlsx_model::workbook::Sheet;
use xlsx_model::{CellRange, CellRef};

use crate::Viewport;
use crate::chart::resolve_chart_anchor;
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
    viewport_for_used_range_within(sheet, |_| true)
}

/// [`viewport_for_used_range`], but a chart only widens the frame while `fits`
/// accepts the widened rectangle. A chart parked far below the data would
/// otherwise inflate the frame past a renderer's limits and take the whole
/// render down with it; declining it leaves the chart out of frame instead,
/// exactly as a chart past the edge of an explicit viewport already is.
///
/// Charts are offered in sheet order, so of two that fit alone but not together
/// the earlier one wins. A used range `fits` already rejects grows no further.
pub fn viewport_for_used_range_within(
    sheet: &Sheet,
    mut fits: impl FnMut(&Viewport) -> bool,
) -> Viewport {
    let range = sheet
        .used_range()
        .unwrap_or_else(|| CellRange::new(CellRef::new(0, 0), CellRef::new(49, 25)));
    let mut viewport = viewport_for_range(sheet, range);
    let geometry = GridGeometry::new(sheet);
    let (frozen_rows, frozen_cols) = sheet
        .freeze_pane
        .map_or((0, 0), |pane| (pane.rows, pane.cols));
    for chart in &sheet.charts {
        let Ok(anchor) = resolve_chart_anchor(chart.anchor, &geometry, frozen_rows, frozen_cols)
        else {
            continue;
        };
        let grown = grown_to(viewport, anchor.rect);
        if fits(&grown) {
            viewport = grown;
        }
    }
    viewport
}

/// `viewport` widened to take in a chart's content-pixel rectangle.
fn grown_to(viewport: Viewport, rect: PlotRect) -> Viewport {
    let viewport_right = viewport.x + viewport.width;
    let viewport_bottom = viewport.y + viewport.height;
    let left = viewport.x.min(rect.x as f32);
    let top = viewport.y.min(rect.y as f32);
    Viewport {
        x: left,
        y: top,
        width: viewport_right.max((rect.x + rect.w) as f32) - left,
        height: viewport_bottom.max((rect.y + rect.h) as f32) - top,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::geometry::{
        DEFAULT_COL_WIDTH_CHARS, DEFAULT_ROW_HEIGHT_PT, col_chars_to_px, row_pt_to_px,
    };
    use xlsx_model::CellValue;
    use xlsx_model::chart::{AnchorCell, AnchorEditAs, ChartAnchor, SheetChart};
    use xlsx_model::workbook::Cell;

    fn text_cell(s: &str) -> Cell {
        Cell {
            value: CellValue::Text { value: s.into() },
            ..Cell::default()
        }
    }

    /// a 15-row chart hanging off column h, `row` rows down the sheet.
    fn chart_at_row(row: u32) -> SheetChart {
        SheetChart {
            part: "xl/charts/chart1.xml".into(),
            drawing: "xl/drawings/drawing1.xml".into(),
            anchor_index: 0,
            anchor: ChartAnchor::TwoCell {
                from: AnchorCell {
                    col: 7,
                    row,
                    ..AnchorCell::default()
                },
                to: AnchorCell {
                    col: 14,
                    row: row + 15,
                    ..AnchorCell::default()
                },
                edit_as: AnchorEditAs::OneCell,
            },
            refs: Vec::new(),
        }
    }

    /// a sheet used out to f11, with one chart.
    fn sheet_with_chart(row: u32) -> Sheet {
        let mut sheet = Sheet::new("S");
        sheet.set_cell(CellRef::parse_a1("A1").unwrap(), text_cell("a"));
        sheet.set_cell(CellRef::parse_a1("F11").unwrap(), text_cell("b"));
        sheet.charts.push(chart_at_row(row));
        sheet
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

    #[test]
    fn a_chart_grows_the_used_range() {
        let sheet = sheet_with_chart(1);
        let dc = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        let dr = row_pt_to_px(DEFAULT_ROW_HEIGHT_PT);
        let vp = viewport_for_used_range(&sheet);
        assert_eq!((vp.x, vp.y), (0.0, 0.0));
        assert!((vp.width - dc * 14.0).abs() < 0.01);
        assert!((vp.height - dr * 16.0).abs() < 0.01);
    }

    #[test]
    fn a_chart_the_budget_allows_still_grows_the_used_range() {
        let sheet = sheet_with_chart(1);
        let dc = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        let dr = row_pt_to_px(DEFAULT_ROW_HEIGHT_PT);
        let vp = viewport_for_used_range_within(&sheet, |grown| grown.height <= 16_384.0);
        assert!((vp.width - dc * 14.0).abs() < 0.01);
        assert!((vp.height - dr * 16.0).abs() < 0.01);
    }

    #[test]
    fn a_chart_past_the_budget_is_left_out_of_frame() {
        let sheet = sheet_with_chart(2_000);
        let used = viewport_for_range(&sheet, CellRange::parse_a1("A1:F11").unwrap());
        let vp = viewport_for_used_range_within(&sheet, |grown| grown.height <= 16_384.0);
        assert_eq!(vp, used);
    }

    #[test]
    fn a_used_range_the_budget_already_refuses_grows_no_further() {
        let sheet = sheet_with_chart(1);
        let used = viewport_for_range(&sheet, CellRange::parse_a1("A1:F11").unwrap());
        assert_eq!(viewport_for_used_range_within(&sheet, |_| false), used);
    }

    #[test]
    fn a_reachable_chart_survives_an_unreachable_one() {
        let mut sheet = sheet_with_chart(2_000);
        sheet.charts.push(chart_at_row(1));
        let dc = col_chars_to_px(DEFAULT_COL_WIDTH_CHARS);
        let dr = row_pt_to_px(DEFAULT_ROW_HEIGHT_PT);
        let vp = viewport_for_used_range_within(&sheet, |grown| grown.height <= 16_384.0);
        assert!((vp.width - dc * 14.0).abs() < 0.01);
        assert!((vp.height - dr * 16.0).abs() < 0.01);
    }
}
