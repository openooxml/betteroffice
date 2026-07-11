//! grid geometry and the target-agnostic display list: turns a workbook
//! viewport into draw commands. never imports canvas, dom, or any raster backend.

pub mod display_list;
pub mod geometry;
pub mod region;

use serde::{Deserialize, Serialize};

use xlsx_model::numfmt::{builtin_format_code, format_value};
use xlsx_model::styles::{Border, BorderEdge, BorderStyle, FormatCode, Stylesheet, Theme};
use xlsx_model::value::CellValue;
use xlsx_model::workbook::Sheet;
use xlsx_model::{CellRange, CellRef, Fill, HAlign, SheetId, VAlign, Workbook};

pub use display_list::{Align, DisplayList, DrawCmd, GridMeta, Rect, scaled};
pub use geometry::GridGeometry;
pub use region::{viewport_for_range, viewport_for_used_range};

/// a scrolled window into a sheet, in pixels from the sheet origin.
#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub struct Viewport {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
}

const BACKGROUND_COLOR: &str = "#ffffff";
const GRIDLINE_COLOR: &str = "#d4d4d4";
const TEXT_COLOR: &str = "#000000";
const BORDER_COLOR: &str = "#000000";
const GRIDLINE_WIDTH: f32 = 1.0;
const FONT_SIZE_PT: f32 = 11.0;
const TEXT_PAD_PX: f32 = 2.0;
// rough calibri-like ascent/descent as fractions of the font size.
const ASCENT_RATIO: f32 = 0.7;
const DESCENT_RATIO: f32 = 0.2;

/// build the display list for one viewport of one sheet. commands are emitted
/// background -> fills -> gridlines -> borders -> text, each layer painting over the previous.
pub fn build_display_list(wb: &Workbook, sheet: SheetId, viewport: &Viewport) -> DisplayList {
    let mut commands = Vec::new();

    commands.push(DrawCmd::FillRect {
        x: 0.0,
        y: 0.0,
        w: viewport.width,
        h: viewport.height,
        color: BACKGROUND_COLOR.to_string(),
    });

    let Some(sheet_ref) = wb.sheet(sheet) else {
        return DisplayList {
            width: viewport.width,
            height: viewport.height,
            commands,
            grid: GridMeta::default(),
        };
    };

    let styles = &wb.styles;
    let theme = &styles.theme;
    let geom = GridGeometry::new(sheet_ref);
    let (rows, cols) = geom.viewport_range(viewport);

    let grid = GridMeta {
        start_row: rows.start,
        start_col: cols.start,
        row_offsets: (rows.start..=rows.end)
            .map(|r| geom.row_y(r) - viewport.y)
            .collect(),
        col_offsets: (cols.start..=cols.end)
            .map(|c| geom.col_x(c) - viewport.x)
            .collect(),
    };

    for (at, cell) in visible_anchors(sheet_ref, &rows, &cols) {
        let Some(style) = cell.style else { continue };
        let Some(Fill::Solid(color)) = styles.fill_for(style) else {
            continue;
        };
        let Some(hex) = color.resolve(theme) else {
            continue;
        };
        let (x, y, w, h) = cell_box(&geom, viewport, sheet_ref, at);
        commands.push(DrawCmd::FillRect {
            x,
            y,
            w,
            h,
            color: hex,
        });
    }

    let top = geom.row_y(rows.start) - viewport.y;
    let bottom = geom.row_y(rows.end) - viewport.y;
    let left = geom.col_x(cols.start) - viewport.x;
    let right = geom.col_x(cols.end) - viewport.x;
    for c in cols.start..=cols.end {
        let x = geom.col_x(c) - viewport.x;
        commands.push(DrawCmd::Line {
            x1: x,
            y1: top,
            x2: x,
            y2: bottom,
            width: GRIDLINE_WIDTH,
            color: GRIDLINE_COLOR.to_string(),
            style: None,
        });
    }
    for r in rows.start..=rows.end {
        let y = geom.row_y(r) - viewport.y;
        commands.push(DrawCmd::Line {
            x1: left,
            y1: y,
            x2: right,
            y2: y,
            width: GRIDLINE_WIDTH,
            color: GRIDLINE_COLOR.to_string(),
            style: None,
        });
    }

    for (at, cell) in visible_anchors(sheet_ref, &rows, &cols) {
        let Some(style) = cell.style else { continue };
        let Some(border) = styles.border_for(style) else {
            continue;
        };
        emit_borders(
            &mut commands,
            &geom,
            viewport,
            sheet_ref,
            styles,
            theme,
            at,
            border,
        );
    }

    for (at, cell) in visible_anchors(sheet_ref, &rows, &cols) {
        let Some((text, color)) = cell_display_text(styles, wb.date_system, cell) else {
            continue;
        };

        let (cx0, cy0, cw, ch) = cell_box(&geom, viewport, sheet_ref, at);
        let font = cell.style.and_then(|s| styles.font_for(s));
        let size = font
            .and_then(|f| f.size_pt)
            .map(|p| p as f32)
            .unwrap_or(FONT_SIZE_PT);
        let align = resolve_align(styles, cell);
        let valign = cell
            .style
            .and_then(|s| styles.alignment_for(s))
            .and_then(|a| a.v);

        let tx = match align {
            Align::Left => cx0 + TEXT_PAD_PX,
            Align::Right => cx0 + cw - TEXT_PAD_PX,
            Align::Center => cx0 + cw / 2.0,
        };
        let ty = baseline_y(cy0, ch, size, valign);

        commands.push(DrawCmd::Text {
            x: tx,
            y: ty,
            text,
            font_size: size,
            color,
            clip: Rect {
                x: cx0,
                y: cy0,
                w: cw,
                h: ch,
            },
            align,
            bold: font.is_some_and(|f| f.bold),
            italic: font.is_some_and(|f| f.italic),
            underline: font.is_some_and(|f| f.underline),
            strike: font.is_some_and(|f| f.strike),
            font_family: font.and_then(|f| f.name.clone()),
        });
    }

    DisplayList {
        width: viewport.width,
        height: viewport.height,
        commands,
        grid,
    }
}

/// visible cells that draw: inside the range and not a covered merge cell
/// (only a merge's anchor draws). yields `(anchor, cell)` in row-major order.
fn visible_anchors<'a>(
    sheet: &'a Sheet,
    rows: &'a std::ops::Range<u32>,
    cols: &'a std::ops::Range<u32>,
) -> impl Iterator<Item = (CellRef, &'a xlsx_model::Cell)> {
    sheet.iter_cells().filter(move |(at, _)| {
        if !rows.contains(&at.row) || !cols.contains(&at.col) {
            return false;
        }
        match covering_merge(&sheet.merges, *at) {
            Some(m) => m.start == *at,
            None => true,
        }
    })
}

/// the merge (if any) that covers a cell.
fn covering_merge(merges: &[CellRange], at: CellRef) -> Option<CellRange> {
    merges.iter().copied().find(|m| m.contains(at))
}

/// viewport-local `(x, y, w, h)` of a cell's box, spanning its merged range
/// when `at` anchors one.
fn cell_box(
    geom: &GridGeometry,
    viewport: &Viewport,
    sheet: &Sheet,
    at: CellRef,
) -> (f32, f32, f32, f32) {
    let (end_col, end_row) = match covering_merge(&sheet.merges, at) {
        Some(m) => (m.end.col + 1, m.end.row + 1),
        None => (at.col + 1, at.row + 1),
    };
    let x = geom.col_x(at.col) - viewport.x;
    let y = geom.row_y(at.row) - viewport.y;
    let w = (geom.col_x(end_col) - viewport.x) - x;
    let h = (geom.row_y(end_row) - viewport.y) - y;
    (x, y, w, h)
}

/// display string and resolved font color for a cell, or `None` when it renders
/// nothing. a `[Red]`-style number-format color overrides the font color.
fn cell_display_text(
    styles: &Stylesheet,
    date_system: xlsx_model::DateSystem,
    cell: &xlsx_model::Cell,
) -> Option<(String, String)> {
    if matches!(cell.value, CellValue::Empty) {
        return None;
    }
    let code = format_code_for_cell(styles, cell);
    let formatted = format_value(&cell.value, &code, date_system);
    if formatted.text.is_empty() {
        return None;
    }
    let font = cell.style.and_then(|s| styles.font_for(s));
    let color = formatted
        .color
        .or_else(|| {
            font.and_then(|f| f.color.as_ref())
                .and_then(|c| c.resolve(&styles.theme))
        })
        .unwrap_or_else(|| TEXT_COLOR.to_string());
    Some((formatted.text, color))
}

/// the number-format code a cell's xf resolves to; general when unset or when
/// a builtin id is not modeled.
fn format_code_for_cell(styles: &Stylesheet, cell: &xlsx_model::Cell) -> String {
    match cell.style.map(|s| styles.format_code_for(s)) {
        Some(FormatCode::Custom(c)) => c.to_string(),
        Some(FormatCode::Builtin(id)) => builtin_format_code(id).unwrap_or("General").to_string(),
        None => "General".to_string(),
    }
}

/// the exact string the grid would paint for `cell`, number-format aware.
/// empty cells and formats that yield nothing render as "".
pub fn display_text(
    styles: &Stylesheet,
    date_system: xlsx_model::DateSystem,
    cell: &xlsx_model::Cell,
) -> String {
    if matches!(cell.value, CellValue::Empty) {
        return String::new();
    }
    let code = format_code_for_cell(styles, cell);
    format_value(&cell.value, &code, date_system).text
}

/// horizontal anchor for a cell: an explicit xf alignment wins, otherwise the
/// value type decides (numbers right, booleans center, text/errors left).
fn resolve_align(styles: &Stylesheet, cell: &xlsx_model::Cell) -> Align {
    let type_default = match cell.value {
        CellValue::Number { .. } => Align::Right,
        CellValue::Bool { .. } => Align::Center,
        _ => Align::Left,
    };
    let h = cell
        .style
        .and_then(|s| styles.alignment_for(s))
        .and_then(|a| a.h);
    match h {
        Some(HAlign::Left) | Some(HAlign::Fill) | Some(HAlign::Justify) => Align::Left,
        Some(HAlign::Right) => Align::Right,
        Some(HAlign::Center) | Some(HAlign::CenterContinuous) | Some(HAlign::Distributed) => {
            Align::Center
        }
        Some(HAlign::General) | None => type_default,
    }
}

/// baseline y for a cell's text given its vertical alignment; unset (or center)
/// keeps the centered baseline.
fn baseline_y(cy0: f32, ch: f32, size: f32, valign: Option<VAlign>) -> f32 {
    match valign {
        Some(VAlign::Top) => cy0 + TEXT_PAD_PX + size * ASCENT_RATIO,
        Some(VAlign::Bottom) => cy0 + ch - TEXT_PAD_PX - size * DESCENT_RATIO,
        _ => cy0 + (ch + size * ASCENT_RATIO) / 2.0,
    }
}

/// emit the set edges of a cell's border. a shared interior edge draws once:
/// the bottom (right) edge is skipped when the neighbor declares its own top (left) edge.
#[allow(clippy::too_many_arguments)]
fn emit_borders(
    commands: &mut Vec<DrawCmd>,
    geom: &GridGeometry,
    viewport: &Viewport,
    sheet: &Sheet,
    styles: &Stylesheet,
    theme: &Theme,
    at: CellRef,
    border: &Border,
) {
    let (x, y, w, h) = cell_box(geom, viewport, sheet, at);
    let (x2, y2) = (x + w, y + h);
    let (end_col, end_row) = match covering_merge(&sheet.merges, at) {
        Some(m) => (m.end.col, m.end.row),
        None => (at.col, at.row),
    };

    if let Some(edge) = &border.top {
        commands.push(border_line(x, y, x2, y, edge, theme));
    }
    if let Some(edge) = &border.left {
        commands.push(border_line(x, y, x, y2, edge, theme));
    }
    if let Some(edge) = &border.bottom
        && !neighbor_edge(sheet, styles, end_row + 1, at.col, |b| b.top.is_some())
    {
        commands.push(border_line(x, y2, x2, y2, edge, theme));
    }
    if let Some(edge) = &border.right
        && !neighbor_edge(sheet, styles, at.row, end_col + 1, |b| b.left.is_some())
    {
        commands.push(border_line(x2, y, x2, y2, edge, theme));
    }
}

/// true when the cell at `(row, col)` has a border satisfying `pick`.
fn neighbor_edge(
    sheet: &Sheet,
    styles: &Stylesheet,
    row: u32,
    col: u32,
    pick: impl Fn(&Border) -> bool,
) -> bool {
    sheet
        .cell(CellRef::new(row, col))
        .and_then(|c| c.style)
        .and_then(|s| styles.border_for(s))
        .is_some_and(pick)
}

/// one border edge as a `Line`, mapping the weight to a stroke width and dash
/// style; an unset edge color resolves to black, matching excel's automatic color.
fn border_line(x1: f32, y1: f32, x2: f32, y2: f32, edge: &BorderEdge, theme: &Theme) -> DrawCmd {
    let (width, style) = border_stroke(edge.style);
    let color = edge
        .color
        .as_ref()
        .and_then(|c| c.resolve(theme))
        .unwrap_or_else(|| BORDER_COLOR.to_string());
    DrawCmd::Line {
        x1,
        y1,
        x2,
        y2,
        width,
        color,
        style,
    }
}

/// map a border weight to a `(stroke width, dash style)`.
fn border_stroke(style: BorderStyle) -> (f32, Option<String>) {
    match style {
        BorderStyle::Hair => (1.0, Some("dotted".to_string())),
        BorderStyle::Thin => (1.0, None),
        BorderStyle::Medium => (2.0, None),
        BorderStyle::Thick => (3.0, None),
        BorderStyle::Dashed => (1.0, Some("dashed".to_string())),
        BorderStyle::Dotted => (1.0, Some("dotted".to_string())),
        BorderStyle::Double => (1.0, Some("double".to_string())),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::workbook::{Cell, Sheet};

    fn text_cell(s: &str) -> Cell {
        Cell {
            value: CellValue::Text { value: s.into() },
            ..Cell::default()
        }
    }
    fn num_cell(n: f64) -> Cell {
        Cell {
            value: CellValue::Number { value: n },
            ..Cell::default()
        }
    }

    #[test]
    fn structural_order_and_clip_rect() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 0), num_cell(42.0));
        sheet.set_cell(CellRef::new(0, 1), text_cell("hi"));
        let long = "a very long label that overflows its cell";
        sheet.set_cell(CellRef::new(0, 2), text_cell(long));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 400.0,
            height: 100.0,
        };
        let dl = build_display_list(&wb, SheetId(0), &vp);

        assert_eq!(dl.width, 400.0);
        assert!(matches!(dl.commands[0], DrawCmd::FillRect { .. }));

        let first_text = dl
            .commands
            .iter()
            .position(|c| matches!(c, DrawCmd::Text { .. }));
        let last_line = dl
            .commands
            .iter()
            .rposition(|c| matches!(c, DrawCmd::Line { .. }));
        assert!(first_text.is_some() && last_line.is_some());
        assert!(last_line.unwrap() < first_text.unwrap());

        let texts: Vec<_> = dl
            .commands
            .iter()
            .filter(|c| matches!(c, DrawCmd::Text { .. }))
            .collect();
        assert_eq!(texts.len(), 3);

        let long_text = texts
            .iter()
            .find_map(|c| match c {
                DrawCmd::Text { text, clip, .. } if text == long => Some(clip),
                _ => None,
            })
            .unwrap();
        let dc = geometry::col_chars_to_px(geometry::DEFAULT_COL_WIDTH_CHARS);
        assert_eq!(long_text.x, dc * 2.0);
        assert_eq!(long_text.w, dc);
    }

    #[test]
    fn grid_meta_covers_visible_boundaries() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));
        let dc = geometry::col_chars_to_px(geometry::DEFAULT_COL_WIDTH_CHARS);
        let dr = geometry::row_pt_to_px(geometry::DEFAULT_ROW_HEIGHT_PT);

        let vp = Viewport {
            x: dc * 1.5,
            y: dr * 2.5,
            width: dc * 2.0,
            height: dr * 1.0,
        };
        let dl = build_display_list(&wb, SheetId(0), &vp);

        assert_eq!(dl.grid.start_col, 1);
        assert_eq!(dl.grid.start_row, 2);
        assert_eq!(dl.grid.col_offsets.len(), 4);
        assert_eq!(dl.grid.row_offsets.len(), 3);
        assert!((dl.grid.col_offsets[0] - (dc * 1.0 - vp.x)).abs() < 0.01);
        assert!((dl.grid.row_offsets[0] - (dr * 2.0 - vp.y)).abs() < 0.01);
    }

    #[test]
    fn merge_anchor_draws_covered_cells_skip() {
        let mut sheet = Sheet::new("Sheet1");
        sheet
            .merges
            .push(CellRange::new(CellRef::new(0, 0), CellRef::new(0, 1)));
        sheet.set_cell(CellRef::new(0, 0), text_cell("merged"));
        sheet.set_cell(CellRef::new(0, 1), text_cell("covered"));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 400.0,
            height: 100.0,
        };
        let dl = build_display_list(&wb, SheetId(0), &vp);

        let texts: Vec<_> = dl
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCmd::Text { text, clip, .. } => Some((text.clone(), *clip)),
                _ => None,
            })
            .collect();
        assert_eq!(texts.len(), 1);
        assert_eq!(texts[0].0, "merged");
        let dc = geometry::col_chars_to_px(geometry::DEFAULT_COL_WIDTH_CHARS);
        assert!((texts[0].1.w - dc * 2.0).abs() < 0.01);
    }
}
