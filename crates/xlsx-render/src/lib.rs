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

// ghost pair colors, matching the docx revision palette (del struck / ins).
const GHOST_DEL_COLOR: &str = "#c62828";
const GHOST_INS_COLOR: &str = "#2e7d32";
const GHOST_DEL_HIGHLIGHT: &str = "#c628281a";
const GHOST_INS_HIGHLIGHT: &str = "#2e7d321a";
// conservative per-char advance estimate as a fraction of the font size.
const GHOST_CHAR_W_RATIO: f32 = 0.6;
const GHOST_GAP_PX: f32 = 6.0;
const GHOST_MIN_SCALE: f32 = 0.6;

/// a pending edit rendered as a ghost pair in place of the cell's committed
/// text: `old_text` struck in red, `new_text` in green.
#[derive(Debug, Clone, PartialEq)]
pub struct GhostEdit {
    pub row: u32,
    pub col: u32,
    pub old_text: String,
    pub new_text: String,
    pub alignment_value: CellValue,
}

struct GhostFont {
    size: f32,
    family: Option<String>,
    bold: bool,
    italic: bool,
    underline: bool,
}

/// build the display list for one viewport of one sheet. commands are emitted
/// background -> fills -> gridlines -> borders -> text, each layer painting over the previous.
pub fn build_display_list(wb: &Workbook, sheet: SheetId, viewport: &Viewport) -> DisplayList {
    build_display_list_with_ghosts(wb, sheet, viewport, &[])
}

/// like [`build_display_list`], with ghost pairs painted in place of the
/// committed text of each ghost's cell.
pub fn build_display_list_with_ghosts(
    wb: &Workbook,
    sheet: SheetId,
    viewport: &Viewport,
    ghosts: &[GhostEdit],
) -> DisplayList {
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
    let changed_ghost_cells: std::collections::HashSet<(u32, u32)> = ghosts
        .iter()
        .filter(|g| g.old_text != g.new_text)
        .map(|g| (g.row, g.col))
        .collect();

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
        if changed_ghost_cells.contains(&(at.row, at.col)) {
            continue;
        }
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
            highlight: None,
            dashed_underline: false,
            font_family: font.and_then(|f| f.name.clone()),
            ghost: false,
        });
    }

    for ghost in ghosts {
        if !rows.contains(&ghost.row) || !cols.contains(&ghost.col) {
            continue;
        }
        let at = CellRef::new(ghost.row, ghost.col);
        let cell = sheet_ref.cell(at);
        let font = cell.and_then(|c| c.style).and_then(|s| styles.font_for(s));
        let font = GhostFont {
            size: font
                .and_then(|font| font.size_pt)
                .map(|size| size as f32)
                .unwrap_or(FONT_SIZE_PT),
            family: font.and_then(|font| font.name.clone()),
            bold: font.is_some_and(|font| font.bold),
            italic: font.is_some_and(|font| font.italic),
            underline: font.is_some_and(|font| font.underline),
        };
        let bx = cell_box(&geom, viewport, sheet_ref, at);
        let align = resolve_align_with_value(styles, cell, &ghost.alignment_value);
        emit_ghost(&mut commands, ghost, bx, font, align);
    }

    DisplayList {
        width: viewport.width,
        height: viewport.height,
        commands,
        grid,
    }
}

/// paint one pending edit inside a cell box.
fn emit_ghost(
    commands: &mut Vec<DrawCmd>,
    ghost: &GhostEdit,
    (cx0, cy0, cw, ch): (f32, f32, f32, f32),
    font: GhostFont,
    single_align: Align,
) {
    let old = ghost.old_text.as_str();
    let new = ghost.new_text.as_str();
    if old == new {
        return;
    }

    let clip = Rect {
        x: cx0,
        y: cy0,
        w: cw,
        h: ch,
    };
    let x = cx0 + TEXT_PAD_PX;
    let avail = (cw - 2.0 * TEXT_PAD_PX).max(0.0);
    let full_size = font.size;

    let mut line = |x: f32,
                    y: f32,
                    text: String,
                    size: f32,
                    color: &str,
                    align: Align,
                    strike: bool,
                    preview: bool| {
        commands.push(DrawCmd::Text {
            x,
            y,
            text,
            font_size: size,
            color: color.to_string(),
            clip,
            align,
            bold: font.bold,
            italic: font.italic,
            underline: font.underline,
            strike,
            highlight: Some(
                if preview {
                    GHOST_INS_HIGHLIGHT
                } else {
                    GHOST_DEL_HIGHLIGHT
                }
                .to_string(),
            ),
            dashed_underline: preview,
            font_family: font.family.clone(),
            ghost: preview,
        });
    };

    if old.is_empty() || new.is_empty() {
        let (text, color, strike, preview) = if old.is_empty() {
            (new, GHOST_INS_COLOR, false, true)
        } else {
            (old, GHOST_DEL_COLOR, true, false)
        };
        let x = match single_align {
            Align::Left => cx0 + TEXT_PAD_PX,
            Align::Right => cx0 + cw - TEXT_PAD_PX,
            Align::Center => cx0 + cw / 2.0,
        };
        line(
            x,
            baseline_y(cy0, ch, full_size, None),
            ellipsize(text, avail, full_size),
            full_size,
            color,
            single_align,
            strike,
            preview,
        );
        return;
    }

    let old_width = ghost_text_width(old, full_size);
    let new_width = ghost_text_width(new, full_size);
    if old_width + GHOST_GAP_PX + new_width <= avail {
        let baseline = baseline_y(cy0, ch, full_size, None);
        line(
            x,
            baseline,
            old.to_string(),
            full_size,
            GHOST_DEL_COLOR,
            Align::Left,
            true,
            false,
        );
        line(
            x + old_width + GHOST_GAP_PX,
            baseline,
            new.to_string(),
            full_size,
            GHOST_INS_COLOR,
            Align::Left,
            false,
            true,
        );
        return;
    }

    let line_ratio = ASCENT_RATIO + DESCENT_RATIO;
    let scale = (ch / (2.0 * full_size * line_ratio)).clamp(GHOST_MIN_SCALE, 1.0);
    let size = full_size * scale;
    let line_h = size * line_ratio;
    let top = cy0 + ((ch - 2.0 * line_h) / 2.0).max(0.0);
    let first_baseline = top + size * ASCENT_RATIO;
    line(
        x,
        first_baseline,
        ellipsize(new, avail, size),
        size,
        GHOST_INS_COLOR,
        Align::Left,
        false,
        true,
    );
    line(
        x,
        first_baseline + line_h,
        ellipsize(old, avail, size),
        size,
        GHOST_DEL_COLOR,
        Align::Left,
        true,
        false,
    );
}

/// estimated advance width of `text` at `size`, deliberately generous so fit
/// decisions err toward ellipsizing rather than overlap.
fn ghost_text_width(text: &str, size: f32) -> f32 {
    text.chars().count() as f32 * size * GHOST_CHAR_W_RATIO
}

/// `text` unchanged when its estimate fits `budget`, else a truncated prefix
/// ending in `…`.
fn ellipsize(text: &str, budget: f32, size: f32) -> String {
    if ghost_text_width(text, size) <= budget {
        return text.to_string();
    }
    let char_w = size * GHOST_CHAR_W_RATIO;
    let keep = ((budget / char_w) as i32 - 1).max(0) as usize;
    let prefix: String = text.chars().take(keep).collect();
    format!("{prefix}…")
}

/// visible cells that draw: inside the range and not a covered merge cell
/// (only a merge's anchor draws). yields `(anchor, cell)` in row-major order.
fn visible_anchors<'a>(
    sheet: &'a Sheet,
    rows: &'a std::ops::Range<u32>,
    cols: &'a std::ops::Range<u32>,
) -> impl Iterator<Item = (CellRef, &'a xlsx_model::Cell)> {
    sheet.iter_cells_in_rect(rows.clone(), cols.clone()).filter(
        move |(at, _)| match covering_merge(&sheet.merges, *at) {
            Some(m) => m.start == *at,
            None => true,
        },
    )
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
    resolve_align_with_value(styles, Some(cell), &cell.value)
}

fn resolve_align_with_value(
    styles: &Stylesheet,
    cell: Option<&xlsx_model::Cell>,
    value: &CellValue,
) -> Align {
    let type_default = match value {
        CellValue::Number { .. } => Align::Right,
        CellValue::Bool { .. } => Align::Center,
        _ => Align::Left,
    };
    let h = cell
        .and_then(|cell| cell.style)
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

    fn ghost(row: u32, col: u32, old: &str, new: &str) -> GhostEdit {
        ghost_with_alignment_value(
            row,
            col,
            old,
            new,
            CellValue::Text {
                value: new.to_string(),
            },
        )
    }

    fn ghost_with_alignment_value(
        row: u32,
        col: u32,
        old: &str,
        new: &str,
        alignment_value: CellValue,
    ) -> GhostEdit {
        GhostEdit {
            row,
            col,
            old_text: old.into(),
            new_text: new.into(),
            alignment_value,
        }
    }

    fn text_cmds(dl: &DisplayList) -> Vec<(&str, &str, bool, Align)> {
        dl.commands
            .iter()
            .filter_map(|c| match c {
                DrawCmd::Text {
                    text,
                    color,
                    strike,
                    align,
                    ..
                } => Some((text.as_str(), color.as_str(), *strike, *align)),
                _ => None,
            })
            .collect()
    }

    fn ghost_flags(dl: &DisplayList) -> Vec<bool> {
        dl.commands
            .iter()
            .filter_map(|c| match c {
                DrawCmd::Text { ghost, .. } => Some(*ghost),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn ghost_pair_prefers_old_then_new_on_one_line() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 0), num_cell(10.0));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 60.0,
        };
        let dl = build_display_list_with_ghosts(&wb, SheetId(0), &vp, &[ghost(0, 0, "10", "42")]);

        let texts = text_cmds(&dl);
        assert_eq!(texts.len(), 2);
        assert_eq!(texts[0], ("10", GHOST_DEL_COLOR, true, Align::Left));
        assert_eq!(texts[1], ("42", GHOST_INS_COLOR, false, Align::Left));
        assert_eq!(ghost_flags(&dl), vec![false, true]);

        let lines: Vec<(f32, f32, f32)> = dl
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCmd::Text {
                    x, y, font_size, ..
                } => Some((*x, *y, *font_size)),
                _ => None,
            })
            .collect();
        assert!(lines[0].0 < lines[1].0);
        assert_eq!(lines[0].1, lines[1].1);
        assert_eq!((lines[0].2, lines[1].2), (FONT_SIZE_PT, FONT_SIZE_PT));
    }

    #[test]
    fn ghost_insertion_paints_green_only() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Sheet1"));

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 60.0,
        };
        let dl = build_display_list_with_ghosts(
            &wb,
            SheetId(0),
            &vp,
            &[ghost_with_alignment_value(
                1,
                1,
                "",
                "7",
                CellValue::Number { value: 7.0 },
            )],
        );

        let texts = text_cmds(&dl);
        assert_eq!(texts, vec![("7", GHOST_INS_COLOR, false, Align::Right)]);
        assert_eq!(ghost_flags(&dl), vec![true]);
        let (x, clip) = dl
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCmd::Text { x, clip, .. } => Some((*x, *clip)),
                _ => None,
            })
            .unwrap();
        assert_eq!(x, clip.x + clip.w - TEXT_PAD_PX);
    }

    #[test]
    fn single_ghosts_honor_explicit_alignment_and_deleted_value_type() {
        let mut wb = Workbook::default();
        let style = wb
            .styles
            .intern_cell_format(&xlsx_model::CellFormat {
                alignment: xlsx_model::Alignment {
                    h: Some(HAlign::Left),
                    ..xlsx_model::Alignment::default()
                },
                ..xlsx_model::CellFormat::default()
            })
            .unwrap();
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                style: Some(style),
                ..Cell::default()
            },
        );
        sheet.set_cell(CellRef::new(1, 0), num_cell(7.0));
        wb.sheets.push(sheet);

        let dl = build_display_list_with_ghosts(
            &wb,
            SheetId(0),
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 60.0,
            },
            &[
                ghost_with_alignment_value(0, 0, "", "7", CellValue::Number { value: 7.0 }),
                ghost_with_alignment_value(1, 0, "7", "", CellValue::Number { value: 7.0 }),
            ],
        );

        assert_eq!(
            text_cmds(&dl),
            vec![
                ("7", GHOST_INS_COLOR, false, Align::Left),
                ("7", GHOST_DEL_COLOR, true, Align::Right),
            ]
        );
    }

    #[test]
    fn stacked_ghost_pair_puts_new_value_on_top() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 0), text_cell("previous long value"));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 60.0,
        };
        let dl = build_display_list_with_ghosts(
            &wb,
            SheetId(0),
            &vp,
            &[ghost(0, 0, "previous long value", "replacement long value")],
        );

        let texts = text_cmds(&dl);
        assert_eq!(texts.len(), 2);
        assert!(texts[0].0.ends_with('…') && !texts[0].2);
        assert!(texts[1].0.ends_with('…') && texts[1].2);
        assert_eq!((texts[0].3, texts[1].3), (Align::Left, Align::Left));

        let lines: Vec<_> = dl
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCmd::Text { y, font_size, .. } => Some((*y, *font_size)),
                _ => None,
            })
            .collect();
        assert!(lines[0].0 < lines[1].0);
        assert_eq!((lines[0].1, lines[1].1), (FONT_SIZE_PT, FONT_SIZE_PT));
    }

    #[test]
    fn short_rows_shrink_stacked_ghosts_without_overlap() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.row_heights.insert(0, 7.5);
        sheet.set_cell(CellRef::new(0, 0), num_cell(10.0));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 80.0,
        };
        let dl = build_display_list_with_ghosts(
            &wb,
            SheetId(0),
            &vp,
            &[ghost(0, 0, "previous", "replacement")],
        );

        let lines: Vec<(f32, f32)> = dl
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCmd::Text { y, font_size, .. } => Some((*y, *font_size)),
                _ => None,
            })
            .collect();
        let texts = text_cmds(&dl);
        assert!(!texts[0].2 && texts[1].2);
        assert_eq!(lines.len(), 2);
        assert_eq!(
            (lines[0].1, lines[1].1),
            (
                FONT_SIZE_PT * GHOST_MIN_SCALE,
                FONT_SIZE_PT * GHOST_MIN_SCALE
            )
        );
        assert!(
            lines[1].0 - lines[0].0
                >= FONT_SIZE_PT * GHOST_MIN_SCALE * (ASCENT_RATIO + DESCENT_RATIO) - 0.01
        );
    }

    #[test]
    fn equal_formatted_values_keep_the_committed_cell_text() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 0), num_cell(4855.0));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 60.0,
        };
        let dl =
            build_display_list_with_ghosts(&wb, SheetId(0), &vp, &[ghost(0, 0, "4855", "4855")]);

        assert_eq!(
            text_cmds(&dl),
            vec![("4855", TEXT_COLOR, false, Align::Right)]
        );
        assert_eq!(ghost_flags(&dl), vec![false]);
        assert_eq!(
            dl.commands
                .iter()
                .filter(|command| matches!(command, DrawCmd::FillRect { .. }))
                .count(),
            1
        );
    }

    #[test]
    fn ghost_runs_carry_revision_highlights_and_new_underline() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 0), num_cell(10.0));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);
        let dl = build_display_list_with_ghosts(
            &wb,
            SheetId(0),
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 60.0,
            },
            &[ghost(0, 0, "10", "42")],
        );
        let styles: Vec<_> = dl
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCmd::Text {
                    highlight,
                    dashed_underline,
                    ..
                } => Some((highlight.as_deref(), *dashed_underline)),
                _ => None,
            })
            .collect();
        assert_eq!(
            styles,
            vec![
                (Some(GHOST_DEL_HIGHLIGHT), false),
                (Some(GHOST_INS_HIGHLIGHT), true)
            ]
        );
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
