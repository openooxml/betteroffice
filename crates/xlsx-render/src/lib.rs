//! grid geometry and the target-agnostic display list: turns a workbook
//! viewport into draw commands. never imports canvas, dom, or any raster backend.

pub mod chart;
pub mod display_list;
pub mod geometry;
pub mod hit;
pub mod region;

pub use ooxml_drawingml::GeometryPathCommand;
use std::sync::Arc;

use hashbrown::HashMap;
use ooxml_drawingml::chart::ChartSpace;
use std::collections::BTreeMap;
use std::ops::{Range, RangeInclusive};

use serde::{Deserialize, Serialize};

use xlsx_model::numfmt::{builtin_format_code, format_value};
use xlsx_model::styles::{BorderEdge, BorderStyle, FormatCode, Stylesheet};
use xlsx_model::value::CellValue;
use xlsx_model::workbook::{Hyperlink, Sheet};
use xlsx_model::{
    CellRange, CellRef, Fill, HAlign, MAX_COLS, MAX_ROWS, SheetChart, SheetId, VAlign, Workbook,
};

pub use chart::{
    AnchorError, MAX_CHART_OPS_PER_FRAME, RenderError, ResolvedChartAnchor, chart_regions,
    moved_chart_anchor, resolve_chart_anchor,
};
pub use display_list::{
    Align, ChartA11yAttrs, ChartRegion, DisplayList, DrawCmd, GridMeta, HyperlinkRegion,
    PathStroke, Rect, scaled,
};
pub use geometry::{GridGeometry, PrintMetrics, autofit_relevant};
pub use hit::chart_at_point;
pub use region::{viewport_for_range, viewport_for_used_range, viewport_for_used_range_within};

/// a scrolled window in pixels; `x`/`y` offset the non-frozen body.
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
const PANE_DIVIDER_COLOR: &str = "#8a8a8a";
const HYPERLINK_COLOR: &str = "#0563c1";
const GRIDLINE_WIDTH: f32 = 1.0;
const PANE_DIVIDER_WIDTH: f32 = 2.0;
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
    family: Option<Arc<str>>,
    bold: bool,
    italic: bool,
    underline: bool,
}

#[derive(Clone, Copy)]
struct AxisTrack {
    index: u32,
    raw_start: f32,
    start: f32,
    end: f32,
    pinned: bool,
}

struct AxisLayout {
    tracks: Vec<AxisTrack>,
    ranges: Vec<Range<u32>>,
    divider: Option<f32>,
    frozen: u32,
    scroll: f32,
    print_extent: Option<f32>,
}

#[derive(Clone, Copy)]
struct AxisSpan {
    raw_start: f32,
    raw_end: f32,
    start: f32,
    end: f32,
}

impl AxisLayout {
    fn new(
        limit: u32,
        frozen: u32,
        scroll: f32,
        extent: f32,
        edge: impl Fn(u32) -> f32,
        at: impl Fn(f32) -> u32,
    ) -> Self {
        let frozen = frozen.min(limit);
        let scroll = if frozen > 0 { scroll.max(0.0) } else { scroll };
        let frozen_extent = edge(frozen);
        let mut tracks = Vec::new();
        if frozen > 0 {
            for index in 0..frozen {
                let start = edge(index);
                if start >= extent {
                    break;
                }
                tracks.push(AxisTrack {
                    index,
                    raw_start: start,
                    start,
                    end: edge(index + 1),
                    pinned: true,
                });
            }
        }

        if frozen < limit && frozen_extent < extent {
            let body_extent = extent - frozen_extent;
            let origin = frozen_extent + scroll;
            let first = at(origin).max(frozen).min(limit - 1);
            let last = at(origin + body_extent).max(first).min(limit - 1);
            for index in first..=last {
                let raw_start = edge(index) - scroll;
                let raw_end = edge(index + 1) - scroll;
                tracks.push(AxisTrack {
                    index,
                    raw_start,
                    start: if frozen > 0 && index == first {
                        raw_start.max(frozen_extent)
                    } else {
                        raw_start
                    },
                    end: raw_end,
                    pinned: false,
                });
            }
        }

        let mut ranges: Vec<Range<u32>> = Vec::new();
        for track in &tracks {
            match ranges.last_mut() {
                Some(range) if range.end == track.index => range.end += 1,
                _ => ranges.push(track.index..track.index + 1),
            }
        }
        Self {
            tracks,
            ranges,
            divider: (frozen > 0 && frozen_extent < extent).then_some(frozen_extent),
            frozen,
            scroll,
            print_extent: None,
        }
    }

    fn start(&self) -> u32 {
        self.tracks.first().map_or(0, |track| track.index)
    }

    fn indices(&self) -> Option<Vec<u32>> {
        let start = self.start();
        self.tracks
            .iter()
            .enumerate()
            .any(|(offset, track)| track.index != start + offset as u32)
            .then(|| self.tracks.iter().map(|track| track.index).collect())
    }

    fn offsets(&self) -> Vec<f32> {
        let Some(first) = self.tracks.first() else {
            return Vec::new();
        };
        std::iter::once(first.start)
            .chain(self.tracks.iter().map(|track| track.end))
            .collect()
    }

    fn contains(&self, index: u32) -> bool {
        self.tracks
            .binary_search_by_key(&index, |track| track.index)
            .is_ok()
    }

    fn intersects(&self, start: u32, end: u32) -> bool {
        self.ranges
            .iter()
            .any(|range| range.start <= end && range.end > start)
    }

    fn span(&self, start: u32, end: u32, edge: impl Fn(u32) -> f32) -> Option<AxisSpan> {
        if let Some(extent) = self.print_extent {
            if !self.intersects(start, end) {
                return None;
            }
            let raw_start = edge(start) - self.scroll;
            let raw_end = edge(end.saturating_add(1)) - self.scroll;
            let start = raw_start.max(0.0);
            let end = raw_end.min(extent);
            return (end > start).then_some(AxisSpan {
                raw_start,
                raw_end,
                start,
                end,
            });
        }
        let first = self
            .tracks
            .binary_search_by_key(&start, |track| track.index)
            .ok()?;
        let track = self.tracks[first];
        let raw_end = if track.pinned {
            edge(end.saturating_add(1).min(self.frozen))
        } else {
            edge(end.saturating_add(1)) - self.scroll
        };
        Some(AxisSpan {
            raw_start: track.raw_start,
            raw_end,
            start: track.start,
            end: raw_end,
        })
    }
}

#[derive(Clone, Copy)]
struct CellBox {
    x: f32,
    y: f32,
    w: f32,
    h: f32,
    clip: Rect,
}

/// Builds a display list for one worksheet viewport, with every chart drawn as
/// a placeholder because no chart source is supplied.
pub fn build_display_list(
    wb: &Workbook,
    sheet: SheetId,
    viewport: &Viewport,
) -> Result<DisplayList, RenderError> {
    build_display_list_with_ghosts(wb, sheet, viewport, &[])
}

/// Builds a chart-placeholder display list with pending edit ghosts.
pub fn build_display_list_with_ghosts(
    wb: &Workbook,
    sheet: SheetId,
    viewport: &Viewport,
    ghosts: &[GhostEdit],
) -> Result<DisplayList, RenderError> {
    build_display_list_with_charts_and_ghosts(wb, sheet, viewport, ghosts, |chart| {
        Err::<ChartSpace, _>(RenderError::ChartSourceUnavailable {
            part: chart.part.clone(),
        })
    })
}

/// Builds a display list using a lazy chart-part resolver.
pub fn build_display_list_with_charts<F, R>(
    wb: &Workbook,
    sheet: SheetId,
    viewport: &Viewport,
    resolver: F,
) -> Result<DisplayList, RenderError>
where
    F: FnMut(&SheetChart) -> Result<R, RenderError>,
    R: Into<Arc<ChartSpace>>,
{
    build_display_list_with_charts_and_ghosts(wb, sheet, viewport, &[], resolver)
}

/// Builds a display list with charts and pending edit ghosts.
pub fn build_display_list_with_charts_and_ghosts<F, R>(
    wb: &Workbook,
    sheet: SheetId,
    viewport: &Viewport,
    ghosts: &[GhostEdit],
    resolver: F,
) -> Result<DisplayList, RenderError>
where
    F: FnMut(&SheetChart) -> Result<R, RenderError>,
    R: Into<Arc<ChartSpace>>,
{
    build_frame(wb, sheet, viewport, ghosts, resolver, None)
}

pub fn build_print_display_list_with_charts<F, R>(
    wb: &Workbook,
    sheet: SheetId,
    viewport: &Viewport,
    metrics: &PrintMetrics,
    gridlines: bool,
    resolver: F,
) -> Result<DisplayList, RenderError>
where
    F: FnMut(&SheetChart) -> Result<R, RenderError>,
    R: Into<Arc<ChartSpace>>,
{
    build_frame(
        wb,
        sheet,
        viewport,
        &[],
        resolver,
        Some((metrics, gridlines)),
    )
}

fn build_frame<F, R>(
    wb: &Workbook,
    sheet: SheetId,
    viewport: &Viewport,
    ghosts: &[GhostEdit],
    mut resolver: F,
    print: Option<(&PrintMetrics, bool)>,
) -> Result<DisplayList, RenderError>
where
    F: FnMut(&SheetChart) -> Result<R, RenderError>,
    R: Into<Arc<ChartSpace>>,
{
    let background_color: Arc<str> = BACKGROUND_COLOR.into();
    let mut commands = Vec::new();

    commands.push(DrawCmd::FillRect {
        x: 0.0,
        y: 0.0,
        w: viewport.width,
        h: viewport.height,
        color: background_color.clone(),
        clip: None,
    });

    let Some(sheet_ref) = wb.sheet(sheet) else {
        return Ok(DisplayList {
            width: viewport.width,
            height: viewport.height,
            commands,
            grid: GridMeta::default(),
            hyperlinks: Vec::new(),
            charts: Vec::new(),
        });
    };

    let styles = &wb.styles;
    let theme = &styles.theme;
    let hyperlink_color: Arc<str> = theme.slot(10).unwrap_or(HYPERLINK_COLOR).into();
    let text_color: Arc<str> = TEXT_COLOR.into();
    let gridline_color: Arc<str> = if print.is_some() {
        TEXT_COLOR
    } else {
        GRIDLINE_COLOR
    }
    .into();
    let pane_divider_color: Arc<str> = PANE_DIVIDER_COLOR.into();
    let print_font_family: Option<Arc<str>> = print.map(|(m, _)| m.font_family.as_str().into());
    let mut fx = FrameStyles::new(styles);
    let geom = print.map_or_else(
        || GridGeometry::new(sheet_ref, styles),
        |(metrics, _)| GridGeometry::for_print(sheet_ref, styles, metrics),
    );
    let (frozen_rows, frozen_cols) = if print.is_some() {
        (0, 0)
    } else {
        sheet_ref
            .freeze_pane
            .map_or((0, 0), |pane| (pane.rows, pane.cols))
    };
    let mut rows = AxisLayout::new(
        MAX_ROWS,
        frozen_rows,
        viewport.y,
        viewport.height,
        |row| geom.row_y(row),
        |y| geom.row_at_y(y),
    );
    let mut cols = AxisLayout::new(
        MAX_COLS,
        frozen_cols,
        viewport.x,
        viewport.width,
        |col| geom.col_x(col),
        |x| geom.col_at_x(x),
    );
    if print.is_some() {
        rows.print_extent = Some(viewport.height);
        cols.print_extent = Some(viewport.width);
    }

    let grid = GridMeta {
        start_row: rows.start(),
        start_col: cols.start(),
        row_indices: rows.indices(),
        col_indices: cols.indices(),
        row_offsets: rows.offsets(),
        col_offsets: cols.offsets(),
    };
    let hyperlinks = sheet_ref
        .hyperlinks
        .iter()
        .filter(|link| {
            rows.intersects(link.range.start.row, link.range.end.row)
                && cols.intersects(link.range.start.col, link.range.end.col)
        })
        .map(|link| HyperlinkRegion {
            top: link.range.start.row,
            left: link.range.start.col,
            bottom: link.range.end.row,
            right: link.range.end.col,
            external_target: link.external_target.clone(),
            location: link.location.clone(),
            tooltip: link.tooltip.clone(),
        })
        .collect();
    let link_index = HyperlinkIndex::new(&sheet_ref.hyperlinks);
    let row_hint = match (rows.ranges.first(), rows.ranges.last()) {
        (Some(first), Some(last)) => first.start..=last.end.saturating_sub(1),
        _ => RangeInclusive::new(1, 0),
    };
    let merge_index = MergeIndex::new(&sheet_ref.merges, row_hint);
    let mut anchors = visible_anchors(sheet_ref, &rows, &cols, &merge_index);
    if print.is_some() {
        for merge in &sheet_ref.merges {
            if rows.intersects(merge.start.row, merge.end.row)
                && cols.intersects(merge.start.col, merge.end.col)
                && let Some(cell) = sheet_ref.cell(merge.start)
            {
                anchors.push((merge.start, cell));
            }
        }
        anchors.sort_unstable_by_key(|(at, _)| (at.row, at.col));
        anchors.dedup_by_key(|(at, _)| (at.row, at.col));
        anchors.retain(|(at, _)| {
            geom.row_y(at.row) < viewport.y + viewport.height
                && geom.col_x(at.col) < viewport.x + viewport.width
        });
    }
    let changed_ghost_cells: std::collections::HashSet<(u32, u32)> = ghosts
        .iter()
        .filter(|g| g.old_text != g.new_text)
        .map(|g| (g.row, g.col))
        .collect();

    let mut grid_commands = Vec::new();
    let grid_offset = print.map_or(0.0, |(m, _)| 48.0 / m.dpi);
    let row_offsets = rows.offsets();
    let col_offsets = cols.offsets();
    let top = row_offsets.first().copied().unwrap_or(0.0);
    let bottom = row_offsets.last().copied().unwrap_or(0.0);
    let left = col_offsets.first().copied().unwrap_or(0.0);
    let right = col_offsets.last().copied().unwrap_or(0.0);
    for &x in &col_offsets {
        grid_commands.push(DrawCmd::Line {
            x1: x + grid_offset,
            y1: top + grid_offset,
            x2: x + grid_offset,
            y2: bottom + grid_offset,
            width: print.map_or(GRIDLINE_WIDTH, |(m, _)| 96.0 / m.dpi),
            color: gridline_color.clone(),
            style: None,
            clip: None,
        });
    }
    for &y in &row_offsets {
        grid_commands.push(DrawCmd::Line {
            x1: left + grid_offset,
            y1: y + grid_offset,
            x2: right + grid_offset,
            y2: y + grid_offset,
            width: print.map_or(GRIDLINE_WIDTH, |(m, _)| 96.0 / m.dpi),
            color: gridline_color.clone(),
            style: None,
            clip: None,
        });
    }

    if let Some((metrics, gridlines)) = print {
        if gridlines {
            commands.extend(grid_commands.iter().cloned());
        }
        for merge in &sheet_ref.merges {
            if let Some(cell_box) = cell_box(&geom, &rows, &cols, &merge_index, merge.start) {
                let inset = 96.0 / metrics.dpi / 2.0;
                commands.push(DrawCmd::FillRect {
                    x: cell_box.x + inset,
                    y: cell_box.y + inset,
                    w: (cell_box.w - 2.0 * inset).max(0.0),
                    h: (cell_box.h - 2.0 * inset).max(0.0),
                    color: background_color.clone(),
                    clip: Some(cell_box.clip),
                });
            }
        }
    }

    for &(at, cell) in &anchors {
        let Some(hex) = &fx.xf(cell.style).fill else {
            continue;
        };
        let Some(cell_box) = cell_box(&geom, &rows, &cols, &merge_index, at) else {
            continue;
        };
        let clip = cell_box.clip;
        commands.push(DrawCmd::FillRect {
            x: clip.x,
            y: clip.y,
            w: clip.w,
            h: clip.h,
            color: hex.clone(),
            clip: None,
        });
    }

    if print.is_none() {
        commands.extend(grid_commands);
    }
    if let Some((metrics, true)) = print {
        let width = 96.0 / metrics.dpi;
        let offset = width / 2.0;
        for (x1, y1, x2, y2) in [
            (offset, offset, viewport.width + offset, offset),
            (offset, offset, offset, viewport.height + offset),
            (
                offset,
                viewport.height + offset,
                viewport.width + offset,
                viewport.height + offset,
            ),
            (
                viewport.width + offset,
                offset,
                viewport.width + offset,
                viewport.height + offset,
            ),
        ] {
            commands.push(DrawCmd::Line {
                x1,
                y1,
                x2,
                y2,
                width,
                color: text_color.clone(),
                style: None,
                clip: None,
            });
        }
    }

    for &(at, cell) in &anchors {
        let xf = fx.xf(cell.style).clone();
        let Some(border) = &xf.border else {
            continue;
        };
        emit_borders(
            &mut commands,
            &geom,
            &rows,
            &cols,
            sheet_ref,
            &merge_index,
            &mut fx,
            at,
            border,
        );
    }

    for &(at, cell) in &anchors {
        if changed_ghost_cells.contains(&(at.row, at.col)) {
            continue;
        }
        let hyperlink = link_index.at(at);
        let xf = &**fx.xf(cell.style);
        let Some((text, color)) =
            cell_display_text(xf, wb.date_system, cell, &text_color).or_else(|| {
                hyperlink
                    .filter(|link| link.range.start == at)
                    .and_then(|link| link.display.clone())
                    .filter(|display| !display.is_empty())
                    .map(|display| (display.into(), hyperlink_color.clone()))
            })
        else {
            continue;
        };
        let color = if hyperlink.is_some() {
            hyperlink_color.clone()
        } else {
            color
        };

        let Some(cell_box) = cell_box(&geom, &rows, &cols, &merge_index, at) else {
            continue;
        };
        let font = xf.font.as_ref();
        let size = font
            .and_then(|f| f.size_pt)
            .map(|p| p as f32)
            .unwrap_or_else(|| print.map_or(FONT_SIZE_PT, |(m, _)| m.font_size_pt));
        let align = resolve_align(xf.h, &cell.value);
        let valign = xf.v;
        let bold = font.is_some_and(|f| f.bold);
        let italic = font.is_some_and(|f| f.italic);
        let underline = font.is_some_and(|f| f.underline);
        let strike = font.is_some_and(|f| f.strike);
        let font_family = font
            .and_then(|f| f.family.clone())
            .or_else(|| print_font_family.clone());

        let tx = match align {
            Align::Left => cell_box.x + print.map_or(TEXT_PAD_PX, |(m, _)| 3.0 * 96.0 / m.dpi),
            Align::Right => {
                cell_box.x + cell_box.w - print.map_or(TEXT_PAD_PX, |(m, _)| 2.0 * 96.0 / m.dpi)
            }
            Align::Center => cell_box.x + cell_box.w / 2.0,
        };
        let ty = text_baseline(cell_box, size, valign, print.map(|(m, _)| m));
        let clip = spill_clip(
            &geom,
            &cols,
            sheet_ref,
            &link_index,
            &merge_index,
            &mut fx,
            wb.date_system,
            at,
            cell,
            align,
            cell_box,
            &text_color,
        );

        commands.push(DrawCmd::Text {
            x: tx,
            y: ty,
            text,
            font_size: size,
            color,
            clip,
            align,
            bold,
            italic,
            underline: hyperlink.is_some() || underline,
            strike,
            highlight: None,
            dashed_underline: false,
            font_family,
            ghost: false,
            chart: false,
        });
    }

    for link in &sheet_ref.hyperlinks {
        let at = link.range.start;
        if sheet_ref.cell(at).is_some()
            || (print.is_none() && (!rows.contains(at.row) || !cols.contains(at.col)))
        {
            continue;
        }
        let Some(text) = link.display.as_ref().filter(|display| !display.is_empty()) else {
            continue;
        };
        let Some(cell_box) = cell_box(&geom, &rows, &cols, &merge_index, at) else {
            continue;
        };
        let size = print.map_or(FONT_SIZE_PT, |(m, _)| m.font_size_pt);
        commands.push(DrawCmd::Text {
            x: cell_box.x + print.map_or(TEXT_PAD_PX, |(m, _)| 3.0 * 96.0 / m.dpi),
            y: text_baseline(cell_box, size, None, print.map(|(m, _)| m)),
            text: text.as_str().into(),
            font_size: size,
            color: hyperlink_color.clone(),
            clip: cell_box.clip,
            align: Align::Left,
            bold: false,
            italic: false,
            underline: true,
            strike: false,
            highlight: None,
            dashed_underline: false,
            font_family: print_font_family.clone(),
            ghost: false,
            chart: false,
        });
    }

    for ghost in ghosts {
        if !rows.contains(ghost.row) || !cols.contains(ghost.col) {
            continue;
        }
        let at = CellRef::new(ghost.row, ghost.col);
        let cell = sheet_ref.cell(at);
        let xf = fx.xf(cell.and_then(|c| c.style)).clone();
        let font = xf.font.as_ref();
        let font = GhostFont {
            size: font
                .and_then(|font| font.size_pt)
                .map(|size| size as f32)
                .unwrap_or(FONT_SIZE_PT),
            family: font.and_then(|font| font.family.clone()),
            bold: font.is_some_and(|font| font.bold),
            italic: font.is_some_and(|font| font.italic),
            underline: font.is_some_and(|font| font.underline),
        };
        let Some(bx) = cell_box(&geom, &rows, &cols, &merge_index, at) else {
            continue;
        };
        let align = resolve_align(xf.h, &ghost.alignment_value);
        emit_ghost(&mut commands, ghost, bx, font, align);
    }

    let mut charts = Vec::new();
    chart::render_charts(
        sheet_ref,
        &geom,
        viewport,
        frozen_rows,
        frozen_cols,
        &mut commands,
        &mut charts,
        &mut resolver,
    )?;

    if let Some(x) = cols.divider {
        commands.push(DrawCmd::Line {
            x1: x,
            y1: 0.0,
            x2: x,
            y2: viewport.height,
            width: PANE_DIVIDER_WIDTH,
            color: pane_divider_color.clone(),
            style: None,
            clip: None,
        });
    }
    if let Some(y) = rows.divider {
        commands.push(DrawCmd::Line {
            x1: 0.0,
            y1: y,
            x2: viewport.width,
            y2: y,
            width: PANE_DIVIDER_WIDTH,
            color: pane_divider_color.clone(),
            style: None,
            clip: None,
        });
    }

    Ok(DisplayList {
        width: viewport.width,
        height: viewport.height,
        commands,
        grid,
        hyperlinks,
        charts,
    })
}

/// paint one pending edit inside a cell box.
fn emit_ghost(
    commands: &mut Vec<DrawCmd>,
    ghost: &GhostEdit,
    cell_box: CellBox,
    font: GhostFont,
    single_align: Align,
) {
    let old = ghost.old_text.as_str();
    let new = ghost.new_text.as_str();
    if old == new {
        return;
    }

    let (cx0, cy0, cw, ch) = (cell_box.x, cell_box.y, cell_box.w, cell_box.h);
    let clip = cell_box.clip;
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
            text: text.into(),
            font_size: size,
            color: color.into(),
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
                .into(),
            ),
            dashed_underline: preview,
            font_family: font.family.clone(),
            ghost: preview,
            chart: false,
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

    let line_ratio = (ASCENT_RATIO + DESCENT_RATIO) * geometry::PX_PER_PT as f32;
    let scale = (ch / (2.0 * full_size * line_ratio)).clamp(GHOST_MIN_SCALE, 1.0);
    let size = full_size * scale;
    let line_h = size * line_ratio;
    let top = cy0 + ((ch - 2.0 * line_h) / 2.0).max(0.0);
    let first_baseline = top + size * geometry::PX_PER_PT as f32 * ASCENT_RATIO;
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
    text.chars().count() as f32 * size * geometry::PX_PER_PT as f32 * GHOST_CHAR_W_RATIO
}

/// `text` unchanged when its estimate fits `budget`, else a truncated prefix
/// ending in `…`.
fn ellipsize(text: &str, budget: f32, size: f32) -> String {
    if ghost_text_width(text, size) <= budget {
        return text.to_string();
    }
    let char_w = size * geometry::PX_PER_PT as f32 * GHOST_CHAR_W_RATIO;
    let keep = ((budget / char_w) as i32 - 1).max(0) as usize;
    let prefix: String = text.chars().take(keep).collect();
    format!("{prefix}…")
}

/// visible cells that draw: inside the range and not a covered merge cell
/// (only a merge's anchor draws). yields `(anchor, cell)` in row-major order.
fn visible_anchors<'a>(
    sheet: &'a Sheet,
    rows: &AxisLayout,
    cols: &AxisLayout,
    merges: &MergeIndex,
) -> Vec<(CellRef, &'a xlsx_model::Cell)> {
    let mut cells = Vec::new();
    for row_range in &rows.ranges {
        for col_range in &cols.ranges {
            cells.extend(sheet.iter_cells_in_rect(row_range.clone(), col_range.clone()));
        }
    }
    cells.sort_unstable_by_key(|(at, _)| (at.row, at.col));
    cells.dedup_by_key(|(at, _)| (at.row, at.col));
    cells.retain(|(at, _)| match merges.covering(*at) {
        Some(merge) => merge.start == *at,
        None => true,
    });
    cells
}

/// whether the painter draws a hyperlink's own label at `at`, which it does
/// only at the link range's start and only when the cell has no text of its own.
fn draws_hyperlink_label(links: &HyperlinkIndex, at: CellRef) -> bool {
    links.at(at).is_some_and(|link| {
        link.range.start == at && link.display.as_ref().is_some_and(|d| !d.is_empty())
    })
}

/// per-pass merge coverage: merges shorter than `MERGE_ROW_CAP` are indexed
/// per hinted row, taller or inverted merges are `contains`-scanned.
struct MergeIndex<'a> {
    merges: &'a [CellRange],
    rows: BTreeMap<u32, MergeRow>,
    entries: Vec<(u32, u32, u32, u32)>,
    scanned: Vec<u32>,
}

#[derive(Clone, Copy)]
struct MergeRow {
    /// span of this row's merges in `MergeIndex::entries`.
    start: u32,
    end: u32,
    /// column-disjoint entries sort by start column and binary-search.
    by_col: bool,
}

/// row-height cutoff between indexed and scanned merges.
const MERGE_ROW_CAP: u32 = 64;

impl<'a> MergeIndex<'a> {
    /// indexes short merges over `hint` rows (plus every anchor row); other
    /// merges — inverted, taller than `MERGE_ROW_CAP` — are scanned per query.
    fn new(merges: &'a [CellRange], hint: RangeInclusive<u32>) -> Self {
        let (hint_start, hint_end) = (*hint.start(), *hint.end());
        let mut entries: Vec<(u32, u32, u32, u32)> = Vec::new();
        let mut scanned = Vec::new();
        for (index, merge) in merges.iter().enumerate() {
            if merge.end.row < merge.start.row || merge.end.row - merge.start.row >= MERGE_ROW_CAP {
                scanned.push(index as u32);
                continue;
            }
            entries.push((
                merge.start.row,
                merge.start.col,
                merge.end.col,
                index as u32,
            ));
            for row in merge.start.row.max(hint_start)..=merge.end.row.min(hint_end) {
                if row != merge.start.row {
                    entries.push((row, merge.start.col, merge.end.col, index as u32));
                }
            }
        }
        entries.sort_by_key(|entry| (entry.0, entry.1));
        let mut rows = BTreeMap::new();
        let mut i = 0;
        while i < entries.len() {
            let mut j = i + 1;
            while j < entries.len() && entries[j].0 == entries[i].0 {
                j += 1;
            }
            let group = &mut entries[i..j];
            let by_col = group.windows(2).all(|w| w[0].2 < w[1].1);
            if !by_col {
                group.sort_by_key(|entry| entry.3);
            }
            rows.insert(
                entries[i].0,
                MergeRow {
                    start: i as u32,
                    end: j as u32,
                    by_col,
                },
            );
            i = j;
        }
        Self {
            merges,
            rows,
            entries,
            scanned,
        }
    }

    /// the merge (if any) that covers `at`.
    fn covering(&self, at: CellRef) -> Option<CellRange> {
        let in_row = self.rows.get(&at.row).and_then(|row| {
            let entries = &self.entries[row.start as usize..row.end as usize];
            if row.by_col {
                let p = entries.partition_point(|entry| entry.1 <= at.col);
                p.checked_sub(1)
                    .and_then(|i| entries.get(i))
                    .filter(|entry| entry.2 >= at.col)
                    .map(|entry| entry.3)
            } else {
                entries
                    .iter()
                    .find(|entry| entry.1 <= at.col && at.col <= entry.2)
                    .map(|entry| entry.3)
            }
        });
        let scanned = self
            .scanned
            .iter()
            .copied()
            .find(|&i| self.merges[i as usize].contains(at));
        match (in_row, scanned) {
            (Some(m), Some(s)) => Some(self.merges[m.min(s) as usize]),
            (Some(m), None) => Some(self.merges[m as usize]),
            (None, s) => s.map(|i| self.merges[i as usize]),
        }
    }
}

/// Excel lets text that outgrows its cell run into the blank cells beside it,
/// towards whichever side its alignment points. The run stops at the first
/// neighbour that draws something, at a merge, at a frozen-pane split, and at
/// the edge of the columns this frame lays out.
#[allow(clippy::too_many_arguments)]
fn spill_clip(
    geom: &GridGeometry,
    cols: &AxisLayout,
    sheet: &Sheet,
    links: &HyperlinkIndex,
    merges: &MergeIndex,
    fx: &mut FrameStyles,
    date_system: xlsx_model::DateSystem,
    at: CellRef,
    cell: &xlsx_model::Cell,
    align: Align,
    cell_box: CellBox,
    default_color: &Arc<str>,
) -> Rect {
    if !matches!(cell.value, CellValue::Text { .. })
        || merges.covering(at).is_some()
        || fx.xf(cell.style).wrap_or_shrink
    {
        return cell_box.clip;
    }
    let Ok(anchor) = cols
        .tracks
        .binary_search_by_key(&at.col, |track| track.index)
    else {
        return cell_box.clip;
    };
    let mut blank = |col: u32| {
        let neighbour = CellRef::new(at.row, col);
        merges.covering(neighbour).is_none()
            && sheet.cell(neighbour).is_none_or(|cell| {
                cell_display_text(fx.xf(cell.style), date_system, cell, default_color).is_none()
            })
            && !draws_hyperlink_label(links, neighbour)
    };
    let pane = cols.tracks[anchor].pinned;
    let mut first = anchor;
    if matches!(align, Align::Right | Align::Center) {
        while first > 0
            && cols.tracks[first - 1].pinned == pane
            && cols.tracks[first - 1].index + 1 == cols.tracks[first].index
            && blank(cols.tracks[first - 1].index)
        {
            first -= 1;
        }
    }
    let mut last = anchor;
    if matches!(align, Align::Left | Align::Center) {
        while last + 1 < cols.tracks.len()
            && cols.tracks[last + 1].pinned == pane
            && cols.tracks[last].index + 1 == cols.tracks[last + 1].index
            && blank(cols.tracks[last + 1].index)
        {
            last += 1;
        }
    }
    if first == anchor && last == anchor {
        return cell_box.clip;
    }
    cols.span(
        cols.tracks[first].index,
        cols.tracks[last].index,
        |column| geom.col_x(column),
    )
    .map_or(cell_box.clip, |span| Rect {
        x: span.start,
        y: cell_box.clip.y,
        w: span.end - span.start,
        h: cell_box.clip.h,
    })
}

/// hyperlink lookup for one render pass: per-row entries when ranges are
/// column-disjoint, else scans `hyperlinks` order like a linear search.
struct HyperlinkIndex<'a> {
    links: &'a [Hyperlink],
    rows: BTreeMap<u32, LinkRow>,
    entries: Vec<(u32, u32, u32, u32)>,
    scanned: Vec<u32>,
}

#[derive(Clone, Copy)]
struct LinkRow {
    /// span of this row's links in `HyperlinkIndex::entries`.
    start: u32,
    end: u32,
    /// true when the row's ranges are column-disjoint and binary-searchable.
    by_col: bool,
}

/// a range spanning more rows than this is scanned instead of indexed per row.
const LINK_ROW_CAP: u32 = 64;

/// total expanded row entries an index build may hold before further ranges
/// fall back to scanning.
const LINK_ENTRY_CAP: usize = 1 << 20;

impl<'a> HyperlinkIndex<'a> {
    fn new(links: &'a [Hyperlink]) -> Self {
        let mut entries: Vec<(u32, u32, u32, u32)> = Vec::new();
        let mut scanned = Vec::new();
        for (index, link) in links.iter().enumerate() {
            let range = link.range;
            if range.end.row - range.start.row >= LINK_ROW_CAP
                || entries.len() + (range.end.row - range.start.row + 1) as usize > LINK_ENTRY_CAP
            {
                scanned.push(index as u32);
                continue;
            }
            for row in range.start.row..=range.end.row {
                entries.push((row, range.start.col, range.end.col, index as u32));
            }
        }
        entries.sort_by_key(|entry| (entry.0, entry.1));
        let mut rows = BTreeMap::new();
        let mut i = 0;
        while i < entries.len() {
            let mut j = i + 1;
            while j < entries.len() && entries[j].0 == entries[i].0 {
                j += 1;
            }
            let group = &mut entries[i..j];
            let by_col = group.windows(2).all(|w| w[0].2 < w[1].1);
            if !by_col {
                group.sort_by_key(|entry| entry.3);
            }
            rows.insert(
                entries[i].0,
                LinkRow {
                    start: i as u32,
                    end: j as u32,
                    by_col,
                },
            );
            i = j;
        }
        Self {
            links,
            rows,
            entries,
            scanned,
        }
    }

    /// the first link in `hyperlinks` order (if any) that covers `at`.
    fn at(&self, at: CellRef) -> Option<&'a Hyperlink> {
        let in_row = self.rows.get(&at.row).and_then(|row| {
            let entries = &self.entries[row.start as usize..row.end as usize];
            if row.by_col {
                let p = entries.partition_point(|entry| entry.1 <= at.col);
                p.checked_sub(1)
                    .and_then(|i| entries.get(i))
                    .filter(|entry| entry.2 >= at.col)
                    .map(|entry| entry.3)
            } else {
                entries
                    .iter()
                    .find(|entry| entry.1 <= at.col && at.col <= entry.2)
                    .map(|entry| entry.3)
            }
        });
        let scanned = self
            .scanned
            .iter()
            .copied()
            .find(|&i| self.links[i as usize].range.contains(at));
        let index = match (in_row, scanned) {
            (Some(m), Some(s)) => m.min(s),
            (Some(m), None) => m,
            (None, s) => s?,
        };
        Some(&self.links[index as usize])
    }
}

/// viewport-local `(x, y, w, h)` of a cell's box, spanning its merged range
/// when `at` anchors one.
fn cell_box(
    geom: &GridGeometry,
    rows: &AxisLayout,
    cols: &AxisLayout,
    merges: &MergeIndex,
    at: CellRef,
) -> Option<CellBox> {
    let (end_col, end_row) = match merges.covering(at) {
        Some(merge) => (merge.end.col, merge.end.row),
        None => (at.col, at.row),
    };
    let col = cols.span(at.col, end_col, |column| geom.col_x(column))?;
    let row = rows.span(at.row, end_row, |row| geom.row_y(row))?;
    Some(CellBox {
        x: col.raw_start,
        y: row.raw_start,
        w: col.raw_end - col.raw_start,
        h: row.raw_end - row.raw_start,
        clip: Rect {
            x: col.start,
            y: row.start,
            w: col.end - col.start,
            h: row.end - row.start,
        },
    })
}

/// everything the frame resolves once per xf id instead of once per cell.
struct ResolvedXf {
    format_code: Arc<str>,
    font: Option<ResolvedFont>,
    fill: Option<Arc<str>>,
    border: Option<ResolvedBorder>,
    h: Option<HAlign>,
    v: Option<VAlign>,
    /// wrap_text or shrink_to_fit — both suppress text spill.
    wrap_or_shrink: bool,
}

struct ResolvedFont {
    size_pt: Option<f64>,
    family: Option<Arc<str>>,
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
    color: Option<Arc<str>>,
}

struct ResolvedBorder {
    top: Option<ResolvedBorderEdge>,
    left: Option<ResolvedBorderEdge>,
    bottom: Option<ResolvedBorderEdge>,
    right: Option<ResolvedBorderEdge>,
}

struct ResolvedBorderEdge {
    width: f32,
    style: Option<Arc<str>>,
    color: Arc<str>,
}

/// per-frame xf cache: each style id resolves once. sparse map — cell style
/// ids are workbook-controlled `u32`s, so indexing a vec by them is unsafe.
struct FrameStyles<'a> {
    styles: &'a Stylesheet,
    xfs: HashMap<u32, Arc<ResolvedXf>>,
    none: Option<Arc<ResolvedXf>>,
}

impl<'a> FrameStyles<'a> {
    fn new(styles: &'a Stylesheet) -> Self {
        FrameStyles {
            styles,
            xfs: HashMap::new(),
            none: None,
        }
    }

    fn xf(&mut self, style: Option<u32>) -> &Arc<ResolvedXf> {
        let styles = self.styles;
        match style {
            Some(style) => self
                .xfs
                .entry(style)
                .or_insert_with(|| Arc::new(Self::resolve(styles, Some(style)))),
            None => self
                .none
                .get_or_insert_with(|| Arc::new(Self::resolve(styles, None))),
        }
    }

    fn resolve(styles: &Stylesheet, style: Option<u32>) -> ResolvedXf {
        let format_code = match style.map(|s| styles.format_code_for(s)) {
            Some(FormatCode::Custom(code)) => Arc::from(code),
            Some(FormatCode::Builtin(id)) => {
                Arc::from(builtin_format_code(id).unwrap_or("General"))
            }
            None => Arc::from("General"),
        };
        let font = style
            .and_then(|s| styles.font_for(s))
            .map(|font| ResolvedFont {
                size_pt: font.size_pt,
                family: font.name.as_deref().map(Arc::from),
                bold: font.bold,
                italic: font.italic,
                underline: font.underline,
                strike: font.strike,
                color: font
                    .color
                    .as_ref()
                    .and_then(|c| styles.resolve_color(c))
                    .map(Arc::from),
            });
        let fill = style
            .and_then(|s| styles.fill_for(s))
            .and_then(|fill| match fill {
                Fill::Solid(color) => styles.resolve_color(color).map(Arc::from),
                _ => None,
            });
        let border = style
            .and_then(|s| styles.border_for(s))
            .map(|border| ResolvedBorder {
                top: border
                    .top
                    .as_ref()
                    .map(|edge| Self::resolve_edge(styles, edge)),
                left: border
                    .left
                    .as_ref()
                    .map(|edge| Self::resolve_edge(styles, edge)),
                bottom: border
                    .bottom
                    .as_ref()
                    .map(|edge| Self::resolve_edge(styles, edge)),
                right: border
                    .right
                    .as_ref()
                    .map(|edge| Self::resolve_edge(styles, edge)),
            });
        let (h, v, wrap_or_shrink) = style
            .and_then(|s| styles.alignment_for(s))
            .map_or((None, None, false), |a| {
                (a.h, a.v, a.wrap_text || a.shrink_to_fit)
            });
        ResolvedXf {
            format_code,
            font,
            fill,
            border,
            h,
            v,
            wrap_or_shrink,
        }
    }

    fn resolve_edge(styles: &Stylesheet, edge: &BorderEdge) -> ResolvedBorderEdge {
        let (width, style) = border_stroke(edge.style);
        ResolvedBorderEdge {
            width,
            style: style.map(Arc::from),
            color: edge
                .color
                .as_ref()
                .and_then(|c| styles.resolve_color(c))
                .map_or_else(|| Arc::from(BORDER_COLOR), Arc::from),
        }
    }
}

/// display string and resolved font color for a cell, or `None` when it renders
/// nothing. a `[Red]`-style number-format color overrides the font color.
fn cell_display_text(
    xf: &ResolvedXf,
    date_system: xlsx_model::DateSystem,
    cell: &xlsx_model::Cell,
    default_color: &Arc<str>,
) -> Option<(Arc<str>, Arc<str>)> {
    if matches!(cell.value, CellValue::Empty) {
        return None;
    }
    let formatted = format_value(&cell.value, &xf.format_code, date_system);
    if formatted.text.is_empty() {
        return None;
    }
    let color = formatted.color.map_or_else(
        || {
            xf.font
                .as_ref()
                .and_then(|font| font.color.clone())
                .unwrap_or_else(|| default_color.clone())
        },
        Arc::from,
    );
    Some((formatted.text.into(), color))
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
fn resolve_align(h: Option<HAlign>, value: &CellValue) -> Align {
    let type_default = match value {
        CellValue::Number { .. } => Align::Right,
        CellValue::Bool { .. } => Align::Center,
        _ => Align::Left,
    };
    match h {
        Some(HAlign::Left) | Some(HAlign::Fill) | Some(HAlign::Justify) => Align::Left,
        Some(HAlign::Right) => Align::Right,
        Some(HAlign::Center) | Some(HAlign::CenterContinuous) | Some(HAlign::Distributed) => {
            Align::Center
        }
        Some(HAlign::General) | None => type_default,
    }
}

/// Alphabetic baseline in pixels for a point-sized cell font.
fn baseline_y(cy0: f32, ch: f32, size: f32, valign: Option<VAlign>) -> f32 {
    let size_px = size * geometry::PX_PER_PT as f32;
    match valign {
        Some(VAlign::Top) => cy0 + TEXT_PAD_PX + size_px * ASCENT_RATIO,
        Some(VAlign::Center | VAlign::Justify | VAlign::Distributed) => {
            cy0 + (ch + size_px * (ASCENT_RATIO - DESCENT_RATIO)) / 2.0
        }
        Some(VAlign::Bottom) | None => cy0 + ch - TEXT_PAD_PX - size_px * DESCENT_RATIO,
    }
}

fn text_baseline(
    cell: CellBox,
    size: f32,
    valign: Option<VAlign>,
    print: Option<&PrintMetrics>,
) -> f32 {
    let Some(metrics) = print else {
        return baseline_y(cell.y, cell.h, size, valign);
    };
    let ratio = size / metrics.font_size_pt * 96.0 / metrics.dpi;
    let ascent = metrics.font_ascent * ratio;
    let descent = metrics.font_descent * ratio;
    match valign {
        Some(VAlign::Top) => cell.y + ascent,
        Some(VAlign::Center | VAlign::Justify | VAlign::Distributed) => {
            cell.y + (cell.h + ascent - descent) / 2.0
        }
        Some(VAlign::Bottom) | None => cell.y + cell.h - descent,
    }
}

/// emit the set edges of a cell's border. a shared interior edge draws once:
/// the bottom (right) edge is skipped when the neighbor declares its own top (left) edge.
#[allow(clippy::too_many_arguments)]
fn emit_borders(
    commands: &mut Vec<DrawCmd>,
    geom: &GridGeometry,
    rows: &AxisLayout,
    cols: &AxisLayout,
    sheet: &Sheet,
    merges: &MergeIndex,
    fx: &mut FrameStyles,
    at: CellRef,
    border: &ResolvedBorder,
) {
    let Some(cell_box) = cell_box(geom, rows, cols, merges, at) else {
        return;
    };
    let (x, y) = (cell_box.x, cell_box.y);
    let (x2, y2) = (x + cell_box.w, y + cell_box.h);
    let clip = cell_box.clip;
    let (clip_x2, clip_y2) = (clip.x + clip.w, clip.y + clip.h);
    let (end_col, end_row) = match merges.covering(at) {
        Some(m) => (m.end.col, m.end.row),
        None => (at.col, at.row),
    };

    if let Some(edge) = &border.top
        && y >= clip.y
        && y <= clip_y2
    {
        commands.push(border_line(clip.x, y, clip_x2, y, edge));
    }
    if let Some(edge) = &border.left
        && x >= clip.x
        && x <= clip_x2
    {
        commands.push(border_line(x, clip.y, x, clip_y2, edge));
    }
    if let Some(edge) = &border.bottom
        && y2 >= clip.y
        && y2 <= clip_y2
        && !neighbor_edge(sheet, fx, end_row + 1, at.col, |b| b.top.is_some())
    {
        commands.push(border_line(clip.x, y2, clip_x2, y2, edge));
    }
    if let Some(edge) = &border.right
        && x2 >= clip.x
        && x2 <= clip_x2
        && !neighbor_edge(sheet, fx, at.row, end_col + 1, |b| b.left.is_some())
    {
        commands.push(border_line(x2, clip.y, x2, clip_y2, edge));
    }
}

/// true when the cell at `(row, col)` has a border satisfying `pick`.
fn neighbor_edge(
    sheet: &Sheet,
    fx: &mut FrameStyles,
    row: u32,
    col: u32,
    pick: impl Fn(&ResolvedBorder) -> bool,
) -> bool {
    let Some(cell) = sheet.cell(CellRef::new(row, col)) else {
        return false;
    };
    fx.xf(cell.style).border.as_ref().is_some_and(pick)
}

/// one border edge as a `Line`, carrying the resolved width, dash and color.
fn border_line(x1: f32, y1: f32, x2: f32, y2: f32, edge: &ResolvedBorderEdge) -> DrawCmd {
    DrawCmd::Line {
        x1,
        y1,
        x2,
        y2,
        width: edge.width,
        color: edge.color.clone(),
        style: edge.style.clone(),
        clip: None,
    }
}

/// map a border weight to a `(stroke width, dash style)`.
fn border_stroke(style: BorderStyle) -> (f32, Option<&'static str>) {
    match style {
        BorderStyle::Hair => (1.0, Some("dotted")),
        BorderStyle::Thin => (1.0, None),
        BorderStyle::Medium => (2.0, None),
        BorderStyle::Thick => (3.0, None),
        BorderStyle::Dashed => (1.0, Some("dashed")),
        BorderStyle::Dotted => (1.0, Some("dotted")),
        BorderStyle::Double => (1.0, Some("double")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::Hyperlink;
    use xlsx_model::styles::{Alignment, HAlign, Xf};
    use xlsx_model::workbook::{Cell, FreezePane, Sheet};

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
        let dl = build_display_list(&wb, SheetId(0), &vp).unwrap();

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
                DrawCmd::Text { text, clip, .. } if &**text == long => Some(clip),
                _ => None,
            })
            .unwrap();
        let dc = geometry::col_chars_to_px(geometry::DEFAULT_COL_WIDTH_CHARS);
        assert_eq!(long_text.x, dc * 2.0);
        assert_eq!(long_text.w, dc * 5.0);
    }

    /// A neighbour the painter draws is not blank, whatever draws it: it also
    /// paints a hyperlink's own label where the cell carries no text.
    #[test]
    fn spilling_text_stops_at_a_hyperlink_label() {
        let long = "a very long label that overflows its cell";
        let dc = geometry::col_chars_to_px(geometry::DEFAULT_COL_WIDTH_CHARS);
        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: dc * 6.0,
            height: 100.0,
        };
        let clip_of = |sheet: Sheet| {
            let mut wb = Workbook::default();
            wb.sheets.push(sheet);
            build_display_list(&wb, SheetId(0), &vp)
                .unwrap()
                .commands
                .iter()
                .find_map(|command| match command {
                    DrawCmd::Text { text, clip, .. } if &**text == long => Some(clip.w),
                    _ => None,
                })
                .unwrap()
        };

        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 1), text_cell(long));
        sheet.hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("D1:D1").unwrap(),
            external_target: Some("https://example.com".into()),
            location: None,
            tooltip: None,
            display: Some("Open".into()),
        });
        assert_eq!(clip_of(sheet), dc * 2.0);
    }

    /// A frozen split is a pane boundary: Excel does not run text across it,
    /// and `span` clamps a pinned start to the frozen extent, so a spill that
    /// crossed the split would be clipped away from where its text sits.
    #[test]
    fn spilling_text_stops_at_a_frozen_split() {
        let long = "a very long label that overflows its cell";
        let dc = geometry::col_chars_to_px(geometry::DEFAULT_COL_WIDTH_CHARS);
        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: dc * 6.0,
            height: 100.0,
        };
        let mut styles = Stylesheet::default();
        styles.cell_xfs.push(Xf {
            alignment: Some(Alignment {
                h: Some(HAlign::Right),
                ..Alignment::default()
            }),
            ..Xf::default()
        });
        let mut sheet = Sheet::new("Sheet1");
        sheet.freeze_pane = Some(FreezePane {
            rows: 0,
            cols: 1,
            top_left: CellRef::new(0, 1),
        });
        let mut cell = text_cell(long);
        cell.style = Some(0);
        sheet.set_cell(CellRef::new(0, 1), cell);
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);
        wb.styles = styles;
        let clip = build_display_list(&wb, SheetId(0), &vp)
            .unwrap()
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCmd::Text { text, clip, .. } if &**text == long => Some(*clip),
                _ => None,
            })
            .unwrap();
        assert_eq!(clip.x, dc, "the run must not reach into the frozen column");
    }

    #[test]
    fn text_spills_over_blank_neighbours_and_stops_at_the_next_value() {
        let long = "a very long label that overflows its cell";
        let dc = geometry::col_chars_to_px(geometry::DEFAULT_COL_WIDTH_CHARS);
        let vp = Viewport {
            x: 0.0,
            y: 0.0,
            width: dc * 6.0,
            height: 100.0,
        };
        let clip_of = |sheet: Sheet, styles: Stylesheet| {
            let mut wb = Workbook::default();
            wb.sheets.push(sheet);
            wb.styles = styles;
            let clip = build_display_list(&wb, SheetId(0), &vp)
                .unwrap()
                .commands
                .iter()
                .find_map(|command| match command {
                    DrawCmd::Text { text, clip, .. } if &**text == long => Some(*clip),
                    _ => None,
                })
                .unwrap();
            (clip.x, clip.w)
        };

        let mut open = Sheet::new("Sheet1");
        open.set_cell(CellRef::new(0, 1), text_cell(long));
        assert_eq!(clip_of(open, Stylesheet::default()), (dc, dc * 6.0));

        let mut blocked = Sheet::new("Sheet1");
        blocked.set_cell(CellRef::new(0, 1), text_cell(long));
        blocked.set_cell(CellRef::new(0, 3), num_cell(1.0));
        assert_eq!(clip_of(blocked, Stylesheet::default()), (dc, dc * 2.0));

        let mut merged = Sheet::new("Sheet1");
        merged.set_cell(CellRef::new(0, 1), text_cell(long));
        merged.merges.push(CellRange::parse_a1("D1:E1").unwrap());
        assert_eq!(clip_of(merged, Stylesheet::default()), (dc, dc * 2.0));

        let mut wrapped_styles = Stylesheet::default();
        wrapped_styles.cell_xfs.push(Xf {
            alignment: Some(Alignment {
                wrap_text: true,
                ..Alignment::default()
            }),
            ..Xf::default()
        });
        let mut wrapped = Sheet::new("Sheet1");
        wrapped.set_cell(
            CellRef::new(0, 1),
            Cell {
                style: Some(0),
                ..text_cell(long)
            },
        );
        assert_eq!(clip_of(wrapped, wrapped_styles), (dc, dc));

        let mut numeric = Sheet::new("Sheet1");
        numeric.set_cell(CellRef::new(0, 1), num_cell(123_456_789_012_345.0));
        let mut wb = Workbook::default();
        wb.sheets.push(numeric);
        let clip = build_display_list(&wb, SheetId(0), &vp)
            .unwrap()
            .commands
            .iter()
            .find_map(|command| match command {
                DrawCmd::Text { clip, .. } => Some(*clip),
                _ => None,
            })
            .unwrap();
        assert_eq!((clip.x, clip.w), (dc, dc));
    }

    #[test]
    fn renders_hyperlink_indication_and_hit_regions() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(CellRef::new(0, 0), text_cell("Website"));
        sheet.hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("A1:B1").unwrap(),
            external_target: Some("https://example.com".into()),
            location: None,
            tooltip: Some("Open site".into()),
            display: None,
        });
        sheet.hyperlinks.push(Hyperlink {
            range: CellRange::parse_a1("C3").unwrap(),
            external_target: None,
            location: Some("Sheet1!A1".into()),
            tooltip: None,
            display: Some("Jump".into()),
        });
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);

        let dl = build_display_list(
            &wb,
            SheetId(0),
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: 400.0,
                height: 120.0,
            },
        )
        .unwrap();

        assert_eq!(dl.hyperlinks.len(), 2);
        assert_eq!(dl.hyperlinks[0].right, 1);
        assert_eq!(dl.hyperlinks[0].tooltip.as_deref(), Some("Open site"));
        let text = dl
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCmd::Text {
                    text,
                    color,
                    underline,
                    ..
                } => Some((&**text, &**color, *underline)),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert!(text.contains(&("Website", HYPERLINK_COLOR, true)));
        assert!(text.contains(&("Jump", HYPERLINK_COLOR, true)));
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
        let dl = build_display_list(&wb, SheetId(0), &vp).unwrap();

        assert_eq!(dl.grid.start_col, 1);
        assert_eq!(dl.grid.start_row, 2);
        assert_eq!(dl.grid.col_indices, None);
        assert_eq!(dl.grid.row_indices, None);
        assert_eq!(dl.grid.col_offsets.len(), 4);
        assert_eq!(dl.grid.row_offsets.len(), 3);
        assert!((dl.grid.col_offsets[0] - (dc * 1.0 - vp.x)).abs() < 0.01);
        assert!((dl.grid.row_offsets[0] - (dr * 2.0 - vp.y)).abs() < 0.01);
    }

    #[test]
    fn frozen_tracks_stay_pinned_while_the_body_scrolls() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.freeze_pane = Some(FreezePane::new(1, 1, CellRef::new(4, 3)));
        sheet.set_cell(CellRef::new(0, 0), text_cell("pinned"));
        sheet.set_cell(CellRef::new(4, 3), text_cell("body"));
        let mut wb = Workbook::default();
        wb.sheets.push(sheet);
        let dc = geometry::col_chars_to_px(geometry::DEFAULT_COL_WIDTH_CHARS);
        let dr = geometry::row_pt_to_px(geometry::DEFAULT_ROW_HEIGHT_PT);
        let dl = build_display_list(
            &wb,
            SheetId(0),
            &Viewport {
                x: dc * 2.0,
                y: dr * 3.0,
                width: dc * 3.0,
                height: dr * 3.0,
            },
        )
        .unwrap();

        assert_eq!(dl.grid.start_col, 0);
        assert_eq!(dl.grid.start_row, 0);
        assert_eq!(dl.grid.col_indices.as_deref(), Some(&[0, 3, 4, 5][..]));
        assert_eq!(dl.grid.row_indices.as_deref(), Some(&[0, 4, 5, 6][..]));
        assert_eq!(
            dl.grid.col_offsets,
            vec![0.0, dc, dc * 2.0, dc * 3.0, dc * 4.0]
        );
        assert_eq!(
            dl.grid.row_offsets,
            vec![0.0, dr, dr * 2.0, dr * 3.0, dr * 4.0]
        );

        let clips = dl
            .commands
            .iter()
            .filter_map(|command| match command {
                DrawCmd::Text { text, clip, .. } => Some((&**text, *clip)),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(
            clips[0],
            (
                "pinned",
                Rect {
                    x: 0.0,
                    y: 0.0,
                    w: dc,
                    h: dr
                }
            )
        );
        assert_eq!(
            clips[1],
            (
                "body",
                Rect {
                    x: dc,
                    y: dr,
                    w: dc * 3.0,
                    h: dr
                }
            )
        );
        assert!(dl.commands.iter().any(|command| matches!(
            command,
            DrawCmd::Line { x1, x2, width, color, .. }
                if *x1 == dc && *x2 == dc && *width == PANE_DIVIDER_WIDTH
                    && &**color == PANE_DIVIDER_COLOR
        )));
        assert!(dl.commands.iter().any(|command| matches!(
            command,
            DrawCmd::Line { y1, y2, width, color, .. }
                if *y1 == dr && *y2 == dr && *width == PANE_DIVIDER_WIDTH
                    && &**color == PANE_DIVIDER_COLOR
        )));
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
                } => Some((&**text, &**color, *strike, *align)),
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
        let dl = build_display_list_with_ghosts(&wb, SheetId(0), &vp, &[ghost(0, 0, "10", "42")])
            .unwrap();

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
        )
        .unwrap();

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
            .unwrap()
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
        )
        .unwrap();

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
        )
        .unwrap();

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
        assert_eq!(lines[0].1, lines[1].1);
        assert!(lines[0].1 < FONT_SIZE_PT);
        assert!(lines[0].1 >= FONT_SIZE_PT * GHOST_MIN_SCALE);
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
        )
        .unwrap();

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
            build_display_list_with_ghosts(&wb, SheetId(0), &vp, &[ghost(0, 0, "4855", "4855")])
                .unwrap();

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
        )
        .unwrap();
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
        let dl = build_display_list(&wb, SheetId(0), &vp).unwrap();

        let texts: Vec<_> = dl
            .commands
            .iter()
            .filter_map(|c| match c {
                DrawCmd::Text { text, clip, .. } => Some((text.clone(), *clip)),
                _ => None,
            })
            .collect();
        assert_eq!(texts.len(), 1);
        assert_eq!(&*texts[0].0, "merged");
        let dc = geometry::col_chars_to_px(geometry::DEFAULT_COL_WIDTH_CHARS);
        assert!((texts[0].1.w - dc * 2.0).abs() < 0.01);
    }

    #[test]
    fn hyperlink_index_matches_a_linear_scan() {
        let link = |a1: &str, tag: &str| Hyperlink {
            range: CellRange::parse_a1(a1).unwrap(),
            external_target: Some(tag.to_string()),
            location: None,
            tooltip: None,
            display: None,
        };
        let links = vec![
            link("B2:D4", "wide"),
            link("C3", "nested"),
            link("A1:A200", "tall"),
            link("D4:E4", "dup"),
            link("B2:D4", "dup2"),
        ];
        let index = HyperlinkIndex::new(&links);
        for row in 0..6u32 {
            for col in 0..6u32 {
                let at = CellRef::new(row, col);
                let want = links
                    .iter()
                    .find(|l| l.range.contains(at))
                    .and_then(|l| l.external_target.as_deref());
                let got = index.at(at).and_then(|l| l.external_target.as_deref());
                assert_eq!(got, want, "at {at:?}");
            }
        }
    }

    #[test]
    fn hyperlink_index_over_budget_falls_back_to_scan() {
        let link = |a1: &str, tag: &str| Hyperlink {
            range: CellRange::parse_a1(a1).unwrap(),
            external_target: Some(tag.to_string()),
            location: None,
            tooltip: None,
            display: None,
        };
        let mut links = Vec::new();
        let rows = LINK_ROW_CAP - 1;
        for i in 0..(LINK_ENTRY_CAP / rows as usize) as u32 {
            let top = i * rows + 1;
            links.push(link(&format!("A{top}:B{}", top + rows - 1), "in"));
        }
        links.push(link("C1:C5", "spill"));
        let index = HyperlinkIndex::new(&links);
        assert!(index.scanned.contains(&((links.len() - 1) as u32)));
        assert_eq!(
            index
                .at(CellRef::new(0, 2))
                .and_then(|l| l.external_target.as_deref()),
            Some("spill")
        );
        assert_eq!(
            index
                .at(CellRef::new(0, 0))
                .and_then(|l| l.external_target.as_deref()),
            Some("in")
        );
    }

    #[test]
    fn merge_index_matches_linear_scan() {
        let range = |r1: u32, c1: u32, r2: u32, c2: u32| CellRange {
            start: CellRef::new(r1, c1),
            end: CellRef::new(r2, c2),
        };
        let merges = [
            range(2, 0, 65, 3),    // MERGE_ROW_CAP - 1 rows tall: indexed
            range(10, 5, 74, 8),   // MERGE_ROW_CAP rows tall: scanned
            range(30, 1, 40, 2),   // nested inside the indexed merge's rows
            range(50, 6, 60, 7),   // nested inside the scanned merge
            range(100, 0, 50, 5),  // inverted: never covers
            range(200, 0, 210, 2), // below a narrow hint
        ];
        let linear = |at: CellRef| merges.iter().find(|m| m.contains(at)).copied();
        for hint in [0u32..=120, 60..=70, 205..=205, RangeInclusive::new(1, 0)] {
            let index = MergeIndex::new(&merges, hint.clone());
            for row in hint.clone() {
                for col in 0u32..10 {
                    let at = CellRef::new(row, col);
                    assert_eq!(index.covering(at), linear(at), "{at:?} hint {hint:?}");
                }
            }
            for merge in &merges {
                if hint.contains(&merge.start.row) {
                    assert_eq!(
                        index.covering(merge.start),
                        linear(merge.start),
                        "anchor {:?} hint {hint:?}",
                        merge.start
                    );
                }
            }
        }
    }
}
