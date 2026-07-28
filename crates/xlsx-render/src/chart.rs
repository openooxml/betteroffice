use std::fmt;

use ooxml_drawingml::GeometryPathCommand;
use ooxml_drawingml::chart::{
    ChartSpace, PlotChart, PlotOp, PlotRect, PlotSink, chart_aria_label, plot_chart_into,
};
use xlsx_model::chart::{AnchorCell, ChartAnchor, SheetChart};
use xlsx_model::workbook::Sheet;
use xlsx_model::{MAX_COLS, MAX_ROWS};

use crate::Viewport;
use crate::display_list::{Align, ChartA11yAttrs, DrawCmd, PathStroke, Rect};
use crate::geometry::{GridGeometry, emu_to_px};

pub const MAX_CHART_OPS_PER_FRAME: usize = 65_536;

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ResolvedChartAnchor {
    pub rect: PlotRect,
    pub pinned_x: bool,
    pub pinned_y: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AnchorError {
    CellOutOfBounds,
    NegativeOffset,
    NonPositiveExtent,
    InvertedCorners,
    OutsideGrid,
    NonFinite,
}

impl fmt::Display for AnchorError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::CellOutOfBounds => "anchor cell is outside the worksheet grid",
            Self::NegativeOffset => "anchor has a negative cell offset or position",
            Self::NonPositiveExtent => "anchor extent must be positive",
            Self::InvertedCorners => "anchor corners are inverted or coincident",
            Self::OutsideGrid => "anchor extends beyond the worksheet grid",
            Self::NonFinite => "anchor does not resolve to finite pixel geometry",
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RenderError {
    ChartSourceUnavailable {
        part: String,
    },
    ChartPartMissing {
        part: String,
    },
    ChartParseFailed {
        part: String,
    },
    UnsupportedChartFamily {
        part: String,
        family: String,
    },
    UnsupportedChartFeature {
        part: String,
        feature: &'static str,
    },
    InvalidChartAnchor {
        part: String,
        anchor_index: usize,
        error: AnchorError,
    },
    InvalidChartGeometry {
        part: String,
    },
    ChartOpBudgetExceeded {
        max: usize,
    },
    DisplayListAllocationFailed,
}

impl fmt::Display for RenderError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ChartSourceUnavailable { part } => {
                write!(f, "cannot render chart {part}: source part is unavailable")
            }
            Self::ChartPartMissing { part } => {
                write!(f, "cannot render chart {part}: package part is missing")
            }
            Self::ChartParseFailed { part } => {
                write!(f, "cannot render chart {part}: chart data is invalid")
            }
            Self::UnsupportedChartFamily { part, family } => {
                write!(
                    f,
                    "cannot render chart {part}: {family} charts are unsupported"
                )
            }
            Self::UnsupportedChartFeature { part, feature } => {
                write!(f, "cannot render chart {part}: {feature} is unsupported")
            }
            Self::InvalidChartAnchor {
                part,
                anchor_index,
                error,
            } => write!(
                f,
                "cannot render chart {part} at anchor {anchor_index}: {error}"
            ),
            Self::InvalidChartGeometry { part } => {
                write!(f, "cannot render chart {part}: geometry is invalid")
            }
            Self::ChartOpBudgetExceeded { max } => {
                write!(f, "chart rendering exceeds the {max}-operation frame limit")
            }
            Self::DisplayListAllocationFailed => {
                f.write_str("chart display-list allocation failed")
            }
        }
    }
}

impl std::error::Error for RenderError {}

pub fn resolve_chart_anchor(
    anchor: ChartAnchor,
    geometry: &GridGeometry,
    frozen_rows: u32,
    frozen_cols: u32,
) -> Result<ResolvedChartAnchor, AnchorError> {
    let sheet_width = f64::from(geometry.col_x(MAX_COLS));
    let sheet_height = f64::from(geometry.row_y(MAX_ROWS));
    let (rect, pinned_x, pinned_y) = match anchor {
        ChartAnchor::TwoCell { from, to, .. } => {
            let (x1, y1) = cell_position(from, geometry)?;
            let (x2, y2) = cell_position(to, geometry)?;
            if x2 <= x1 || y2 <= y1 {
                return Err(AnchorError::InvertedCorners);
            }
            (
                PlotRect {
                    x: x1,
                    y: y1,
                    w: x2 - x1,
                    h: y2 - y1,
                },
                from.col < frozen_cols,
                from.row < frozen_rows,
            )
        }
        ChartAnchor::OneCell { from, extent } => {
            if extent.cx <= 0 || extent.cy <= 0 {
                return Err(AnchorError::NonPositiveExtent);
            }
            let (x, y) = cell_position(from, geometry)?;
            (
                PlotRect {
                    x,
                    y,
                    w: emu_to_px(extent.cx),
                    h: emu_to_px(extent.cy),
                },
                from.col < frozen_cols,
                from.row < frozen_rows,
            )
        }
        ChartAnchor::Absolute { pos, extent } => {
            if pos.x < 0 || pos.y < 0 {
                return Err(AnchorError::NegativeOffset);
            }
            if extent.cx <= 0 || extent.cy <= 0 {
                return Err(AnchorError::NonPositiveExtent);
            }
            (
                PlotRect {
                    x: emu_to_px(pos.x),
                    y: emu_to_px(pos.y),
                    w: emu_to_px(extent.cx),
                    h: emu_to_px(extent.cy),
                },
                false,
                false,
            )
        }
    };
    if !rect.x.is_finite() || !rect.y.is_finite() || !rect.w.is_finite() || !rect.h.is_finite() {
        return Err(AnchorError::NonFinite);
    }
    let right = rect.x + rect.w;
    let bottom = rect.y + rect.h;
    if !right.is_finite() || !bottom.is_finite() {
        return Err(AnchorError::NonFinite);
    }
    if rect.x < 0.0
        || rect.y < 0.0
        || rect.w <= 0.0
        || rect.h <= 0.0
        || right > sheet_width
        || bottom > sheet_height
    {
        return Err(AnchorError::OutsideGrid);
    }
    Ok(ResolvedChartAnchor {
        rect,
        pinned_x,
        pinned_y,
    })
}

fn cell_position(cell: AnchorCell, geometry: &GridGeometry) -> Result<(f64, f64), AnchorError> {
    if cell.col >= MAX_COLS || cell.row >= MAX_ROWS {
        return Err(AnchorError::CellOutOfBounds);
    }
    if cell.col_off < 0 || cell.row_off < 0 {
        return Err(AnchorError::NegativeOffset);
    }
    Ok((
        f64::from(geometry.col_x(cell.col)) + emu_to_px(cell.col_off),
        f64::from(geometry.row_y(cell.row)) + emu_to_px(cell.row_off),
    ))
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn render_charts<F>(
    sheet: &Sheet,
    geometry: &GridGeometry,
    viewport: &Viewport,
    frozen_rows: u32,
    frozen_cols: u32,
    commands: &mut Vec<DrawCmd>,
    a11y: &mut Vec<ChartA11yAttrs>,
    resolver: &mut F,
) -> Result<(), RenderError>
where
    F: FnMut(&SheetChart) -> Result<ChartSpace, RenderError>,
{
    render_charts_with_budget(
        sheet,
        geometry,
        viewport,
        frozen_rows,
        frozen_cols,
        commands,
        a11y,
        resolver,
        MAX_CHART_OPS_PER_FRAME,
    )
}

#[allow(clippy::too_many_arguments)]
fn render_charts_with_budget<F>(
    sheet: &Sheet,
    geometry: &GridGeometry,
    viewport: &Viewport,
    frozen_rows: u32,
    frozen_cols: u32,
    commands: &mut Vec<DrawCmd>,
    a11y: &mut Vec<ChartA11yAttrs>,
    resolver: &mut F,
    max_ops: usize,
) -> Result<(), RenderError>
where
    F: FnMut(&SheetChart) -> Result<ChartSpace, RenderError>,
{
    let mut remaining = max_ops;
    for chart in &sheet.charts {
        let resolved = resolve_chart_anchor(chart.anchor, geometry, frozen_rows, frozen_cols)
            .map_err(|error| RenderError::InvalidChartAnchor {
                part: chart.part.clone(),
                anchor_index: chart.anchor_index,
                error,
            })?;
        let Some((rect, clip)) =
            viewport_geometry(resolved, geometry, viewport, frozen_rows, frozen_cols)
        else {
            continue;
        };
        if remaining == 0 {
            return Err(RenderError::ChartOpBudgetExceeded { max: max_ops });
        }
        let space = resolver(chart)?;
        if let Some(family) = unsupported_family(&space) {
            return Err(RenderError::UnsupportedChartFamily {
                part: chart.part.clone(),
                family: family.to_string(),
            });
        }
        if let Some(feature) = unsupported_feature(&space) {
            return Err(RenderError::UnsupportedChartFeature {
                part: chart.part.clone(),
                feature,
            });
        }
        let plot = PlotChart::from(&space);
        let label = chart_aria_label(&plot);
        let mut sink = ChartSink {
            commands,
            clip,
            remaining: &mut remaining,
            failure: None,
        };
        plot_chart_into(&plot, rect, &mut sink);
        if let Some(failure) = sink.failure {
            return Err(match failure {
                SinkFailure::Budget => RenderError::ChartOpBudgetExceeded { max: max_ops },
                SinkFailure::Geometry => RenderError::InvalidChartGeometry {
                    part: chart.part.clone(),
                },
                SinkFailure::Allocation => RenderError::DisplayListAllocationFailed,
            });
        }
        a11y.try_reserve(1)
            .map_err(|_| RenderError::DisplayListAllocationFailed)?;
        a11y.push(ChartA11yAttrs { label });
    }
    Ok(())
}

fn unsupported_family(space: &ChartSpace) -> Option<&str> {
    let supported = |family: &str| {
        matches!(
            family,
            "column"
                | "bar"
                | "line"
                | "pie"
                | "doughnut"
                | "area"
                | "scatter"
                | "bubble"
                | "radar"
                | "stock"
                | "ofPie"
                | "surface"
        )
    };
    if space.plot_groups.is_empty() {
        return (!supported(&space.chart_type)).then_some(space.chart_type.as_str());
    }
    space
        .plot_groups
        .iter()
        .filter_map(|group| group.chart_type.as_deref())
        .find(|family| !supported(family))
}

fn unsupported_feature(space: &ChartSpace) -> Option<&'static str> {
    if space.plot_groups.iter().any(|group| {
        matches!(
            group.grouping.as_deref(),
            Some("stacked" | "percentStacked")
        )
    }) {
        return Some("stacked chart grouping");
    }
    if space.axis_list.as_deref().is_some_and(|axes| {
        axes.iter().filter(|axis| axis.axis_type == "value").count() > 1
            || axes
                .iter()
                .filter(|axis| matches!(axis.axis_type.as_str(), "category" | "date"))
                .count()
                > 1
    }) {
        return Some("secondary-axis chart combinations");
    }
    let mut axes = space
        .plot_groups
        .iter()
        .map(|group| group.axis_ids.as_slice())
        .filter(|axis_ids| !axis_ids.is_empty());
    if let Some(first) = axes.next()
        && axes.any(|axis_ids| axis_ids != first)
    {
        return Some("secondary-axis chart combinations");
    }
    if space
        .axis_list
        .iter()
        .flatten()
        .any(|axis| axis.logarithmic_base.is_some())
    {
        return Some("logarithmic chart axes");
    }
    None
}

fn viewport_geometry(
    anchor: ResolvedChartAnchor,
    geometry: &GridGeometry,
    viewport: &Viewport,
    frozen_rows: u32,
    frozen_cols: u32,
) -> Option<(PlotRect, Rect)> {
    let frozen_x = f64::from(geometry.col_x(frozen_cols.min(MAX_COLS)));
    let frozen_y = f64::from(geometry.row_y(frozen_rows.min(MAX_ROWS)));
    let x = if anchor.pinned_x {
        anchor.rect.x
    } else {
        anchor.rect.x - f64::from(viewport.x)
    };
    let y = if anchor.pinned_y {
        anchor.rect.y
    } else {
        anchor.rect.y - f64::from(viewport.y)
    };
    let rect = PlotRect {
        x,
        y,
        w: anchor.rect.w,
        h: anchor.rect.h,
    };
    let clip_x1 = if anchor.pinned_x { 0.0 } else { frozen_x };
    let clip_y1 = if anchor.pinned_y { 0.0 } else { frozen_y };
    let clip_x2 = if anchor.pinned_x {
        frozen_x.min(f64::from(viewport.width))
    } else {
        f64::from(viewport.width)
    };
    let clip_y2 = if anchor.pinned_y {
        frozen_y.min(f64::from(viewport.height))
    } else {
        f64::from(viewport.height)
    };
    let clip = intersect_f64(
        rect,
        PlotRect {
            x: clip_x1,
            y: clip_y1,
            w: clip_x2 - clip_x1,
            h: clip_y2 - clip_y1,
        },
    )?;
    Some((rect, plot_rect_to_rect(clip)?))
}

fn intersect_f64(left: PlotRect, right: PlotRect) -> Option<PlotRect> {
    let x1 = left.x.max(right.x);
    let y1 = left.y.max(right.y);
    let x2 = (left.x + left.w).min(right.x + right.w);
    let y2 = (left.y + left.h).min(right.y + right.h);
    (x2 > x1 && y2 > y1).then_some(PlotRect {
        x: x1,
        y: y1,
        w: x2 - x1,
        h: y2 - y1,
    })
}

fn plot_rect_to_rect(rect: PlotRect) -> Option<Rect> {
    Some(Rect {
        x: finite_f32(rect.x).ok()?,
        y: finite_f32(rect.y).ok()?,
        w: finite_f32(rect.w).ok()?,
        h: finite_f32(rect.h).ok()?,
    })
}

#[derive(Clone, Copy)]
enum SinkFailure {
    Budget,
    Geometry,
    Allocation,
}

struct ChartSink<'a> {
    commands: &'a mut Vec<DrawCmd>,
    clip: Rect,
    remaining: &'a mut usize,
    failure: Option<SinkFailure>,
}

impl PlotSink for ChartSink<'_> {
    fn accepts_more(&mut self) -> bool {
        if self.failure.is_some() {
            return false;
        }
        if *self.remaining == 0 {
            self.failure = Some(SinkFailure::Budget);
            return false;
        }
        true
    }

    fn push_op(&mut self, op: PlotOp) -> bool {
        if self.failure.is_some() {
            return false;
        }
        let Some(next) = self.remaining.checked_sub(1) else {
            self.failure = Some(SinkFailure::Budget);
            return false;
        };
        *self.remaining = next;
        let command = match translate_op(op, self.clip) {
            Ok(command) => command,
            Err(()) => {
                self.failure = Some(SinkFailure::Geometry);
                return false;
            }
        };
        let Some(command) = command else {
            return true;
        };
        if self.commands.try_reserve(1).is_err() {
            self.failure = Some(SinkFailure::Allocation);
            return false;
        }
        self.commands.push(command);
        true
    }
}

fn translate_op(op: PlotOp, chart_clip: Rect) -> Result<Option<DrawCmd>, ()> {
    match op {
        PlotOp::Rect { x, y, w, h, fill } => Ok(Some(DrawCmd::FillRect {
            x: finite_f32(x)?,
            y: finite_f32(y)?,
            w: positive_f32(w)?,
            h: positive_f32(h)?,
            color: fill,
            clip: Some(chart_clip),
        })),
        PlotOp::Text {
            text,
            x,
            baseline_y,
            width,
            font,
            color,
        } => {
            let x = finite_f32(x)?;
            let y = finite_f32(baseline_y)?;
            let width = positive_f32(width)?;
            let font_size_px = positive_f32(font.size_px)?;
            let font_size = positive_f32(font.size_px * 72.0 / 96.0)?;
            let text_clip = text_clip(x, y, width, font_size_px, chart_clip)?;
            Ok(text_clip.map(|clip| DrawCmd::Text {
                x,
                y,
                text,
                font_size,
                color,
                clip,
                align: Align::Left,
                bold: font.weight >= 600,
                italic: false,
                underline: false,
                strike: false,
                highlight: None,
                dashed_underline: false,
                font_family: Some(font.family.to_string()),
                ghost: false,
                chart: true,
            }))
        }
        PlotOp::Line {
            x1,
            y1,
            x2,
            y2,
            color,
            width,
        } => Ok(Some(DrawCmd::Line {
            x1: finite_f32(x1)?,
            y1: finite_f32(y1)?,
            x2: finite_f32(x2)?,
            y2: finite_f32(y2)?,
            width: positive_f32(width)?,
            color,
            style: None,
            clip: Some(chart_clip),
        })),
        PlotOp::Path {
            commands,
            fill,
            stroke,
            ..
        } => {
            validate_path(&commands)?;
            Ok(Some(DrawCmd::Path {
                commands,
                fill,
                stroke: stroke
                    .map(|stroke| {
                        Ok(PathStroke {
                            color: stroke.color,
                            width: positive_f32(stroke.width)?,
                        })
                    })
                    .transpose()?,
                clip: Some(chart_clip),
            }))
        }
    }
}

fn text_clip(
    x: f32,
    baseline_y: f32,
    width: f32,
    font_size_px: f32,
    chart_clip: Rect,
) -> Result<Option<Rect>, ()> {
    let bounds = PlotRect {
        x: f64::from(x),
        y: f64::from(baseline_y - font_size_px),
        w: f64::from(width),
        h: f64::from(font_size_px * 1.25),
    };
    let chart = PlotRect {
        x: f64::from(chart_clip.x),
        y: f64::from(chart_clip.y),
        w: f64::from(chart_clip.w),
        h: f64::from(chart_clip.h),
    };
    let Some(intersection) = intersect_f64(bounds, chart) else {
        return Ok(None);
    };
    plot_rect_to_rect(intersection).map(Some).ok_or(())
}

fn validate_path(commands: &[GeometryPathCommand]) -> Result<(), ()> {
    commands.iter().try_for_each(|command| match command {
        GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
            finite_f32(*x)?;
            finite_f32(*y)?;
            Ok(())
        }
        GeometryPathCommand::Quad { cpx, cpy, x, y } => {
            finite_f32(*cpx)?;
            finite_f32(*cpy)?;
            finite_f32(*x)?;
            finite_f32(*y)?;
            Ok(())
        }
        GeometryPathCommand::Cubic {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        } => {
            finite_f32(*cp1x)?;
            finite_f32(*cp1y)?;
            finite_f32(*cp2x)?;
            finite_f32(*cp2y)?;
            finite_f32(*x)?;
            finite_f32(*y)?;
            Ok(())
        }
        GeometryPathCommand::Close => Ok(()),
    })
}

fn finite_f32(value: f64) -> Result<f32, ()> {
    let value = value as f32;
    value.is_finite().then_some(value).ok_or(())
}

fn positive_f32(value: f64) -> Result<f32, ()> {
    let value = finite_f32(value)?;
    (value > 0.0).then_some(value).ok_or(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use ooxml_drawingml::chart::ChartLegend;
    use xlsx_model::chart::{AnchorEditAs, AnchorExtent, AnchorPos};

    fn geometry() -> GridGeometry {
        let mut sheet = Sheet::new("Sheet1");
        sheet.col_widths.insert(1, 0.0);
        sheet.row_heights.insert(1, 0.0);
        GridGeometry::new(&sheet)
    }

    fn cell(col: u32, col_off: i64, row: u32, row_off: i64) -> AnchorCell {
        AnchorCell {
            col,
            col_off,
            row,
            row_off,
        }
    }

    #[test]
    fn resolves_all_anchor_variants_with_hidden_tracks() {
        let geometry = geometry();
        let two = resolve_chart_anchor(
            ChartAnchor::TwoCell {
                from: cell(0, 9_525, 0, 9_525),
                to: cell(2, 0, 2, 0),
                edit_as: AnchorEditAs::TwoCell,
            },
            &geometry,
            0,
            0,
        )
        .unwrap();
        let one = resolve_chart_anchor(
            ChartAnchor::OneCell {
                from: cell(2, 0, 2, 0),
                extent: AnchorExtent {
                    cx: 952_500,
                    cy: 476_250,
                },
            },
            &geometry,
            0,
            0,
        )
        .unwrap();
        let absolute = resolve_chart_anchor(
            ChartAnchor::Absolute {
                pos: AnchorPos {
                    x: 95_250,
                    y: 190_500,
                },
                extent: AnchorExtent {
                    cx: 952_500,
                    cy: 476_250,
                },
            },
            &geometry,
            0,
            0,
        )
        .unwrap();

        assert_eq!(two.rect.x, 1.0);
        assert_eq!(two.rect.y, 1.0);
        assert!(two.rect.w > 0.0);
        assert!(two.rect.h > 0.0);
        assert_eq!((one.rect.w, one.rect.h), (100.0, 50.0));
        assert_eq!((absolute.rect.x, absolute.rect.y), (10.0, 20.0));
    }

    #[test]
    fn rejects_degenerate_and_off_grid_anchors() {
        let geometry = geometry();
        let same = cell(1, 0, 1, 0);
        assert_eq!(
            resolve_chart_anchor(
                ChartAnchor::TwoCell {
                    from: same,
                    to: same,
                    edit_as: AnchorEditAs::TwoCell,
                },
                &geometry,
                0,
                0,
            ),
            Err(AnchorError::InvertedCorners)
        );
        assert_eq!(
            resolve_chart_anchor(
                ChartAnchor::OneCell {
                    from: cell(0, 0, 0, 0),
                    extent: AnchorExtent { cx: 0, cy: 1 },
                },
                &geometry,
                0,
                0,
            ),
            Err(AnchorError::NonPositiveExtent)
        );
        assert_eq!(
            resolve_chart_anchor(
                ChartAnchor::Absolute {
                    pos: AnchorPos { x: i64::MAX, y: 0 },
                    extent: AnchorExtent { cx: 1, cy: 1 },
                },
                &geometry,
                0,
                0,
            ),
            Err(AnchorError::OutsideGrid)
        );
    }

    #[test]
    fn frame_budget_fails_on_the_first_op_past_the_boundary() {
        let mut commands = Vec::new();
        let mut remaining = 1;
        let mut sink = ChartSink {
            commands: &mut commands,
            clip: Rect {
                x: 0.0,
                y: 0.0,
                w: 10.0,
                h: 10.0,
            },
            remaining: &mut remaining,
            failure: None,
        };
        let op = || PlotOp::Rect {
            x: 0.0,
            y: 0.0,
            w: 1.0,
            h: 1.0,
            fill: "#000000".into(),
        };
        assert!(sink.push_op(op()));
        assert!(sink.failure.is_none());
        assert!(!sink.push_op(op()));
        assert!(matches!(sink.failure, Some(SinkFailure::Budget)));
        assert_eq!(commands.len(), 1);
    }

    #[test]
    fn frame_budget_is_reported_before_resolving_the_next_chart() {
        let mut sheet = Sheet::new("Sheet1");
        let anchor = ChartAnchor::Absolute {
            pos: AnchorPos { x: 0, y: 0 },
            extent: AnchorExtent {
                cx: 952_500,
                cy: 952_500,
            },
        };
        for index in 0..2 {
            sheet.charts.push(SheetChart {
                part: format!("xl/charts/chart{}.xml", index + 1),
                drawing: "xl/drawings/drawing1.xml".into(),
                anchor_index: index,
                anchor,
                refs: Vec::new(),
            });
        }
        let geometry = GridGeometry::new(&sheet);
        let mut commands = Vec::new();
        let mut a11y = Vec::new();
        let resolved = std::cell::Cell::new(0);
        let mut resolver = |_: &SheetChart| {
            resolved.set(resolved.get() + 1);
            Ok(ChartSpace {
                chart_type: "line".into(),
                title: None,
                legend: Some(ChartLegend {
                    position: None,
                    visible: false,
                    ..Default::default()
                }),
                series: Vec::new(),
                axes: None,
                plot_groups: Vec::new(),
                axis_list: None,
                ..Default::default()
            })
        };
        let error = render_charts_with_budget(
            &sheet,
            &geometry,
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 200.0,
            },
            0,
            0,
            &mut commands,
            &mut a11y,
            &mut resolver,
            1,
        )
        .unwrap_err();

        assert_eq!(error, RenderError::ChartOpBudgetExceeded { max: 1 });
        assert_eq!(resolved.get(), 1);
        assert_eq!(commands.len(), 1);
    }

    #[test]
    fn unsupported_families_are_not_substituted() {
        let space = ChartSpace {
            chart_type: "treemap".into(),
            title: None,
            legend: None,
            series: Vec::new(),
            axes: None,
            plot_groups: Vec::new(),
            axis_list: None,
            ..Default::default()
        };
        assert_eq!(unsupported_family(&space), Some("treemap"));

        let area = ChartSpace {
            chart_type: "area".into(),
            ..Default::default()
        };
        assert_eq!(unsupported_family(&area), None);
    }
}
