use std::fmt;
use std::sync::Arc;

use hashbrown::{HashMap, HashSet};
use ooxml_drawingml::GeometryPathCommand;
use ooxml_drawingml::chart::{
    ChartSpace, MAX_PLOT_DATA_SCAN, PlotChart, PlotOp, PlotRect, PlotSeries, PlotSink,
    PlotTextAlign, chart_aria_label, plot_chart_into,
};
use xlsx_model::chart::{AnchorCell, ChartAnchor, SheetChart};
use xlsx_model::styles::Stylesheet;
use xlsx_model::workbook::Sheet;
use xlsx_model::{MAX_COLS, MAX_ROWS};

use crate::Viewport;
use crate::display_list::{Align, ChartRegion, DrawCmd, PathStroke, Rect};
use crate::geometry::{EMU_PER_PT, GridGeometry, PX_PER_PT, emu_to_px};

pub const MAX_CHART_OPS_PER_FRAME: usize = 65_536;

const PLACEHOLDER_FILL: &str = "#f2f2f2";
const PLACEHOLDER_BORDER: &str = "#bfbfbf";
const PLACEHOLDER_BORDER_WIDTH: f32 = 1.0;

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
    /// no longer raised: a geometry failure degrades that one chart instead.
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

impl RenderError {
    /// Whether the whole frame must be refused. Everything else is a defect in
    /// one chart and degrades to a placeholder, leaving the sheet renderable.
    /// Governs only errors handed back by the resolver and by the family and
    /// feature checks; sink failures, the placeholder's own budget and the a11y
    /// reservation each decide their disposition where they occur, so adding a
    /// variant here does not change what those paths do.
    pub fn refuses_frame(&self) -> bool {
        matches!(
            self,
            Self::ChartOpBudgetExceeded { .. } | Self::DisplayListAllocationFailed
        )
    }
}

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

/// The anchor `anchor` would carry after sliding `dx`/`dy` viewport pixels,
/// clamped so it never leaves the grid. `None` for an absolute anchor, which
/// is pinned to the sheet and whose position a save cannot rewrite.
pub fn moved_chart_anchor(
    anchor: ChartAnchor,
    geometry: &GridGeometry,
    dx: f64,
    dy: f64,
) -> Option<ChartAnchor> {
    if matches!(anchor, ChartAnchor::Absolute { .. }) || !dx.is_finite() || !dy.is_finite() {
        return None;
    }
    let resolved = resolve_chart_anchor(anchor, geometry, 0, 0).ok()?;
    let rect = resolved.rect;
    let sheet_width = f64::from(geometry.col_x(MAX_COLS));
    let sheet_height = f64::from(geometry.row_y(MAX_ROWS));
    let x = (rect.x + dx).clamp(0.0, (sheet_width - rect.w).max(0.0));
    let y = (rect.y + dy).clamp(0.0, (sheet_height - rect.h).max(0.0));
    if !x.is_finite() || !y.is_finite() {
        return None;
    }
    let from = anchor_cell_at(x, y, geometry)?;
    match anchor {
        ChartAnchor::TwoCell { edit_as, .. } => Some(ChartAnchor::TwoCell {
            from,
            to: anchor_cell_at(x + rect.w, y + rect.h, geometry)?,
            edit_as,
        }),
        ChartAnchor::OneCell { extent, .. } => Some(ChartAnchor::OneCell { from, extent }),
        ChartAnchor::Absolute { .. } => None,
    }
}

/// The grid-anchored corner a sheet-pixel point falls in: the cell containing
/// it plus the EMU offset from that cell's top-left.
fn anchor_cell_at(x: f64, y: f64, geometry: &GridGeometry) -> Option<AnchorCell> {
    let col = geometry.col_at_x(finite_f32(x).ok()?).min(MAX_COLS - 1);
    let row = geometry.row_at_y(finite_f32(y).ok()?).min(MAX_ROWS - 1);
    Some(AnchorCell {
        col,
        col_off: px_to_emu(x - f64::from(geometry.col_x(col)))?,
        row,
        row_off: px_to_emu(y - f64::from(geometry.row_y(row)))?,
    })
}

/// Pixels back to EMU, rounded to the nearest unit and floored at zero — an
/// anchor offset the schema and the writer both accept.
fn px_to_emu(px: f64) -> Option<i64> {
    let emu = (px / PX_PER_PT * EMU_PER_PT).round();
    (emu.is_finite() && emu <= i64::MAX as f64).then(|| (emu as i64).max(0))
}

/// Every chart the viewport shows, in paint order, carrying the geometry the
/// display list publishes. The label is left empty: resolving anchors needs no
/// chart part, and a hit test has no use for one.
pub fn chart_regions(
    sheet: &Sheet,
    styles: &Stylesheet,
    viewport: &Viewport,
) -> Result<Vec<ChartRegion>, RenderError> {
    let geometry = GridGeometry::new(sheet, styles);
    chart_regions_with_geometry(sheet, &geometry, viewport)
}

#[doc(hidden)]
pub fn chart_regions_with_geometry(
    sheet: &Sheet,
    geometry: &GridGeometry,
    viewport: &Viewport,
) -> Result<Vec<ChartRegion>, RenderError> {
    let (frozen_rows, frozen_cols) = sheet
        .freeze_pane
        .map_or((0, 0), |pane| (pane.rows, pane.cols));
    Ok(
        visible_charts(sheet, geometry, viewport, frozen_rows, frozen_cols)?
            .into_iter()
            .map(|visible| visible.region)
            .collect(),
    )
}

/// a chart the viewport shows: which one, where its plot goes, and the region
/// the display list publishes for it.
struct VisibleChart<'a> {
    chart: &'a SheetChart,
    plot: PlotRect,
    region: ChartRegion,
}

fn visible_charts<'a>(
    sheet: &'a Sheet,
    geometry: &GridGeometry,
    viewport: &Viewport,
    frozen_rows: u32,
    frozen_cols: u32,
) -> Result<Vec<VisibleChart<'a>>, RenderError> {
    let mut visible = Vec::new();
    for chart in &sheet.charts {
        let resolved = match resolve_chart_anchor(chart.anchor, geometry, frozen_rows, frozen_cols)
        {
            Ok(resolved) => resolved,
            Err(error) => {
                let error = RenderError::InvalidChartAnchor {
                    part: chart.part.clone(),
                    anchor_index: chart.anchor_index,
                    error,
                };
                if error.refuses_frame() {
                    return Err(error);
                }
                // no resolved rect, so there is nowhere to put a placeholder
                // and nothing to address: this chart is not in the frame.
                continue;
            }
        };
        let Some((plot, clip)) =
            viewport_geometry(resolved, geometry, viewport, frozen_rows, frozen_cols)
        else {
            continue;
        };
        // an unrepresentable rect leaves nothing to paint or point at either.
        let Some(rect) = plot_rect_to_rect(plot) else {
            continue;
        };
        visible
            .try_reserve(1)
            .map_err(|_| RenderError::DisplayListAllocationFailed)?;
        visible.push(VisibleChart {
            chart,
            plot,
            region: ChartRegion {
                id: chart.frame_id(),
                label: String::new(),
                placeholder: false,
                rect,
                clip,
                movable: !matches!(chart.anchor, ChartAnchor::Absolute { .. }),
            },
        });
    }
    Ok(visible)
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn render_charts<F, R>(
    sheet: &Sheet,
    geometry: &GridGeometry,
    viewport: &Viewport,
    frozen_rows: u32,
    frozen_cols: u32,
    commands: &mut Vec<DrawCmd>,
    regions: &mut Vec<ChartRegion>,
    resolver: &mut F,
) -> Result<(), RenderError>
where
    F: FnMut(&SheetChart) -> Result<R, RenderError>,
    R: Into<Arc<ChartSpace>>,
{
    render_charts_with_budget(
        sheet,
        geometry,
        viewport,
        frozen_rows,
        frozen_cols,
        commands,
        regions,
        resolver,
        MAX_CHART_OPS_PER_FRAME,
    )
}

#[allow(clippy::too_many_arguments)]
fn render_charts_with_budget<F, R>(
    sheet: &Sheet,
    geometry: &GridGeometry,
    viewport: &Viewport,
    frozen_rows: u32,
    frozen_cols: u32,
    commands: &mut Vec<DrawCmd>,
    regions: &mut Vec<ChartRegion>,
    resolver: &mut F,
    max_ops: usize,
) -> Result<(), RenderError>
where
    F: FnMut(&SheetChart) -> Result<R, RenderError>,
    R: Into<Arc<ChartSpace>>,
{
    let mut remaining = max_ops;
    for visible in visible_charts(sheet, geometry, viewport, frozen_rows, frozen_cols)? {
        let VisibleChart {
            chart,
            plot: rect,
            mut region,
        } = visible;
        let clip = region.clip;
        if remaining == 0 {
            return Err(RenderError::ChartOpBudgetExceeded { max: max_ops });
        }
        let (label, placeholder) = match plot_one_chart(
            chart,
            rect,
            clip,
            commands,
            &mut remaining,
            max_ops,
            resolver,
        ) {
            ChartOutcome::Drawn(label) => (label, false),
            ChartOutcome::Degraded(label) => {
                paint_chart_placeholder(commands, rect, clip, &mut remaining, max_ops)?;
                (label, true)
            }
            ChartOutcome::Fatal(error) => return Err(error),
        };
        region.label = label;
        region.placeholder = placeholder;
        regions
            .try_reserve(1)
            .map_err(|_| RenderError::DisplayListAllocationFailed)?;
        regions.push(region);
    }
    Ok(())
}

/// What one chart contributed to the frame.
enum ChartOutcome {
    Drawn(String),
    Degraded(String),
    Fatal(RenderError),
}

#[allow(clippy::too_many_arguments)]
fn plot_one_chart<F, R>(
    chart: &SheetChart,
    rect: PlotRect,
    clip: Rect,
    commands: &mut Vec<DrawCmd>,
    remaining: &mut usize,
    max_ops: usize,
    resolver: &mut F,
) -> ChartOutcome
where
    F: FnMut(&SheetChart) -> Result<R, RenderError>,
    R: Into<Arc<ChartSpace>>,
{
    let space: Arc<ChartSpace> = match resolver(chart) {
        Ok(space) => space.into(),
        Err(error) if error.refuses_frame() => return ChartOutcome::Fatal(error),
        Err(_) => return ChartOutcome::Degraded(degraded_label(None)),
    };
    let mut plot = PlotChart::from(space.as_ref());
    let label = chart_aria_label(&plot);
    if let Some(error) = chart_refusal(chart, &space) {
        if error.refuses_frame() {
            return ChartOutcome::Fatal(error);
        }
        let known = match &error {
            RenderError::UnsupportedChartFamily { family, .. } => {
                unsupported_family_label(&space, family)
            }
            _ => label,
        };
        return ChartOutcome::Degraded(degraded_label(Some(&known)));
    }
    if needed_main_gate(&space).is_some() {
        prepare_axis_references(&mut plot);
    }
    let fallback_x = scatter_fallback_x(&plot);
    prepare_scatter_plot(&mut plot, &fallback_x);
    let baseline = commands.len();
    let mut sink = ChartSink {
        commands: &mut *commands,
        clip,
        remaining: &mut *remaining,
        failure: None,
    };
    plot_chart_into(&plot, rect, &mut sink);
    let failure = sink.failure;
    match failure {
        None => ChartOutcome::Drawn(label),
        Some(SinkFailure::Budget) => {
            ChartOutcome::Fatal(RenderError::ChartOpBudgetExceeded { max: max_ops })
        }
        Some(SinkFailure::Allocation) => {
            ChartOutcome::Fatal(RenderError::DisplayListAllocationFailed)
        }
        Some(SinkFailure::Geometry) => {
            commands.truncate(baseline);
            ChartOutcome::Degraded(degraded_label(Some(&label)))
        }
    }
}

fn scatter_fallback_x(plot: &PlotChart<'_>) -> Vec<f64> {
    let count = plot
        .series
        .iter()
        .filter(|_| plot.plot_groups.is_empty() && plot.chart_type == "scatter")
        .chain(
            plot.plot_groups
                .iter()
                .filter(|group| group.chart_type.unwrap_or(plot.chart_type) == "scatter")
                .flat_map(|group| &group.series),
        )
        .filter(|series| !series.x_values.iter().any(|value| value.is_finite()))
        .map(|series| series.values.len())
        .max()
        .unwrap_or(0)
        .min(MAX_PLOT_DATA_SCAN);
    (1..=count).map(|value| value as f64).collect()
}

fn prepare_scatter_plot<'a>(plot: &mut PlotChart<'a>, fallback_x: &'a [f64]) {
    if plot.plot_groups.is_empty() && plot.chart_type == "scatter" {
        apply_scatter_x(&mut plot.series, fallback_x);
    }
    for group in &mut plot.plot_groups {
        let family = group.chart_type.unwrap_or(plot.chart_type);
        if matches!(family, "scatter" | "bubble")
            && group.axis_ids.len() == 2
            && plot.axes.iter().any(|axis| {
                axis.id == group.axis_ids.first().copied()
                    && matches!(axis.position, Some("left" | "right"))
            })
        {
            group.axis_ids.swap(0, 1);
        }
        if family == "scatter" {
            apply_scatter_x(&mut group.series, fallback_x);
        }
    }
}

fn apply_scatter_x<'a>(series: &mut [PlotSeries<'a>], fallback_x: &'a [f64]) {
    for series in series {
        if !series.x_values.iter().any(|value| value.is_finite()) {
            series.x_values = &fallback_x[..series.values.len().min(fallback_x.len())];
        }
    }
}

fn chart_refusal(chart: &SheetChart, space: &ChartSpace) -> Option<RenderError> {
    if let Some(family) = unsupported_family(space) {
        return Some(RenderError::UnsupportedChartFamily {
            part: chart.part.clone(),
            family: family.to_string(),
        });
    }
    unsupported_feature(space).map(|feature| RenderError::UnsupportedChartFeature {
        part: chart.part.clone(),
        feature,
    })
}

/// Screen-reader label for a chart the frame could not draw.
fn degraded_label(known: Option<&str>) -> String {
    format!("{}, not shown", known.unwrap_or("Chart"))
}

/// `chart_aria_label` announces any family it does not recognize as a column
/// chart, so a refused family names itself rather than borrowing that guess.
fn unsupported_family_label(space: &ChartSpace, family: &str) -> String {
    let title = space
        .title
        .as_deref()
        .filter(|title| !title.is_empty())
        .unwrap_or("Untitled chart");
    format!("{title}, {family} chart")
}

/// Neutral box in the chart's own rect, so an undrawable chart still occupies
/// its space instead of taking the rest of the frame down with it.
fn paint_chart_placeholder(
    commands: &mut Vec<DrawCmd>,
    rect: PlotRect,
    clip: Rect,
    remaining: &mut usize,
    max_ops: usize,
) -> Result<(), RenderError> {
    let (Ok(x), Ok(y), Ok(w), Ok(h)) = (
        finite_f32(rect.x),
        finite_f32(rect.y),
        positive_f32(rect.w),
        positive_f32(rect.h),
    ) else {
        return Ok(());
    };
    let mut push = |command: DrawCmd| -> Result<(), RenderError> {
        let Some(next) = remaining.checked_sub(1) else {
            return Err(RenderError::ChartOpBudgetExceeded { max: max_ops });
        };
        *remaining = next;
        commands
            .try_reserve(1)
            .map_err(|_| RenderError::DisplayListAllocationFailed)?;
        commands.push(command);
        Ok(())
    };
    push(DrawCmd::FillRect {
        x,
        y,
        w,
        h,
        color: PLACEHOLDER_FILL.into(),
        clip: Some(clip),
    })?;
    if w <= PLACEHOLDER_BORDER_WIDTH || h <= PLACEHOLDER_BORDER_WIDTH {
        return Ok(());
    }
    let inset = PLACEHOLDER_BORDER_WIDTH / 2.0;
    let (left, top) = (x + inset, y + inset);
    let (right, bottom) = (x + w - inset, y + h - inset);
    for (x1, y1, x2, y2) in [
        (left, top, right, top),
        (right, top, right, bottom),
        (right, bottom, left, bottom),
        (left, bottom, left, top),
    ] {
        push(DrawCmd::Line {
            x1,
            y1,
            x2,
            y2,
            width: PLACEHOLDER_BORDER_WIDTH,
            color: PLACEHOLDER_BORDER.into(),
            style: None,
            rail: None,
            clip: Some(clip),
        })?;
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

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum AxisReference<'a> {
    Integer(i128),
    Text(&'a str),
}

fn axis_reference(id: &str) -> AxisReference<'_> {
    let id = id.trim();
    match id.parse() {
        Ok(value) => AxisReference::Integer(value),
        Err(_) => AxisReference::Text(id),
    }
}

fn prepare_axis_references(plot: &mut PlotChart<'_>) {
    let by_id: HashMap<_, _> = plot
        .axes
        .iter()
        .filter_map(|axis| axis.id)
        .map(|id| (axis_reference(id), id))
        .collect();
    for group in &mut plot.plot_groups {
        for id in &mut group.axis_ids {
            if let Some(resolved) = by_id.get(&axis_reference(id)) {
                *id = resolved;
            }
        }
    }
}

fn unsupported_feature(space: &ChartSpace) -> Option<&'static str> {
    needed_main_gate(space)?;
    let axes = space.axis_list.as_deref().unwrap_or_default();
    let mut by_id = HashMap::new();
    for axis in axes {
        if let Some(id) = axis.id.as_deref() {
            *by_id.entry(axis_reference(id)).or_insert(0usize) += 1;
        }
    }
    for group in &space.plot_groups {
        let mut seen = HashSet::new();
        for id in &group.axis_ids {
            let id = axis_reference(id);
            if !seen.insert(id) {
                return Some("ambiguous chart axis references");
            }
            match by_id.get(&id) {
                Some(1) => {}
                Some(_) => return Some("ambiguous chart axis references"),
                None => return Some("unresolved chart axis references"),
            }
        }
    }
    None
}

fn needed_main_gate(space: &ChartSpace) -> Option<&'static str> {
    if space.plot_groups.iter().any(|group| {
        matches!(
            group.grouping.as_deref(),
            Some("stacked" | "percentStacked")
        )
    }) {
        return Some("stacked chart grouping");
    }
    let primary_xy_axes = has_primary_xy_axes(space);
    if !primary_xy_axes
        && space.axis_list.as_deref().is_some_and(|axes| {
            axes.iter().filter(|axis| axis.axis_type == "value").count() > 1
                || axes
                    .iter()
                    .filter(|axis| matches!(axis.axis_type.as_str(), "category" | "date"))
                    .count()
                    > 1
        })
    {
        return Some("secondary-axis chart combinations");
    }
    let mut axes = space
        .plot_groups
        .iter()
        .map(|group| group.axis_ids.as_slice())
        .filter(|axis_ids| !axis_ids.is_empty());
    if !primary_xy_axes
        && let Some(first) = axes.next()
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

fn has_primary_xy_axes(space: &ChartSpace) -> bool {
    let Some(axes) = space.axis_list.as_deref() else {
        return false;
    };
    if axes.len() != 2 || axes.iter().any(|axis| axis.axis_type != "value") {
        return false;
    }
    let Some(horizontal) = axes
        .iter()
        .find(|axis| matches!(axis.position.as_deref(), Some("bottom" | "top")))
    else {
        return false;
    };
    let Some(vertical) = axes
        .iter()
        .find(|axis| matches!(axis.position.as_deref(), Some("left" | "right")))
    else {
        return false;
    };
    let (Some(x_id), Some(y_id)) = (horizontal.id.as_deref(), vertical.id.as_deref()) else {
        return false;
    };
    if x_id.is_empty()
        || y_id.is_empty()
        || x_id == y_id
        || horizontal.cross_axis_id.as_deref() != Some(y_id)
        || vertical.cross_axis_id.as_deref() != Some(x_id)
    {
        return false;
    }
    space.plot_groups.len() == 1
        && space.plot_groups.iter().all(|group| {
            matches!(
                group.chart_type.as_deref().unwrap_or(&space.chart_type),
                "scatter" | "bubble"
            ) && group.axis_ids.len() == 2
                && ((group.axis_ids[0] == x_id && group.axis_ids[1] == y_id)
                    || (group.axis_ids[0] == y_id && group.axis_ids[1] == x_id))
        })
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

    fn anchors_text_end(&self) -> bool {
        true
    }

    fn push_end_aligned_text(&mut self, op: PlotOp) -> bool {
        self.push_anchored(op, true)
    }

    fn push_op(&mut self, op: PlotOp) -> bool {
        self.push_anchored(op, false)
    }
}

impl ChartSink<'_> {
    fn push_anchored(&mut self, op: PlotOp, end_aligned: bool) -> bool {
        if self.failure.is_some() {
            return false;
        }
        let Some(next) = self.remaining.checked_sub(1) else {
            self.failure = Some(SinkFailure::Budget);
            return false;
        };
        *self.remaining = next;
        let command = match translate_op(op, self.clip, end_aligned) {
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

fn translate_op(op: PlotOp, chart_clip: Rect, end_aligned: bool) -> Result<Option<DrawCmd>, ()> {
    match op {
        PlotOp::Rect { x, y, w, h, fill } => Ok(Some(DrawCmd::FillRect {
            x: finite_f32(x)?,
            y: finite_f32(y)?,
            w: positive_f32(w)?,
            h: positive_f32(h)?,
            color: fill.into(),
            clip: Some(chart_clip),
        })),
        PlotOp::Text {
            text,
            x,
            baseline_y,
            width,
            font,
            color,
            align,
            rotation_deg: _,
        } => {
            let x = finite_f32(x)?;
            let y = finite_f32(baseline_y)?;
            let width = positive_f32(width)?;
            let font_size_px = positive_f32(font.size_px)?;
            let font_size = positive_f32(font.size_px * 72.0 / 96.0)?;
            let text_clip = text_clip(x, y, width, font_size_px, end_aligned, chart_clip)?;
            let (x, align) = if end_aligned {
                (x + width, Align::Right)
            } else {
                match align {
                    PlotTextAlign::Center => (x + width / 2.0, Align::Center),
                    PlotTextAlign::Start => (x, Align::Left),
                }
            };
            Ok(text_clip.map(|clip| DrawCmd::Text {
                x,
                y,
                text: text.into(),
                font_size,
                color: color.into(),
                clip,
                align,
                bold: font.weight >= 600,
                italic: false,
                underline: false,
                strike: false,
                highlight: None,
                dashed_underline: false,
                font_family: Some(font.family.into()),
                number_format: None,
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
            color: color.into(),
            style: None,
            rail: None,
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
                fill: fill.into(),
                stroke: stroke
                    .map(|stroke| {
                        Ok(PathStroke {
                            color: stroke.color.into(),
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
    end_aligned: bool,
    chart_clip: Rect,
) -> Result<Option<Rect>, ()> {
    let (clip_x, clip_width) = if end_aligned {
        (
            f64::from(chart_clip.x),
            f64::from(x + width) - f64::from(chart_clip.x),
        )
    } else {
        (f64::from(x), f64::from(width))
    };
    let bounds = PlotRect {
        x: clip_x,
        y: f64::from(baseline_y - font_size_px),
        w: clip_width,
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
    use ooxml_drawingml::chart::{
        ChartAxis, ChartDataLabels, ChartLegend, ChartManualLayout, ChartMarker, ChartPlotGroup,
        ChartSeries,
    };
    use xlsx_model::chart::{AnchorEditAs, AnchorExtent, AnchorPos};

    fn geometry() -> GridGeometry {
        let mut sheet = Sheet::new("Sheet1");
        sheet.col_widths.insert(1, 0.0);
        sheet.row_heights.insert(1, 0.0);
        GridGeometry::new(&sheet, &Stylesheet::default())
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
        let geometry = GridGeometry::new(&sheet, &Stylesheet::default());
        let mut commands = Vec::new();
        let mut a11y = Vec::new();
        let resolved = std::cell::Cell::new(0);
        let mut resolver = |_: &SheetChart| {
            resolved.set(resolved.get() + 1);
            Ok(Arc::new(ChartSpace {
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
            }))
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

    #[test]
    fn only_resource_guards_refuse_the_frame() {
        let part = || "xl/charts/chart1.xml".to_string();
        for error in [
            RenderError::ChartSourceUnavailable { part: part() },
            RenderError::ChartPartMissing { part: part() },
            RenderError::ChartParseFailed { part: part() },
            RenderError::UnsupportedChartFamily {
                part: part(),
                family: "treemap".into(),
            },
            RenderError::UnsupportedChartFeature {
                part: part(),
                feature: "ambiguous chart axis references",
            },
            RenderError::InvalidChartAnchor {
                part: part(),
                anchor_index: 0,
                error: AnchorError::OutsideGrid,
            },
            RenderError::InvalidChartGeometry { part: part() },
        ] {
            assert!(!error.refuses_frame(), "{error} must degrade one chart");
        }
        assert!(RenderError::ChartOpBudgetExceeded { max: 1 }.refuses_frame());
        assert!(RenderError::DisplayListAllocationFailed.refuses_frame());
    }

    #[test]
    fn a_placeholder_cannot_outrun_the_op_budget() {
        let mut sheet = Sheet::new("Sheet1");
        sheet.charts.push(SheetChart {
            part: "xl/charts/chart1.xml".into(),
            drawing: "xl/drawings/drawing1.xml".into(),
            anchor_index: 0,
            anchor: ChartAnchor::Absolute {
                pos: AnchorPos { x: 0, y: 0 },
                extent: AnchorExtent {
                    cx: 952_500,
                    cy: 952_500,
                },
            },
            refs: Vec::new(),
        });
        let geometry = GridGeometry::new(&sheet, &Stylesheet::default());
        let viewport = Viewport {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 200.0,
        };
        let mut resolver = |chart: &SheetChart| {
            Err::<ChartSpace, _>(RenderError::ChartParseFailed {
                part: chart.part.clone(),
            })
        };
        let mut render = |max_ops: usize| {
            let mut commands = Vec::new();
            let mut a11y = Vec::new();
            render_charts_with_budget(
                &sheet,
                &geometry,
                &viewport,
                0,
                0,
                &mut commands,
                &mut a11y,
                &mut resolver,
                max_ops,
            )
            .map(|()| (commands.len(), a11y.len()))
        };

        assert_eq!(
            render(3).unwrap_err(),
            RenderError::ChartOpBudgetExceeded { max: 3 }
        );
        assert_eq!(render(5).unwrap(), (5, 1));
    }

    fn charted_sheet(anchor: ChartAnchor) -> Sheet {
        let mut sheet = Sheet::new("Sheet1");
        sheet.charts.push(SheetChart {
            part: "xl/charts/chart1.xml".into(),
            drawing: "xl/drawings/drawing1.xml".into(),
            anchor_index: 0,
            anchor,
            refs: Vec::new(),
        });
        sheet
    }

    #[test]
    fn axis_label_regression_currency_ticks_fit_wide_fallback_fonts() {
        for chart_type in ["column", "line"] {
            let space = ChartSpace {
                chart_type: chart_type.into(),
                plot_groups: vec![ChartPlotGroup {
                    chart_type: Some(chart_type.into()),
                    axis_ids: vec!["value".into()],
                    series: vec![ChartSeries {
                        categories: vec!["Q1".into(), "Q2".into()],
                        values: vec![200_000.0, 600_000.0],
                        ..Default::default()
                    }],
                    ..Default::default()
                }],
                axis_list: Some(vec![ChartAxis {
                    id: Some("value".into()),
                    axis_type: "value".into(),
                    position: Some("left".into()),
                    min: Some(0.0),
                    max: Some(600_000.0),
                    number_format: Some("\"$\"#,##0".into()),
                    ..Default::default()
                }]),
                ..Default::default()
            };
            let (commands, _) = render_scatter(&space);
            let plot_left = commands
                .iter()
                .filter_map(|command| match command {
                    DrawCmd::Line { x1, x2, width, .. } if x1 == x2 && *width == 1.0 => Some(*x1),
                    _ => None,
                })
                .fold(f32::INFINITY, f32::min);
            let (x, clip, align, size) = commands
                .iter()
                .find_map(|command| match command {
                    DrawCmd::Text {
                        text,
                        x,
                        clip,
                        align,
                        font_size,
                        ..
                    } if text.as_ref() == "$600,000" => {
                        Some((*x, *clip, *align, *font_size * 96.0 / 72.0))
                    }
                    _ => None,
                })
                .expect("the currency tick is drawn");
            let anchor = plot_left - (7.0 + 0.6 * size);
            let wide_width = "$600,000".chars().count() as f32 * 0.64 * size;
            assert!(
                clip.x <= anchor - wide_width && clip.x + clip.w >= anchor - 1e-4,
                "{chart_type}: clip {clip:?} must span {wide_width} left of {anchor}"
            );
            assert_eq!(align, Align::Right);
            assert!((x - anchor).abs() < 1e-4, "{chart_type}: {x} vs {anchor}");
            assert!(clip.x >= 0.0 && clip.y >= 0.0);
            assert!(clip.x + clip.w <= 400.0 && clip.y + clip.h <= 300.0);
        }
    }

    #[test]
    fn axis_label_regression_billion_currency_ticks_fit_wide_fallback_fonts() {
        let label = "$1,000,000,000";
        for chart_type in ["column", "line"] {
            let space = ChartSpace {
                chart_type: chart_type.into(),
                plot_groups: vec![ChartPlotGroup {
                    chart_type: Some(chart_type.into()),
                    axis_ids: vec!["value".into()],
                    series: vec![ChartSeries {
                        categories: vec!["Q1".into(), "Q2".into()],
                        values: vec![200_000_000.0, 1_000_000_000.0],
                        ..Default::default()
                    }],
                    ..Default::default()
                }],
                axis_list: Some(vec![ChartAxis {
                    id: Some("value".into()),
                    axis_type: "value".into(),
                    position: Some("left".into()),
                    min: Some(0.0),
                    max: Some(1_000_000_000.0),
                    number_format: Some("\"$\"#,##0".into()),
                    ..Default::default()
                }]),
                ..Default::default()
            };
            let (commands, _) = render_scatter(&space);
            let plot_left = commands
                .iter()
                .filter_map(|command| match command {
                    DrawCmd::Line { x1, x2, width, .. } if x1 == x2 && *width == 1.0 => Some(*x1),
                    _ => None,
                })
                .fold(f32::INFINITY, f32::min);
            let (x, clip, align, size) = commands
                .iter()
                .find_map(|command| match command {
                    DrawCmd::Text {
                        text,
                        x,
                        clip,
                        align,
                        font_size,
                        ..
                    } if text.as_ref() == label => {
                        Some((*x, *clip, *align, *font_size * 96.0 / 72.0))
                    }
                    _ => None,
                })
                .expect("the billion currency tick is drawn");
            let anchor = plot_left - (7.0 + 0.6 * size);
            let wide_width = label.chars().count() as f32 * 0.64 * size;
            assert!(
                clip.x <= anchor - wide_width && clip.x + clip.w >= anchor - 1e-4,
                "{chart_type}: clip {clip:?} must span {wide_width} left of {anchor}"
            );
            assert_eq!(align, Align::Right);
            assert!((x - anchor).abs() < 1e-4, "{chart_type}: {x} vs {anchor}");
            assert!((clip.x + clip.w - anchor).abs() < 1e-4);
            assert!(clip.x >= 0.0 && clip.y >= 0.0);
            assert!(clip.x + clip.w <= 400.0 && clip.y + clip.h <= 300.0);
        }
    }

    #[test]
    fn axis_label_regression_start_aligned_legend_keeps_box_clip() {
        let space = ChartSpace {
            legend: Some(ChartLegend {
                position: Some("bottom".into()),
                visible: true,
                ..Default::default()
            }),
            ..scatter_space()
        };
        let frame = PlotRect {
            x: 0.0,
            y: 0.0,
            w: 400.0,
            h: 300.0,
        };
        let mut ops = Vec::new();
        plot_chart_into(&PlotChart::from(&space), frame, &mut ops);
        let op = ops
            .into_iter()
            .find(|op| matches!(op, PlotOp::Text { text, .. } if text == "Series1"))
            .expect("the legend label is drawn");
        let PlotOp::Text {
            x,
            baseline_y,
            width,
            ref font,
            align,
            ..
        } = op
        else {
            unreachable!();
        };
        assert_eq!(align, PlotTextAlign::Start);
        let bounds = Rect {
            x: x as f32,
            y: baseline_y as f32 - font.size_px as f32,
            w: width as f32,
            h: font.size_px as f32 * 1.25,
        };
        let chart_clip = Rect {
            x: 0.0,
            y: 0.0,
            w: 400.0,
            h: 300.0,
        };
        assert!(bounds.x > chart_clip.x);
        assert!(bounds.x + bounds.w < chart_clip.x + chart_clip.w);
        assert!(bounds.y >= chart_clip.y);
        assert!(bounds.y + bounds.h <= chart_clip.y + chart_clip.h);
        let DrawCmd::Text { x, clip, align, .. } =
            translate_op(op, chart_clip, false).unwrap().unwrap()
        else {
            unreachable!();
        };
        assert_eq!(align, Align::Left);
        assert_eq!(x, bounds.x);
        assert_eq!(clip, bounds);
    }

    #[test]
    fn axis_label_regression_bar_categories_fit_wide_fallback_fonts() {
        let label = "Northern region";
        let space = ChartSpace {
            chart_type: "bar".into(),
            plot_groups: vec![ChartPlotGroup {
                chart_type: Some("bar".into()),
                series: vec![ChartSeries {
                    categories: vec![label.into(), "South".into()],
                    values: vec![200_000.0, 600_000.0],
                    ..Default::default()
                }],
                ..Default::default()
            }],
            ..Default::default()
        };
        let (commands, _) = render_scatter(&space);
        let plot_left = commands
            .iter()
            .filter_map(|command| match command {
                DrawCmd::Line { x1, x2, width, .. } if x1 == x2 && *width == 1.0 => Some(*x1),
                _ => None,
            })
            .fold(f32::INFINITY, f32::min);
        let (x, clip, align, size) = commands
            .iter()
            .find_map(|command| match command {
                DrawCmd::Text {
                    text,
                    x,
                    clip,
                    align,
                    font_size,
                    ..
                } if text.as_ref() == label => Some((*x, *clip, *align, *font_size * 96.0 / 72.0)),
                _ => None,
            })
            .expect("the long category label is drawn");
        let anchor = plot_left - (7.0 + 0.6 * size);
        let wide_width = label.chars().count() as f32 * 0.64 * size;
        assert!(
            clip.x <= anchor - wide_width && clip.x + clip.w >= anchor - 1e-4,
            "clip {clip:?} must span {wide_width} left of {anchor}"
        );
        assert_eq!(align, Align::Right);
        assert!((x - anchor).abs() < 1e-4, "{x} vs {anchor}");
        assert!(clip.x >= 0.0 && clip.y >= 0.0);
        assert!(clip.x + clip.w <= 400.0 && clip.y + clip.h <= 300.0);
    }

    fn scatter_space() -> ChartSpace {
        let series = ChartSeries {
            name: Some("Series1".into()),
            values: vec![5.0, 10.0, 15.0],
            x_values: Some(vec![1.0, 9.0, 4.0]),
            color: "#123456".into(),
            category_formula: Some("Sheet1!$A$1:$A$3".into()),
            value_formula: Some("Sheet1!$B$1:$B$3".into()),
            marker: Some(ChartMarker {
                symbol: Some("square".into()),
                size: Some(6.0),
                ..Default::default()
            }),
            ..Default::default()
        };
        ChartSpace {
            chart_type: "scatter".into(),
            series: vec![series.clone()],
            plot_groups: vec![ChartPlotGroup {
                chart_type: Some("scatter".into()),
                axis_ids: vec!["1".into(), "2".into()],
                scatter_style: Some("marker".into()),
                series: vec![series],
                ..Default::default()
            }],
            axis_list: Some(vec![
                ChartAxis {
                    id: Some("1".into()),
                    axis_type: "value".into(),
                    position: Some("bottom".into()),
                    cross_axis_id: Some("2".into()),
                    min: Some(0.0),
                    max: Some(10.0),
                    ..Default::default()
                },
                ChartAxis {
                    id: Some("2".into()),
                    axis_type: "value".into(),
                    position: Some("left".into()),
                    cross_axis_id: Some("1".into()),
                    min: Some(0.0),
                    max: Some(20.0),
                    ..Default::default()
                },
            ]),
            plot_layout: Some(ChartManualLayout {
                x: 0.2,
                y: 0.2,
                w: 0.6,
                h: 0.6,
            }),
            ..Default::default()
        }
    }

    fn render_scatter(space: &ChartSpace) -> (Vec<DrawCmd>, Vec<ChartRegion>) {
        let sheet = charted_sheet(ChartAnchor::Absolute {
            pos: AnchorPos { x: 0, y: 0 },
            extent: AnchorExtent {
                cx: 3_810_000,
                cy: 2_857_500,
            },
        });
        let geometry = GridGeometry::new(&sheet, &Stylesheet::default());
        let mut commands = Vec::new();
        let mut regions = Vec::new();
        let mut resolver = |_: &SheetChart| Ok::<_, RenderError>(space.clone());
        render_charts(
            &sheet,
            &geometry,
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: 400.0,
                height: 300.0,
            },
            0,
            0,
            &mut commands,
            &mut regions,
            &mut resolver,
        )
        .unwrap();
        assert_eq!(regions.len(), 1);
        (commands, regions)
    }

    fn scatter_centers(commands: &[DrawCmd]) -> Vec<(f32, f32)> {
        commands
            .iter()
            .filter_map(|command| match command {
                DrawCmd::FillRect {
                    x, y, w, h, color, ..
                } if color.as_ref() == "#123456" => Some((x + w / 2.0, y + h / 2.0)),
                _ => None,
            })
            .collect()
    }

    fn assert_scatter_centers(commands: &[DrawCmd], expected: &[(f32, f32)]) {
        let centers = scatter_centers(commands);
        assert_eq!(centers.len(), expected.len());
        for ((x, y), (expected_x, expected_y)) in centers.iter().zip(expected) {
            assert!((x - expected_x).abs() < 0.01, "x: {x} != {expected_x}");
            assert!((y - expected_y).abs() < 0.01, "y: {y} != {expected_y}");
        }
    }

    #[test]
    fn scatter_crossing_value_axes_plot_numeric_xy() {
        let (commands, regions) = render_scatter(&scatter_space());
        assert!(!regions[0].placeholder);
        assert_scatter_centers(&commands, &[(104.0, 195.0), (296.0, 150.0), (176.0, 105.0)]);
    }

    #[test]
    fn scatter_auto_axes_reuse_numeric_bounds() {
        let mut space = scatter_space();
        for axis in space.axis_list.as_mut().unwrap() {
            axis.min = None;
            axis.max = None;
        }
        let (commands, regions) = render_scatter(&space);
        assert!(!regions[0].placeholder);
        assert_scatter_centers(&commands, &[(80.0, 183.75), (320.0, 127.5), (170.0, 71.25)]);
    }

    #[test]
    fn scatter_missing_or_non_numeric_x_uses_one_based_indexes() {
        for x_values in [None, Some(Vec::new()), Some(vec![f64::NAN; 3])] {
            let mut space = scatter_space();
            space.plot_groups[0].series[0].x_values = x_values;
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder);
            assert_scatter_centers(&commands, &[(104.0, 195.0), (128.0, 150.0), (152.0, 105.0)]);
        }
    }

    #[test]
    fn scatter_numeric_x_gaps_do_not_shift_point_pairs() {
        let mut space = scatter_space();
        space.plot_groups[0].series[0].x_values = Some(vec![1.0, f64::NAN, 4.0]);
        let (commands, regions) = render_scatter(&space);
        assert!(!regions[0].placeholder);
        assert_scatter_centers(&commands, &[(104.0, 195.0), (176.0, 105.0)]);
    }

    #[test]
    fn scatter_axis_order_and_orientation_preserve_numeric_scaling() {
        let mut space = scatter_space();
        space.plot_groups[0].axis_ids.reverse();
        let axes = space.axis_list.as_mut().unwrap();
        axes[0].position = Some("top".into());
        axes[1].position = Some("right".into());
        axes[0].reversed = true;
        axes[1].reversed = true;
        axes.reverse();
        let (commands, regions) = render_scatter(&space);
        assert!(!regions[0].placeholder);
        assert_scatter_centers(&commands, &[(296.0, 105.0), (104.0, 150.0), (224.0, 195.0)]);
    }

    #[test]
    fn scatter_styles_control_connecting_lines_and_markers() {
        for (style, lines, markers) in [
            ("marker", 0, 3),
            ("lineMarker", 2, 3),
            ("line", 2, 0),
            ("smoothMarker", 2, 3),
            ("smooth", 2, 0),
        ] {
            let mut space = scatter_space();
            space.plot_groups[0].scatter_style = Some(style.into());
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder);
            let series_lines: Vec<_> = commands
                .iter()
                .filter_map(|command| match command {
                    DrawCmd::Line {
                        x1,
                        y1,
                        x2,
                        y2,
                        color,
                        ..
                    } if color.as_ref() == "#123456" => Some((*x1, *y1, *x2, *y2)),
                    _ => None,
                })
                .collect();
            assert_eq!(series_lines.len(), lines, "{style}");
            assert_eq!(scatter_centers(&commands).len(), markers, "{style}");
            if lines > 0 {
                assert_eq!(series_lines[0], (104.0, 195.0, 296.0, 150.0));
            }
        }
    }

    #[test]
    fn scatter_default_auto_and_none_markers_follow_series_settings() {
        for (marker, expected) in [
            (None, 3),
            (
                Some(ChartMarker {
                    symbol: Some("auto".into()),
                    ..Default::default()
                }),
                3,
            ),
            (
                Some(ChartMarker {
                    symbol: Some("none".into()),
                    ..Default::default()
                }),
                0,
            ),
        ] {
            let mut space = scatter_space();
            space.plot_groups[0].series[0].marker = marker;
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder);
            let markers = commands
                .iter()
                .filter(|command| {
                    matches!(command, DrawCmd::Path { fill, .. } if fill.as_ref() == "#123456")
                })
                .count();
            assert_eq!(markers, expected);
        }
    }

    #[test]
    fn scatter_multiple_series_reuse_labels_and_x_fallback() {
        let mut space = scatter_space();
        let group = &mut space.plot_groups[0];
        group.data_labels = Some(ChartDataLabels {
            show_series_name: Some(true),
            ..Default::default()
        });
        let mut second = group.series[0].clone();
        second.name = Some("Series2".into());
        second.values = vec![10.0, 5.0];
        second.x_values = None;
        group.series.push(second);
        let (commands, regions) = render_scatter(&space);
        assert!(!regions[0].placeholder);
        assert_scatter_centers(
            &commands,
            &[
                (104.0, 195.0),
                (296.0, 150.0),
                (176.0, 105.0),
                (104.0, 150.0),
                (128.0, 195.0),
            ],
        );
        for (name, expected) in [("Series1", 3), ("Series2", 2)] {
            assert_eq!(
                commands
                    .iter()
                    .filter(|command| {
                        matches!(command, DrawCmd::Text { text, .. } if text.as_ref() == name)
                    })
                    .count(),
                expected
            );
        }
    }

    #[test]
    fn scatter_pair_is_admitted_in_either_reference_order() {
        let mut space = scatter_space();
        space.plot_groups[0].axis_ids.reverse();
        let (commands, regions) = render_scatter(&space);
        assert!(!regions[0].placeholder);
        assert_scatter_centers(&commands, &[(104.0, 195.0), (296.0, 150.0), (176.0, 105.0)]);
    }

    fn cartesian_space(family: &str, values: Vec<f64>, max: f64) -> ChartSpace {
        let mut space = scatter_space();
        space.chart_type = family.into();
        let group = &mut space.plot_groups[0];
        group.chart_type = Some(family.into());
        group.series[0].categories = (0..values.len()).map(|index| format!("C{index}")).collect();
        group.series[0].values = values;
        group.series[0].x_values = None;
        let axes = space.axis_list.as_mut().unwrap();
        axes[0].axis_type = "category".into();
        axes[0].min = None;
        axes[0].max = None;
        axes[1].max = Some(max);
        space
    }

    fn series_rects(commands: &[DrawCmd], series_color: &str) -> Vec<(f32, f32, f32, f32)> {
        commands
            .iter()
            .filter_map(|command| match command {
                DrawCmd::FillRect {
                    x, y, w, h, color, ..
                } if color.as_ref() == series_color => Some((*x, *y, *w, *h)),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn stacked_columns_place_the_second_segment_above_the_first() {
        let mut space = cartesian_space("column", vec![5.0], 20.0);
        let group = &mut space.plot_groups[0];
        group.grouping = Some("stacked".into());
        let mut second = group.series[0].clone();
        second.values = vec![10.0];
        second.color = "#654321".into();
        group.series.push(second);
        assert_eq!(unsupported_feature(&space), None);
        let (commands, regions) = render_scatter(&space);
        assert!(!regions[0].placeholder);
        let first = series_rects(&commands, "#123456");
        let second = series_rects(&commands, "#654321");
        assert_eq!((first.len(), second.len()), (1, 1));
        let (x, y, w, h) = first[0];
        let (x2, y2, w2, h2) = second[0];
        assert!((x - x2).abs() < 0.01 && (w - w2).abs() < 0.01);
        assert!((y2 + h2 - y).abs() < 0.01);
        assert!((h - 45.0).abs() < 0.01 && (h2 - 90.0).abs() < 0.01);
        assert!((y + h - 240.0).abs() < 0.01);
    }

    #[test]
    fn percent_stacked_bars_normalize_each_category_to_the_full_width() {
        let mut space = cartesian_space("bar", vec![1.0, 2.0], 1.0);
        let group = &mut space.plot_groups[0];
        group.grouping = Some("percentStacked".into());
        let mut second = group.series[0].clone();
        second.values = vec![3.0, 2.0];
        second.color = "#654321".into();
        group.series.push(second);
        assert_eq!(unsupported_feature(&space), None);
        let (commands, regions) = render_scatter(&space);
        assert!(!regions[0].placeholder);
        let first = series_rects(&commands, "#123456");
        let second = series_rects(&commands, "#654321");
        assert_eq!((first.len(), second.len()), (2, 2));
        for ((first, second), expected_width) in first.iter().zip(&second).zip([60.0, 120.0]) {
            let (x, y, w, h) = *first;
            let (x2, y2, w2, h2) = *second;
            assert!((y - y2).abs() < 0.01 && (h - h2).abs() < 0.01);
            assert!((x - 80.0).abs() < 0.01 && (x + w - x2).abs() < 0.01);
            assert!((w - expected_width).abs() < 0.01);
            assert!((w + w2 - 240.0).abs() < 0.01);
        }
        assert!(commands.iter().any(|command| {
            matches!(command, DrawCmd::Text { text, .. } if text.as_ref() == "100%")
        }));
    }

    #[test]
    fn bar_line_combo_scales_the_line_against_the_right_value_axis() {
        let mut space = cartesian_space("column", vec![10.0, 10.0], 20.0);
        let mut line = space.plot_groups[0].clone();
        line.chart_type = Some("line".into());
        line.axis_ids = vec!["3".into(), "4".into()];
        line.series[0].color = "#654321".into();
        space.plot_groups.push(line);
        let axes = space.axis_list.as_mut().unwrap();
        let mut category = axes[0].clone();
        category.id = Some("3".into());
        category.cross_axis_id = Some("4".into());
        let mut right = axes[1].clone();
        right.id = Some("4".into());
        right.cross_axis_id = Some("3".into());
        right.position = Some("right".into());
        right.max = Some(100.0);
        axes.extend([category, right]);
        assert_eq!(unsupported_feature(&space), None);
        let (commands, regions) = render_scatter(&space);
        assert!(!regions[0].placeholder);
        let bars = series_rects(&commands, "#123456");
        let markers = series_rects(&commands, "#654321");
        assert_eq!((bars.len(), markers.len()), (2, 2));
        for (bar, marker) in bars.iter().zip(markers) {
            assert!((bar.1 - 150.0).abs() < 0.01 && (bar.3 - 90.0).abs() < 0.01);
            assert!((marker.1 + marker.3 / 2.0 - 222.0).abs() < 0.01);
        }
        assert!(commands.iter().any(|command| {
            matches!(command, DrawCmd::Text { text, x, align: Align::Left, .. }
                if text.as_ref() == "100" && *x > 320.0)
        }));
    }

    #[test]
    fn logarithmic_value_axis_has_power_ticks_and_evenly_spaced_decades() {
        let mut space = cartesian_space("line", vec![1.0, 10.0, 100.0, 1000.0], 2000.0);
        let axis = &mut space.axis_list.as_mut().unwrap()[1];
        axis.min = Some(1.0);
        axis.logarithmic_base = Some(10.0);
        assert_eq!(unsupported_feature(&space), None);
        let (commands, regions) = render_scatter(&space);
        assert!(!regions[0].placeholder);
        let markers = scatter_centers(&commands);
        assert_eq!(markers.len(), 4);
        for ((_, y), expected_y) in markers.iter().zip([240.0, 185.473, 130.945, 76.418]) {
            assert!((y - expected_y).abs() < 0.01);
        }
        let ticks: Vec<_> = commands
            .iter()
            .filter_map(|command| match command {
                DrawCmd::Text {
                    text,
                    align: Align::Right,
                    ..
                } => Some(text.as_ref()),
                _ => None,
            })
            .collect();
        assert_eq!(ticks, ["1000", "100", "10", "1"]);
    }

    #[test]
    fn logarithmic_axis_with_nonpositive_minimum_keeps_the_linear_fallback() {
        for min in [0.0, -10.0] {
            let mut space = cartesian_space("line", vec![0.0, 10.0, 100.0], 100.0);
            let axis = &mut space.axis_list.as_mut().unwrap()[1];
            axis.min = Some(min);
            axis.major_unit = Some(10.0);
            axis.logarithmic_base = Some(10.0);
            assert_eq!(unsupported_feature(&space), None);
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder);
            let markers = scatter_centers(&commands);
            assert_eq!(markers.len(), 3);
            for ((_, y), value) in markers.iter().zip([0.0, 10.0, 100.0]) {
                let expected_y = 240.0 - 180.0 * (value - min) / (100.0 - min);
                assert!((f64::from(*y) - expected_y).abs() < 0.01);
            }
            assert!(commands.iter().any(|command| {
                matches!(command, DrawCmd::Text { text, .. } if text.as_ref() == "20")
            }));
        }
    }

    #[test]
    fn scatter_groups_sharing_the_pair_render_both_groups() {
        for chart_type in ["scatter", "bubble"] {
            let mut space = scatter_space();
            let mut second = space.plot_groups[0].clone();
            second.chart_type = Some(chart_type.into());
            second.axis_ids.reverse();
            second.series[0].bubble_sizes = Some(vec![1.0, 2.0, 3.0]);
            space.plot_groups.push(second);
            assert_eq!(unsupported_feature(&space), None, "{chart_type}");
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder, "{chart_type}");
            if chart_type == "scatter" {
                let expected = [(104.0, 195.0), (296.0, 150.0), (176.0, 105.0)];
                assert_scatter_centers(&commands, &expected.repeat(2));
            } else {
                assert_scatter_centers(
                    &commands,
                    &[(104.0, 195.0), (296.0, 150.0), (176.0, 105.0)],
                );
                assert_eq!(
                    commands
                        .iter()
                        .filter(|command| {
                            matches!(command, DrawCmd::Path { fill, .. } if fill.as_ref() == "#123456")
                        })
                        .count(),
                    3
                );
            }
        }
    }

    #[test]
    fn scatter_auto_y_bounds_include_values_without_an_x_as_on_main() {
        for (x_values, values) in [
            (vec![1.0, 9.0], vec![5.0, 10.0, 1e9]),
            (vec![1.0, f64::NAN, 9.0], vec![5.0, 1e9, 10.0]),
        ] {
            let mut space = scatter_space();
            for axis in space.axis_list.as_mut().unwrap() {
                axis.min = None;
                axis.max = None;
            }
            space.plot_groups[0].series[0].x_values = Some(x_values);
            space.plot_groups[0].series[0].values = values;
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder);
            assert_scatter_centers(
                &commands,
                &[
                    (80.0, (240.0 - 180.0 * (5.0 / 1.2e9)) as f32),
                    (320.0, (240.0 - 180.0 * (10.0 / 1.2e9)) as f32),
                ],
            );
            assert!(commands.iter().any(|command| {
                matches!(command, DrawCmd::Text { text, .. } if text.as_ref() == "1200000000")
            }));
        }
    }

    #[test]
    fn scatter_constant_auto_x_keeps_its_value_on_the_axis() {
        for (x, label, center_x) in [(40.0, "40", 320.0), (-40.0, "-40", 80.0)] {
            let mut space = scatter_space();
            let axis = &mut space.axis_list.as_mut().unwrap()[0];
            axis.min = None;
            axis.max = None;
            space.plot_groups[0].series[0].x_values = Some(vec![x; 3]);
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder);
            assert_scatter_centers(
                &commands,
                &[(center_x, 195.0), (center_x, 150.0), (center_x, 105.0)],
            );
            assert!(
                commands
                    .iter()
                    .any(|command| matches!(command, DrawCmd::Text { text, .. } if text.as_ref() == label)),
                "{label}"
            );
        }
    }

    #[test]
    fn newly_admitted_groups_can_use_geometry_axis_fallbacks() {
        for family in [
            "column", "bar", "line", "area", "radar", "stock", "surface", "pie", "doughnut",
            "ofPie",
        ] {
            let mut space = scatter_space();
            space.chart_type = family.into();
            space.plot_groups[0].chart_type = Some(family.into());
            assert!(needed_main_gate(&space).is_some(), "{family}");
            assert_eq!(unsupported_feature(&space), None, "{family}");
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder, "{family}");
            assert!(!commands.is_empty(), "{family}");
        }
    }

    #[test]
    fn main_admitted_categorical_axis_references_keep_rendering() {
        for case in 0..5 {
            let mut space = cartesian_space("column", vec![5.0], 20.0);
            match case {
                0 => space.plot_groups[0].axis_ids.clear(),
                1 => space.axis_list = None,
                2 => {
                    space.plot_groups[0].axis_ids = vec!["1".into(), "1".into(), "2".into()];
                }
                3 => space.plot_groups[0].axis_ids = vec!["1".into()],
                _ => {
                    space.plot_groups[0].axis_ids = vec!["1".into(), "missing".into()];
                }
            }
            assert_eq!(needed_main_gate(&space), None, "case {case}");
            assert_eq!(unsupported_feature(&space), None, "case {case}");
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder, "case {case}");
            assert_eq!(series_rects(&commands, "#123456").len(), 1);
        }
    }

    #[test]
    fn radar_with_axes_and_no_group_references_keeps_rendering() {
        let mut space = cartesian_space("radar", vec![5.0, 10.0, 15.0], 20.0);
        space.plot_groups[0].axis_ids.clear();
        assert_eq!(needed_main_gate(&space), None);
        assert_eq!(unsupported_feature(&space), None);
        let (commands, regions) = render_scatter(&space);
        assert!(!regions[0].placeholder);
        assert!(commands.iter().any(|command| {
            matches!(command, DrawCmd::Line { color, .. } if color.as_ref() == "#123456")
        }));
    }

    #[test]
    fn scatter_category_value_and_single_value_axes_keep_rendering() {
        for single_value in [false, true] {
            let mut space = scatter_space();
            let axes = space.axis_list.as_mut().unwrap();
            if single_value {
                axes.remove(0);
                space.plot_groups[0].axis_ids = vec!["2".into()];
            } else {
                axes[0].axis_type = "category".into();
            }
            assert_eq!(needed_main_gate(&space), None);
            assert_eq!(unsupported_feature(&space), None);
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder);
            assert_eq!(scatter_centers(&commands).len(), 3);
        }
    }

    #[test]
    fn surface_with_a_series_axis_keeps_rendering() {
        let mut space = cartesian_space("surface", vec![5.0, 10.0, 15.0], 20.0);
        space.axis_list.as_mut().unwrap().push(ChartAxis {
            id: Some("3".into()),
            axis_type: "series".into(),
            ..Default::default()
        });
        space.plot_groups[0].axis_ids.push("3".into());
        assert_eq!(needed_main_gate(&space), None);
        assert_eq!(unsupported_feature(&space), None);
        let (commands, regions) = render_scatter(&space);
        assert!(!regions[0].placeholder);
        assert!(!commands.is_empty());
    }

    #[test]
    fn stacked_date_axes_and_hidden_axes_render() {
        for hidden in [false, true] {
            let mut space = cartesian_space("column", vec![5.0], 20.0);
            space.plot_groups[0].grouping = Some("stacked".into());
            let axes = space.axis_list.as_mut().unwrap();
            axes[0].axis_type = "date".into();
            axes[1].hidden = hidden;
            assert!(needed_main_gate(&space).is_some());
            assert_eq!(unsupported_feature(&space), None);
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder);
            let bars = series_rects(&commands, "#123456");
            assert_eq!(bars.len(), 1);
            assert!((bars[0].3 - 45.0).abs() < 0.01);
        }
    }

    #[test]
    fn stacked_charts_ignore_unreferenced_duplicate_axis_ids() {
        for hidden in [false, true] {
            let mut space = cartesian_space("column", vec![5.0], 20.0);
            space.plot_groups[0].grouping = Some("stacked".into());
            let axes = space.axis_list.as_mut().unwrap();
            let mut leftover = axes[1].clone();
            leftover.id = Some("unused".into());
            leftover.hidden = hidden;
            axes.extend([leftover.clone(), leftover]);
            assert_eq!(unsupported_feature(&space), None);
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder);
            let bars = series_rects(&commands, "#123456");
            assert_eq!(bars.len(), 1);
            assert!((bars[0].3 - 45.0).abs() < 0.01);
        }
    }

    #[test]
    fn newly_admitted_charts_resolve_numeric_and_trimmed_axis_ids() {
        for (definition, reference) in [("01002", "1002"), ("1002", "01002"), (" value ", "value")]
        {
            let mut space = cartesian_space("column", vec![5.0], 20.0);
            space.plot_groups[0].grouping = Some("stacked".into());
            space.axis_list.as_mut().unwrap()[1].id = Some(definition.into());
            space.plot_groups[0].axis_ids[1] = reference.into();
            assert_eq!(unsupported_feature(&space), None);
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder);
            let bars = series_rects(&commands, "#123456");
            assert_eq!(bars.len(), 1);
            assert!((bars[0].3 - 45.0).abs() < 0.01);
        }
    }

    #[test]
    fn newly_admitted_charts_with_empty_axis_references_render() {
        for family in ["scatter", "radar", "column"] {
            let mut space = cartesian_space(family, vec![5.0, 10.0, 15.0], 20.0);
            space.plot_groups[0].grouping = Some("stacked".into());
            space.plot_groups[0].axis_ids.clear();
            assert!(needed_main_gate(&space).is_some());
            assert_eq!(unsupported_feature(&space), None);
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder, "{family}");
            assert!(!commands.is_empty());
        }
    }

    #[test]
    fn scatter_second_value_axis_pair_renders_with_its_own_scale() {
        let mut space = scatter_space();
        let mut second = space.plot_groups[0].clone();
        second.axis_ids = vec!["3".into(), "4".into()];
        space.plot_groups.push(second);
        let axes = space.axis_list.as_mut().unwrap();
        let mut x_axis = axes[0].clone();
        x_axis.id = Some("3".into());
        x_axis.cross_axis_id = Some("4".into());
        let mut y_axis = axes[1].clone();
        y_axis.id = Some("4".into());
        y_axis.cross_axis_id = Some("3".into());
        y_axis.max = Some(100.0);
        axes.extend([x_axis, y_axis]);
        assert_eq!(unsupported_feature(&space), None);
        let (commands, regions) = render_scatter(&space);
        assert!(!regions[0].placeholder);
        assert_scatter_centers(
            &commands,
            &[
                (104.0, 195.0),
                (296.0, 150.0),
                (176.0, 105.0),
                (104.0, 231.0),
                (296.0, 222.0),
                (176.0, 213.0),
            ],
        );
    }

    #[test]
    fn scatter_pair_uses_references_without_requiring_crossing_metadata() {
        for case in 0..3 {
            let mut space = scatter_space();
            match case {
                0 => space.axis_list.as_mut().unwrap()[0].cross_axis_id = None,
                1 => space.axis_list.as_mut().unwrap()[1].cross_axis_id = Some("2".into()),
                _ => space.axis_list.as_mut().unwrap()[1].position = Some("top".into()),
            }
            assert_eq!(unsupported_feature(&space), None, "case {case}");
            let (commands, regions) = render_scatter(&space);
            assert!(!regions[0].placeholder);
            assert_scatter_centers(&commands, &[(104.0, 195.0), (296.0, 150.0), (176.0, 105.0)]);
        }
    }

    #[test]
    fn invalid_axis_references_still_degrade() {
        for case in 0..10 {
            let mut space = scatter_space();
            let expected = match case {
                0 => {
                    space.axis_list.as_mut().unwrap()[1].id = Some("1".into());
                    Some("ambiguous chart axis references")
                }
                1 => {
                    space.plot_groups[0].axis_ids = vec!["1".into(), "1".into()];
                    Some("ambiguous chart axis references")
                }
                2 => {
                    space.plot_groups[0].axis_ids = vec!["1".into(), "3".into()];
                    Some("unresolved chart axis references")
                }
                3 => {
                    space.plot_groups[0].axis_ids = vec!["1".into()];
                    None
                }
                4 => {
                    space.axis_list = None;
                    None
                }
                5 => {
                    let duplicate = space.axis_list.as_ref().unwrap()[1].clone();
                    space.axis_list.as_mut().unwrap().push(duplicate);
                    Some("ambiguous chart axis references")
                }
                6 => {
                    space.axis_list.as_mut().unwrap()[1].axis_type = "category".into();
                    None
                }
                7 => {
                    space.plot_groups[0].axis_ids.clear();
                    None
                }
                8 => {
                    let mut third = space.axis_list.as_ref().unwrap()[1].clone();
                    third.id = Some("3".into());
                    space.axis_list.as_mut().unwrap().push(third);
                    space.plot_groups[0].axis_ids.push("3".into());
                    None
                }
                _ => {
                    space.axis_list.as_mut().unwrap()[1].id = Some(String::new());
                    Some("unresolved chart axis references")
                }
            };
            assert_eq!(unsupported_feature(&space), expected, "case {case}");
            let sheet = charted_sheet(ChartAnchor::Absolute {
                pos: AnchorPos { x: 0, y: 0 },
                extent: AnchorExtent {
                    cx: 3_810_000,
                    cy: 2_857_500,
                },
            });
            let refusal = chart_refusal(&sheet.charts[0], &space);
            let (commands, regions) = render_scatter(&space);
            if let Some(expected) = expected {
                assert!(matches!(
                    refusal,
                    Some(RenderError::UnsupportedChartFeature { feature, .. }) if feature == expected
                ));
                assert!(regions[0].placeholder, "case {case}");
                assert_eq!(commands.len(), 5);
            } else {
                assert!(refusal.is_none(), "case {case}");
                assert!(!regions[0].placeholder, "case {case}");
                assert!(!commands.is_empty());
            }
        }
    }

    #[test]
    fn newly_admitted_charts_with_dangling_references_still_degrade() {
        for gate in ["stacked", "secondary", "logarithmic"] {
            let mut space = cartesian_space("column", vec![5.0], 20.0);
            match gate {
                "stacked" => space.plot_groups[0].grouping = Some("stacked".into()),
                "secondary" => {
                    let mut secondary = space.plot_groups[0].clone();
                    secondary.axis_ids[1] = "missing".into();
                    space.plot_groups.push(secondary);
                }
                _ => space.axis_list.as_mut().unwrap()[1].logarithmic_base = Some(10.0),
            }
            space.plot_groups[0].axis_ids[1] = "dangling".into();
            assert!(needed_main_gate(&space).is_some(), "{gate}");
            assert_eq!(
                unsupported_feature(&space),
                Some("unresolved chart axis references"),
                "{gate}"
            );
            let (commands, regions) = render_scatter(&space);
            assert!(regions[0].placeholder, "{gate}");
            assert_eq!(commands.len(), 5);
        }
    }

    #[test]
    fn referenced_numeric_aliases_remain_ambiguous_even_when_hidden() {
        for duplicate_reference in [false, true] {
            let mut space = cartesian_space("column", vec![5.0], 20.0);
            space.plot_groups[0].grouping = Some("stacked".into());
            if duplicate_reference {
                space.plot_groups[0].axis_ids.push("02".into());
            } else {
                let mut duplicate = space.axis_list.as_ref().unwrap()[1].clone();
                duplicate.id = Some("02".into());
                duplicate.hidden = true;
                space.axis_list.as_mut().unwrap().push(duplicate);
            }
            assert_eq!(
                unsupported_feature(&space),
                Some("ambiguous chart axis references")
            );
            let (commands, regions) = render_scatter(&space);
            assert!(regions[0].placeholder);
            assert_eq!(commands.len(), 5);
        }
    }

    #[test]
    fn mixed_scatter_and_column_groups_can_share_a_value_axis() {
        let mut mixed = scatter_space();
        let mut column = mixed.plot_groups[0].clone();
        column.chart_type = Some("column".into());
        column.axis_ids = vec!["3".into(), "2".into()];
        mixed.plot_groups.push(column);
        mixed.axis_list.as_mut().unwrap().push(ChartAxis {
            id: Some("3".into()),
            axis_type: "category".into(),
            ..Default::default()
        });
        mixed.plot_groups[1].series[0].color = "#654321".into();
        assert_eq!(unsupported_feature(&mixed), None);
        let (commands, regions) = render_scatter(&mixed);
        assert!(!regions[0].placeholder);
        assert_scatter_centers(&commands, &[(104.0, 195.0), (296.0, 150.0), (176.0, 105.0)]);
        assert_eq!(series_rects(&commands, "#654321").len(), 3);
    }

    #[test]
    fn bubble_crossing_value_axes_are_a_primary_pair() {
        let mut space = scatter_space();
        space.chart_type = "bubble".into();
        space.plot_groups[0].chart_type = Some("bubble".into());
        space.plot_groups[0].series[0].bubble_sizes = Some(vec![1.0, 2.0, 3.0]);
        assert_eq!(unsupported_feature(&space), None);
        let (commands, regions) = render_scatter(&space);
        assert!(!regions[0].placeholder);
        assert_eq!(
            commands
                .iter()
                .filter(|command| {
                    matches!(command, DrawCmd::Path { fill, .. } if fill.as_ref() == "#123456")
                })
                .count(),
            3
        );
    }

    #[test]
    fn a_moved_anchor_lands_where_the_pixels_asked_and_keeps_its_size() {
        let sheet = charted_sheet(ChartAnchor::TwoCell {
            from: cell(1, 0, 1, 0),
            to: cell(4, 0, 6, 0),
            edit_as: AnchorEditAs::TwoCell,
        });
        let geometry = GridGeometry::new(&sheet, &Stylesheet::default());
        let before = resolve_chart_anchor(sheet.charts[0].anchor, &geometry, 0, 0).unwrap();
        let moved = moved_chart_anchor(sheet.charts[0].anchor, &geometry, 17.0, -5.0).unwrap();
        let after = resolve_chart_anchor(moved, &geometry, 0, 0).unwrap();

        assert!((after.rect.x - (before.rect.x + 17.0)).abs() < 0.01);
        assert!((after.rect.y - (before.rect.y - 5.0)).abs() < 0.01);
        assert!((after.rect.w - before.rect.w).abs() < 0.01);
        assert!((after.rect.h - before.rect.h).abs() < 0.01);
        assert!(matches!(
            moved,
            ChartAnchor::TwoCell {
                edit_as: AnchorEditAs::TwoCell,
                ..
            }
        ));
    }

    #[test]
    fn a_one_cell_anchor_moves_without_resizing_and_clamps_at_the_origin() {
        let extent = AnchorExtent {
            cx: 952_500,
            cy: 476_250,
        };
        let sheet = charted_sheet(ChartAnchor::OneCell {
            from: cell(1, 0, 1, 0),
            extent,
        });
        let geometry = GridGeometry::new(&sheet, &Stylesheet::default());
        let moved =
            moved_chart_anchor(sheet.charts[0].anchor, &geometry, -10_000.0, -10_000.0).unwrap();
        assert_eq!(
            moved,
            ChartAnchor::OneCell {
                from: cell(0, 0, 0, 0),
                extent,
            }
        );
    }

    #[test]
    fn an_absolute_anchor_refuses_to_move() {
        let sheet = charted_sheet(ChartAnchor::Absolute {
            pos: AnchorPos { x: 0, y: 0 },
            extent: AnchorExtent {
                cx: 952_500,
                cy: 476_250,
            },
        });
        let geometry = GridGeometry::new(&sheet, &Stylesheet::default());
        assert!(moved_chart_anchor(sheet.charts[0].anchor, &geometry, 5.0, 5.0).is_none());
        assert!(
            moved_chart_anchor(
                ChartAnchor::OneCell {
                    from: cell(0, 0, 0, 0),
                    extent: AnchorExtent { cx: 100, cy: 100 },
                },
                &geometry,
                f64::NAN,
                0.0,
            )
            .is_none()
        );
    }

    #[test]
    fn regions_carry_the_geometry_the_display_list_publishes() {
        let sheet = charted_sheet(ChartAnchor::OneCell {
            from: cell(1, 0, 1, 0),
            extent: AnchorExtent {
                cx: 952_500,
                cy: 476_250,
            },
        });
        let viewport = Viewport {
            x: 0.0,
            y: 0.0,
            width: 400.0,
            height: 300.0,
        };
        let regions = chart_regions(&sheet, &Stylesheet::default(), &viewport).unwrap();
        assert_eq!(regions.len(), 1);
        assert_eq!(regions[0].id, "xl/drawings/drawing1.xml#0");
        assert!(regions[0].label.is_empty());
        assert_eq!((regions[0].rect.w, regions[0].rect.h), (100.0, 50.0));
        assert_eq!(regions[0].clip, regions[0].rect);

        let scrolled = chart_regions(
            &sheet,
            &Stylesheet::default(),
            &Viewport {
                x: 2_000.0,
                y: 2_000.0,
                width: 400.0,
                height: 300.0,
            },
        )
        .unwrap();
        assert!(scrolled.is_empty());
    }

    /// A chart part backs a frame, not the other way round: two anchors may
    /// name one part, and each must still address itself.
    #[test]
    fn two_anchors_on_one_chart_part_get_their_own_ids() {
        let extent = AnchorExtent {
            cx: 952_500,
            cy: 476_250,
        };
        let mut sheet = charted_sheet(ChartAnchor::OneCell {
            from: cell(1, 0, 1, 0),
            extent,
        });
        sheet.charts.push(SheetChart {
            part: sheet.charts[0].part.clone(),
            drawing: sheet.charts[0].drawing.clone(),
            anchor_index: 1,
            anchor: ChartAnchor::OneCell {
                from: cell(3, 0, 5, 0),
                extent,
            },
            refs: Vec::new(),
        });
        let regions = chart_regions(
            &sheet,
            &Stylesheet::default(),
            &Viewport {
                x: 0.0,
                y: 0.0,
                width: 400.0,
                height: 300.0,
            },
        )
        .unwrap();
        assert_eq!(regions.len(), 2);
        assert_ne!(regions[0].id, regions[1].id);
        assert_eq!(regions[1].id, "xl/drawings/drawing1.xml#1");
    }
}
