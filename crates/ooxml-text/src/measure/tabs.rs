//! Tab-stop grid and tab-width math — a port of the measurement subset of
//! the TS shared tab model (`packages/core/src/prosemirror/utils/
//! tabCalculator.ts`, `computeTabStops` + `calculateTabWidth`), so the Rust
//! measurer agrees with the TS measurer (and therefore the painter) on tab
//! advances.
//!
//! Semantics mirrored exactly, including the TS measurement path's
//! simplifications:
//!
//! - The default grid interval is **always 720 twips** here: `measureParagraph`
//!   builds its `TabSettings` without `defaultTabInterval`, so the document's
//!   `w:defaultTabStop` (`attrs.defaultTabStopTwips`) does NOT reach tab-width
//!   measurement in TS either (it only feeds the list-marker width path).
//! - `decimal` stops measure like `start` stops: the TS measurer passes only
//!   `followingWidth` (never `decimalPrefixWidth`), so `calculateTabWidth`
//!   subtracts 0 for decimal alignment. The painter positions the decimal
//!   point properly at paint time; measurement reserves the full span.
//! - `bar` stops consume no horizontal space (width 0).
//! - A stop value that is none of `clear`/`end`/`center`/`decimal`/`bar`
//!   behaves like `start`, matching how the TS string comparisons fall
//!   through.

use super::input::TabStopIn;

/// TS `DEFAULT_TAB_INTERVAL_TWIPS`: 720 twips = 0.5in = 48px.
pub(super) const DEFAULT_TAB_INTERVAL_TWIPS: f32 = 720.0;
/// TS `STOP_COINCIDENCE_TWIPS`: two positions within this count as one stop.
const STOP_COINCIDENCE_TWIPS: f32 = 20.0;
/// TS lays the implicit grid out to ten inches past the indent.
const GRID_CEILING_SPAN_TWIPS: f32 = 14_400.0;

/// px (96dpi) → twips, the TS `pixelsToTwips`.
pub(super) fn px_to_twips(px: f32) -> f32 {
    px / 96.0 * 1440.0
}

/// twips → px (96dpi), the TS `twipsToPixels`.
pub(super) fn twips_to_px(twips: f32) -> f32 {
    twips / 1440.0 * 96.0
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StopKind {
    Start,
    End,
    Center,
    Decimal,
    Bar,
}

fn stop_kind(val: &str) -> StopKind {
    match val {
        "end" => StopKind::End,
        "center" => StopKind::Center,
        "decimal" => StopKind::Decimal,
        "bar" => StopKind::Bar,
        // any unrecognized value falls through every TS comparison → start
        _ => StopKind::Start,
    }
}

fn same_stop_position(a: f32, b: f32) -> bool {
    (a - b).abs() < STOP_COINCIDENCE_TWIPS
}

/// TS `computeTabStops`: the paragraph's declared stops overlaid on the
/// implicit 720-twip grid, `clear` entries knocked out, stops left of the
/// indent dropped, plus the implicit stop at a positive left indent. Sorted
/// by position. Positions are twips.
fn compute_tab_stops(declared: &[TabStopIn], left_indent_twips: f32) -> Vec<(f32, StopKind)> {
    let mut kept: Vec<(f32, StopKind)> = Vec::new();
    let mut cleared_at: Vec<f32> = Vec::new();
    for stop in declared {
        if stop.val == "clear" {
            cleared_at.push(stop.pos);
        } else if stop.pos >= left_indent_twips {
            kept.push((stop.pos, stop_kind(&stop.val)));
        }
    }

    let rightmost_kept = kept.iter().fold(0.0f32, |acc, s| acc.max(s.0));
    let mut grid = kept.clone();

    // hanging-indent paragraphs get an implicit stop at the indent itself so
    // a tab in the first line lands on the body text edge, matching Word
    if left_indent_twips > 0.0 && !kept.iter().any(|s| s.0 <= left_indent_twips) {
        let indent_cleared = cleared_at
            .iter()
            .any(|&p| same_stop_position(p, left_indent_twips));
        if !indent_cleared {
            grid.push((left_indent_twips, StopKind::Start));
        }
    }

    // implicit default grid past the last declared stop, out to ten inches
    // beyond the indent (bounded: ceiling − seed ≤ 14400 → ≤ 21 iterations)
    let grid_seed = if rightmost_kept > 0.0 {
        rightmost_kept.max(left_indent_twips)
    } else {
        left_indent_twips
    };
    let grid_ceiling = left_indent_twips + GRID_CEILING_SPAN_TWIPS;
    let mut grid_pos = grid_seed + DEFAULT_TAB_INTERVAL_TWIPS;
    while grid_pos - DEFAULT_TAB_INTERVAL_TWIPS < grid_ceiling {
        let shadowed = kept.iter().any(|s| same_stop_position(s.0, grid_pos));
        let knocked_out = cleared_at.iter().any(|&p| same_stop_position(p, grid_pos));
        let duplicates_indent =
            left_indent_twips > 0.0 && same_stop_position(grid_pos, left_indent_twips);
        if !shadowed && !knocked_out && !duplicates_indent {
            grid.push((grid_pos, StopKind::Start));
        }
        grid_pos += DEFAULT_TAB_INTERVAL_TWIPS;
    }

    grid.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap_or(std::cmp::Ordering::Equal));
    grid
}

/// TS `defaultGridAdvance`: from a pixel x to the next default-grid line; a
/// full stride when already sitting exactly on one. `%` keeps the dividend's
/// sign in both JS and Rust, so negative x behaves identically.
fn default_grid_advance(from_x_px: f32) -> f32 {
    let stride_px = twips_to_px(DEFAULT_TAB_INTERVAL_TWIPS);
    let advance = stride_px - (from_x_px % stride_px);
    if advance <= 0.0 { stride_px } else { advance }
}

/// TS `calculateTabWidth(...).width`: the advance a tab occupies from
/// `current_x_px` (content-area coordinates — the origin tab positions are
/// measured from). `following_width_px` is the inline width of the runs after
/// the tab (`measureInlineWidthAfterTab`), anchored on `end`/`center` stops.
pub(super) fn calculate_tab_width(
    current_x_px: f32,
    declared: &[TabStopIn],
    left_indent_twips: f32,
    following_width_px: f32,
) -> f32 {
    let current_x_twips = px_to_twips(current_x_px);
    let grid = compute_tab_stops(declared, left_indent_twips);

    // past every stop in the grid: plain default-interval spacing
    let Some(&(pos, kind)) = grid.iter().find(|s| s.0 > current_x_twips) else {
        return default_grid_advance(current_x_px);
    };

    let mut width = twips_to_px(pos) - current_x_px;
    match kind {
        StopKind::Center => width -= following_width_px / 2.0,
        StopKind::End => width -= following_width_px,
        // decimal measures like start (see module docs)
        StopKind::Decimal | StopKind::Start => {}
        // a bar stop draws a vertical rule but consumes no horizontal space
        StopKind::Bar => return 0.0,
    }

    // following content wider than the span: give up on the stop and use the
    // default grid instead
    if width < 1.0 {
        return default_grid_advance(current_x_px);
    }
    width
}
