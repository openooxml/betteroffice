//! Tab-stop grid and tab-width math.
//!
//! The paragraph's declared stops are overlaid on an implicit 720-twip grid.
//! A tab advances from the line's current x — in content-area coordinates, the
//! same origin stop positions use — to the next stop past it, less whatever
//! the stop's alignment reserves for the runs that follow.
//!
//! The rules enforced here:
//!
//! - The grid interval is **always** 720 twips. `w:defaultTabStop` never
//!   reaches tab-width measurement; it feeds only the list-marker width.
//! - `decimal` stops measure like `start` stops: measurement reserves the full
//!   span, and a painter places the decimal point within it.
//! - `bar` stops draw a vertical rule and consume no horizontal space.
//! - Any `val` other than `clear`/`end`/`center`/`decimal`/`bar` behaves like
//!   `start`.
//! - A `clear` entry knocks out the grid line at that position as well as a
//!   declared stop; explicit stops within a hanging indent are retained; a positive
//!   left indent gains an implicit stop at the indent itself, so a tab on a
//!   hanging first line lands on the body text edge.
//! - A tab advances to the next stop *past* the pen; a stop the pen already
//!   rests on is spent, within the rounding whisker px coordinates carry.
//! - When the resolved span shrinks below a pixel — following content wider
//!   than an `end` stop's room — the tab gives up on the stop and takes plain
//!   default-grid spacing instead.
//! - A `start` stop reserves nothing for what follows, so the caller checks
//!   the following word itself; `end`, `center` and `bar` stops already
//!   account for it, which [`TabAdvance::reserves_following`] reports.

use super::input::TabStopIn;

/// Default tab interval: 720 twips = 0.5in = 48px.
pub(super) const DEFAULT_TAB_INTERVAL_TWIPS: f32 = 720.0;
/// Two positions closer than this count as the same stop.
const STOP_COINCIDENCE_TWIPS: f32 = 20.0;
/// A pen this close to a stop rests *on* it: an `end` stop parks the pen exactly
/// on itself, and the twips-px-twips round trip in f32 reads a whisker short.
/// The observed error is 1e-4 twips and authored stops are whole twips apart,
/// so this absorbs the noise without reaching a distinct stop.
const PEN_ON_STOP_TWIPS: f32 = 0.05;
/// The implicit grid is laid out to ten inches past the left indent.
const GRID_CEILING_SPAN_TWIPS: f32 = 14_400.0;

/// Converts 96-DPI pixels to twips.
pub(super) fn px_to_twips(px: f32) -> f32 {
    px / 96.0 * 1440.0
}

/// Converts twips to 96-DPI pixels.
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
        // Unknown values use start alignment.
        _ => StopKind::Start,
    }
}

fn same_stop_position(a: f32, b: f32) -> bool {
    (a - b).abs() < STOP_COINCIDENCE_TWIPS
}

/// The paragraph's effective stop list, in twips, sorted by position:
/// declared stops overlaid on the implicit 720-twip grid, with `clear`
/// entries knocked out.
fn compute_tab_stops(declared: &[TabStopIn], left_indent_twips: f32) -> Vec<(f32, StopKind)> {
    let mut kept: Vec<(f32, StopKind)> = Vec::new();
    let mut cleared_at: Vec<f32> = Vec::new();
    for stop in declared {
        if stop.val == "clear" {
            cleared_at.push(stop.pos);
        } else {
            kept.push((stop.pos, stop_kind(&stop.val)));
        }
    }

    let rightmost_kept = kept.iter().fold(0.0f32, |acc, s| acc.max(s.0));
    let mut grid = kept.clone();

    // hanging-indent paragraphs get an implicit stop at the indent itself so
    // a tab in the first line lands on the body text edge, matching Word
    if left_indent_twips > 0.0
        && !kept
            .iter()
            .any(|s| same_stop_position(s.0, left_indent_twips))
    {
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
    let mut grid_pos =
        ((grid_seed / DEFAULT_TAB_INTERVAL_TWIPS).floor() + 1.0) * DEFAULT_TAB_INTERVAL_TWIPS;
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

/// Distance from `from_x_px` to the next default-grid line, a full stride
/// when already sitting exactly on one.
fn default_grid_advance(from_x_px: f32) -> f32 {
    let stride_px = twips_to_px(DEFAULT_TAB_INTERVAL_TWIPS);
    let advance = stride_px - (from_x_px % stride_px);
    if advance <= 0.0 { stride_px } else { advance }
}

/// What one tab resolved to: its advance, and whether the stop it landed on
/// already reserved room for the runs that follow.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct TabAdvance {
    pub width: f32,
    pub reserves_following: bool,
}

impl TabAdvance {
    fn start(width: f32) -> Self {
        Self {
            width,
            reserves_following: false,
        }
    }
}

/// Advance a tab occupies starting from `current_x_px`, in content-area
/// coordinates. `following_width_px` is the inline width of the runs after
/// the tab, which `end` and `center` stops anchor on.
pub(super) fn calculate_tab_width(
    current_x_px: f32,
    declared: &[TabStopIn],
    left_indent_twips: f32,
    following_width_px: f32,
    available_width_px: f32,
) -> TabAdvance {
    let current_x_twips = px_to_twips(current_x_px);
    let grid = compute_tab_stops(declared, left_indent_twips);

    // past every stop in the grid: plain default-interval spacing
    let Some(&(pos, kind)) = grid
        .iter()
        .find(|s| s.0 > current_x_twips + PEN_ON_STOP_TWIPS)
    else {
        return TabAdvance::start(default_grid_advance(current_x_px));
    };

    let mut width = twips_to_px(pos) - current_x_px;
    match kind {
        StopKind::Center => width -= following_width_px / 2.0,
        StopKind::End => width -= following_width_px,
        // decimal measures like start (see module docs)
        StopKind::Decimal | StopKind::Start => {}
        // a bar stop draws a vertical rule but consumes no horizontal space
        StopKind::Bar => {
            return TabAdvance {
                width: 0.0,
                reserves_following: true,
            };
        }
    }

    // following content wider than the span: give up on the stop and use the
    // default grid instead
    if width < 1.0 {
        return TabAdvance::start(default_grid_advance(current_x_px));
    }
    if kind == StopKind::End && width > available_width_px + 1e-3 {
        let clamped = available_width_px - following_width_px;
        if clamped > 1.0 {
            return TabAdvance {
                width: clamped,
                reserves_following: true,
            };
        }
    }
    TabAdvance {
        width,
        reserves_following: matches!(kind, StopKind::End | StopKind::Center),
    }
}
