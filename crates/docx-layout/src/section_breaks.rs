//! Section geometry across section breaks.
//!
//! ECMA-376 §17.6.22: a section break's `w:type` says how the *next* section
//! starts relative to this one, and an absent `w:type` means `nextPage`.
//! [`handle_section_break`] is the entry point the placement walk uses; it
//! drives the paginator directly:
//!
//! - `nextPage` applies the next section's page size and margins immediately,
//!   then forces a page.
//! - `evenPage` / `oddPage` advance under the preceding section until reaching
//!   the requested parity, then apply the next section's page geometry.
//! - `continuous` keeps the current sheet and defers the new geometry to the
//!   next natural page break — unless the page size changes, which Word and
//!   LibreOffice promote to a page break because two page sizes cannot share
//!   one physical sheet. Sizes are compared after rounding.
//! - `nextColumn` advances in the current band and defers geometry until the
//!   next page. A spent band falls through to a new page; page-size or
//!   column-count changes are promoted because they cannot share the sheet.
//!
//! New column geometry is applied at a fresh region and otherwise queued. An
//! omitted column layout resolves to one full-width column.
//!
//! The tracker half ([`SectionLayoutTracker`] and the `apply` / `promote` /
//! `resolve_next_*` functions) is not wired into placement. It is a two-stage
//! scheduling model, and its `nextColumn` arm neither promotes incompatible
//! geometry nor opens the live column region. Its unit tests cover bookkeeping,
//! not the behavior of [`handle_section_break`]; the two must be reconciled
//! before the tracker gains a consumer.
//!
//! Rounding uses ties-toward-`+∞` and NaN-propagating maximum so the values
//! reaching the canonical JSON match the host's numeric semantics.
#![allow(dead_code)]

use crate::LayoutError;
use crate::prescan::{SectionLayoutConfig, default_columns};
use crate::types::{ColumnLayout, PageMargins, SectionBreakBlock, SectionBreakType, Size};

const SINGLE_COLUMN: ColumnLayout = ColumnLayout {
    count: 1.0,
    gap: 0.0,
    equal_width: None,
    separator: None,
    columns: None,
};

/// Margin fields scheduled by a section break. `None` means "not scheduled",
/// so the in-force value carries forward at the boundary.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct PartialPageMargins {
    pub top: Option<f64>,
    pub right: Option<f64>,
    pub bottom: Option<f64>,
    pub left: Option<f64>,
    pub header: Option<f64>,
    pub footer: Option<f64>,
}

/// Fully resolved page or region geometry.
#[derive(Clone, Debug, PartialEq)]
pub struct SectionGeometry {
    pub margins: PageMargins,
    pub page_size: Size,
    pub columns: ColumnLayout,
    pub orientation: Option<String>,
}

/// Geometry queued for the next page or region.
#[derive(Clone, Debug, PartialEq)]
pub struct QueuedGeometry {
    pub margins: Option<PartialPageMargins>,
    pub page_size: Option<Size>,
    pub columns: Option<ColumnLayout>,
    pub orientation: Option<String>,
}

/// Geometry the current page uses, plus what the next boundary will adopt.
#[derive(Clone, Debug, PartialEq)]
pub struct SectionLayoutTracker {
    pub in_force: SectionGeometry,
    pub queued: QueuedGeometry,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PageParity {
    Even,
    Odd,
}

/// Required pagination action at a section break.
#[derive(Clone, Debug, PartialEq)]
pub struct SectionBreakOutcome {
    /// Move to a new page before continuing.
    pub break_to_new_page: bool,
    /// When breaking, the page number parity the new page must satisfy.
    pub page_parity: Option<PageParity>,
    /// Begin a new column region on the current page (continuous column change).
    pub open_column_region: bool,
    /// Move to the next column of the region in force (`nextColumn`).
    pub advance_column: bool,
}

#[derive(Clone, Debug, PartialEq)]
pub struct ApplySectionBreakResult {
    pub outcome: SectionBreakOutcome,
    pub tracker: SectionLayoutTracker,
}

/// Page-flow operations required by section break handling.
pub trait SectionBreakPaginator {
    /// Updates page geometry, rejecting an empty content area.
    fn update_page_layout(
        &mut self,
        page_size: Option<&Size>,
        margins: Option<&PageMargins>,
        apply_immediately: bool,
    ) -> Result<(), LayoutError>;
    /// Forces a page break and returns the new page number.
    fn force_page_break(&mut self) -> u32;
    /// Moves to the next column, reporting whether placement now sits at the
    /// start of a fresh column region. Idempotent on a pristine page.
    fn advance_to_next_column(&mut self) -> bool;
    /// Create a new page even when the current page is pristine.
    fn insert_blank_page(&mut self) -> u32;
    /// Returns the current page size, creating the first page if needed.
    fn current_page_size(&mut self) -> Size;
    fn current_columns(&self) -> ColumnLayout;
    /// Updates the active columns.
    fn update_columns(&mut self, columns: &ColumnLayout);
    /// Queues columns for the next page, leaving the band in force alone.
    fn queue_columns(&mut self, columns: &ColumnLayout);
}

// JS Math.round: nearest integer, ties toward +infinity (Rust's f64::round
// ties away from zero, which differs for negative halves).
fn js_math_round(x: f64) -> f64 {
    if !x.is_finite() || x == 0.0 {
        return x;
    }
    let f = x.floor();
    if x - f >= 0.5 { f + 1.0 } else { f }
}

// Two page sizes cannot share one physical sheet. Sizes are compared after
// rounding, matching the host's numeric semantics.
fn page_size_differs(current: &Size, next: &Size) -> bool {
    js_math_round(next.w) != js_math_round(current.w)
        || js_math_round(next.h) != js_math_round(current.h)
}

// JS Math.max(a, b): NaN-propagating (Rust's f64::max ignores NaN).
fn js_max(a: f64, b: f64) -> f64 {
    if a.is_nan() || b.is_nan() {
        f64::NAN
    } else {
        a.max(b)
    }
}

fn empty_queue() -> QueuedGeometry {
    QueuedGeometry {
        margins: None,
        page_size: None,
        columns: None,
        orientation: None,
    }
}

// Body margins are always scheduled; header and footer remain optional.
fn schedule_margins(
    current: Option<&PartialPageMargins>,
    incoming: &PageMargins,
) -> PartialPageMargins {
    let mut merged = current.cloned().unwrap_or_default();
    merged.top = Some(js_max(0.0, incoming.top));
    merged.right = Some(js_max(0.0, incoming.right));
    merged.bottom = Some(js_max(0.0, incoming.bottom));
    merged.left = Some(js_max(0.0, incoming.left));
    if let Some(v) = incoming.header {
        merged.header = Some(js_max(0.0, v));
    }
    if let Some(v) = incoming.footer {
        merged.footer = Some(js_max(0.0, v));
    }
    merged
}

// Scheduled fields override in-force margins.
fn overlay_margins(base: &PageMargins, over: &PartialPageMargins) -> PageMargins {
    PageMargins {
        top: over.top.unwrap_or(base.top),
        right: over.right.unwrap_or(base.right),
        bottom: over.bottom.unwrap_or(base.bottom),
        left: over.left.unwrap_or(base.left),
        header: over.header.or(base.header),
        footer: over.footer.or(base.footer),
    }
}

/// Seeds section tracking from document defaults.
pub fn create_section_layout_tracker(
    margins: &PageMargins,
    page_size: &Size,
    columns: Option<&ColumnLayout>,
) -> SectionLayoutTracker {
    SectionLayoutTracker {
        in_force: SectionGeometry {
            margins: margins.clone(),
            page_size: page_size.clone(),
            columns: columns.cloned().unwrap_or(SINGLE_COLUMN),
            orientation: None,
        },
        queued: empty_queue(),
    }
}

/// Schedules a break's geometry and reports what the paginator must do now.
///
/// The tracker is never mutated in place, so a caller can discard the result.
/// `nextPage` / `evenPage` / `oddPage` always break and queue their columns.
/// `nextColumn` queues and advances here; live-geometry promotion remains in
/// [`handle_section_break`]. A
/// `continuous` break opens a new column region on the current page only when
/// its column count or gap differs from the columns in force; a section
/// declaring no columns resolves to a single full-width column.
pub fn apply_section_break(
    block: &SectionBreakBlock,
    tracker: &SectionLayoutTracker,
) -> ApplySectionBreakResult {
    let mut updated = tracker.clone();
    let break_kind = block.break_type.unwrap_or(SectionBreakType::Continuous);

    // Omitted geometry remains queued as None and inherits at the boundary.
    if block.orientation.as_deref().is_some_and(|s| !s.is_empty()) {
        updated.queued.orientation = block.orientation.clone();
    }
    if let Some(page_size) = &block.page_size {
        updated.queued.page_size = Some(Size {
            w: page_size.w,
            h: page_size.h,
        });
    }
    if let Some(margins) = &block.margins {
        updated.queued.margins = Some(schedule_margins(updated.queued.margins.as_ref(), margins));
    }

    let incoming_columns = block.columns.clone().unwrap_or(SINGLE_COLUMN);

    let starts_on_new_page = matches!(
        break_kind,
        SectionBreakType::NextPage | SectionBreakType::EvenPage | SectionBreakType::OddPage
    );
    if starts_on_new_page {
        updated.queued.columns = Some(incoming_columns.clone());
        let mut outcome = SectionBreakOutcome {
            break_to_new_page: true,
            page_parity: None,
            open_column_region: false,
            advance_column: false,
        };
        if break_kind == SectionBreakType::EvenPage {
            outcome.page_parity = Some(PageParity::Even);
        }
        if break_kind == SectionBreakType::OddPage {
            outcome.page_parity = Some(PageParity::Odd);
        }
        return ApplySectionBreakResult {
            outcome,
            tracker: updated,
        };
    }

    if break_kind == SectionBreakType::NextColumn {
        updated.queued.columns = Some(incoming_columns);
        return ApplySectionBreakResult {
            outcome: SectionBreakOutcome {
                break_to_new_page: false,
                page_parity: None,
                open_column_region: false,
                advance_column: true,
            },
            tracker: updated,
        };
    }

    // continuous: only a column change forces a new region on the current page
    let columns_differ = incoming_columns.count != updated.in_force.columns.count
        || incoming_columns.gap != updated.in_force.columns.gap;
    if columns_differ {
        updated.queued.columns = Some(incoming_columns);
        return ApplySectionBreakResult {
            outcome: SectionBreakOutcome {
                break_to_new_page: false,
                page_parity: None,
                open_column_region: true,
                advance_column: false,
            },
            tracker: updated,
        };
    }

    ApplySectionBreakResult {
        outcome: SectionBreakOutcome {
            break_to_new_page: false,
            page_parity: None,
            open_column_region: false,
            advance_column: false,
        },
        tracker: updated,
    }
}

/// Promotes queued geometry at a page or region boundary.
pub fn promote_queued_geometry(tracker: &SectionLayoutTracker) -> SectionLayoutTracker {
    let mut in_force = tracker.in_force.clone();
    let queued = &tracker.queued;

    if let Some(margins) = &queued.margins {
        in_force.margins = overlay_margins(&in_force.margins, margins);
    }
    if let Some(page_size) = &queued.page_size {
        in_force.page_size = page_size.clone();
    }
    if let Some(columns) = &queued.columns {
        in_force.columns = columns.clone();
    }
    if queued.orientation.is_some() {
        in_force.orientation = queued.orientation.clone();
    }

    SectionLayoutTracker {
        in_force,
        queued: empty_queue(),
    }
}

/// Returns the margins for the next page or region.
pub fn resolve_next_margins(tracker: &SectionLayoutTracker) -> PageMargins {
    overlay_margins(
        &tracker.in_force.margins,
        tracker
            .queued
            .margins
            .as_ref()
            .unwrap_or(&PartialPageMargins::default()),
    )
}

/// Returns the page size for the next page or region.
pub fn resolve_next_page_size(tracker: &SectionLayoutTracker) -> Size {
    tracker
        .queued
        .page_size
        .clone()
        .unwrap_or_else(|| tracker.in_force.page_size.clone())
}

/// Returns the columns for the next page or region.
pub fn resolve_next_columns(tracker: &SectionLayoutTracker) -> ColumnLayout {
    tracker
        .queued
        .columns
        .clone()
        .unwrap_or_else(|| tracker.in_force.columns.clone())
}

// One inch at 96 DPI.
const DEFAULT_MARGIN_PX: f64 = 96.0;

/// Fills missing body margins with one inch and defaults the header/footer
/// distances to the resolved top/bottom body margins.
pub fn resolve_page_margins(requested: Option<&PageMargins>) -> PageMargins {
    let top = requested.map_or(DEFAULT_MARGIN_PX, |m| m.top);
    let right = requested.map_or(DEFAULT_MARGIN_PX, |m| m.right);
    let bottom = requested.map_or(DEFAULT_MARGIN_PX, |m| m.bottom);
    let left = requested.map_or(DEFAULT_MARGIN_PX, |m| m.left);
    PageMargins {
        top,
        right,
        bottom,
        left,
        header: Some(requested.and_then(|m| m.header).unwrap_or(top)),
        footer: Some(requested.and_then(|m| m.footer).unwrap_or(bottom)),
    }
}

/// Drives the paginator through a section break, per the module rules.
///
/// Returns whether the new section opened a fresh column region. A same-band
/// `nextColumn` queues its columns and returns false.
///
/// `next_section_config` and `next_section_type` describe the section being
/// entered, not the one ending; the block itself carries no geometry the
/// paginator needs at this point.
pub fn handle_section_break<P: SectionBreakPaginator>(
    _block: &SectionBreakBlock,
    paginator: &mut P,
    next_section_config: &SectionLayoutConfig,
    next_section_type: Option<SectionBreakType>,
) -> Result<bool, LayoutError> {
    // ECMA-376 §17.6.22: w:type specifies how the NEXT section starts relative
    // to this one. Default is 'nextPage' when w:type is absent.
    let break_type = next_section_type.unwrap_or(SectionBreakType::NextPage);
    let page_size = Some(&next_section_config.page_size);
    let margins = Some(&next_section_config.margins);
    let default_cols = default_columns();
    let next_columns = next_section_config
        .columns
        .as_ref()
        .unwrap_or(&default_cols);
    let mut opened_column_region = true;

    match break_type {
        SectionBreakType::NextPage => {
            paginator.update_page_layout(page_size, margins, true)?;
            paginator.force_page_break();
        }

        SectionBreakType::EvenPage => {
            let page_number = paginator.force_page_break();
            // If landed on odd page, add another page
            if page_number % 2 != 0 {
                paginator.insert_blank_page();
            }
            paginator.update_page_layout(page_size, margins, true)?;
        }

        SectionBreakType::OddPage => {
            let page_number = paginator.force_page_break();
            // If landed on even page, add another page
            if page_number % 2 == 0 {
                paginator.insert_blank_page();
            }
            paginator.update_page_layout(page_size, margins, true)?;
        }

        SectionBreakType::Continuous => {
            // ECMA-376 §17.6.22: a `continuous` break normally keeps the current
            // page geometry and defers the new size/margins to the next natural
            // page break. BUT a continuous break that changes page size or
            // orientation cannot share a physical sheet with the preceding
            // section, so Word and LibreOffice promote it to a page break when
            // the next section's page size differs from the current page's.
            let current_size = paginator.current_page_size();
            if page_size_differs(&current_size, &next_section_config.page_size) {
                paginator.update_page_layout(page_size, margins, true)?;
                paginator.force_page_break();
            } else {
                paginator
                    .update_page_layout(page_size, margins, /* apply_immediately */ false)?;
            }
        }

        SectionBreakType::NextColumn => {
            // Page size and column count cannot change mid-sheet, so Word
            // promotes incompatible breaks. Otherwise geometry is deferred as
            // the current band advances, falling through when it is spent.
            let current_size = paginator.current_page_size();
            let incompatible = page_size_differs(&current_size, &next_section_config.page_size)
                || next_columns.count != paginator.current_columns().count;
            if incompatible {
                paginator.update_page_layout(page_size, margins, true)?;
                paginator.force_page_break();
            } else {
                paginator
                    .update_page_layout(page_size, margins, /* apply_immediately */ false)?;
                opened_column_region = paginator.advance_to_next_column();
            }
        }
    }

    if opened_column_region {
        paginator.update_columns(next_columns);
    } else {
        paginator.queue_columns(next_columns);
    }
    Ok(opened_column_region)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::BlockId;
    use serde_json::{Value, json};

    fn margins(top: f64, right: f64, bottom: f64, left: f64) -> PageMargins {
        PageMargins {
            top,
            right,
            bottom,
            left,
            header: None,
            footer: None,
        }
    }

    fn empty_break() -> SectionBreakBlock {
        SectionBreakBlock {
            sdt_groups: None,
            id: BlockId::Num(0.0),
            break_type: None,
            page_size: None,
            orientation: None,
            margins: None,
            columns: None,
        }
    }

    fn base_tracker() -> SectionLayoutTracker {
        create_section_layout_tracker(
            &margins(96.0, 96.0, 96.0, 96.0),
            &Size {
                w: 816.0,
                h: 1056.0,
            },
            None,
        )
    }

    // ---- pure tracker functions -------------------------------------------

    #[test]
    fn create_tracker_defaults_to_single_column_and_empty_queue() {
        let tracker = base_tracker();
        assert_eq!(tracker.in_force.columns, SINGLE_COLUMN);
        assert_eq!(tracker.in_force.orientation, None);
        assert_eq!(tracker.queued, empty_queue());
    }

    #[test]
    fn next_page_break_queues_geometry_and_breaks() {
        let tracker = base_tracker();
        let block = SectionBreakBlock {
            break_type: Some(SectionBreakType::NextPage),
            page_size: Some(Size {
                w: 1056.0,
                h: 816.0,
            }),
            orientation: Some("landscape".to_string()),
            margins: Some(margins(50.0, 50.0, 50.0, 50.0)),
            ..empty_break()
        };
        let result = apply_section_break(&block, &tracker);
        assert!(result.outcome.break_to_new_page);
        assert_eq!(result.outcome.page_parity, None);
        assert!(!result.outcome.open_column_region);
        assert_eq!(
            result.tracker.queued.page_size,
            Some(Size {
                w: 1056.0,
                h: 816.0
            })
        );
        assert_eq!(
            result.tracker.queued.orientation,
            Some("landscape".to_string())
        );
        assert_eq!(
            result.tracker.queued.margins,
            Some(PartialPageMargins {
                top: Some(50.0),
                right: Some(50.0),
                bottom: Some(50.0),
                left: Some(50.0),
                header: None,
                footer: None,
            })
        );
        assert_eq!(result.tracker.queued.columns, Some(SINGLE_COLUMN));
        // inForce untouched until promotion
        assert_eq!(result.tracker.in_force, tracker.in_force);
    }

    #[test]
    fn even_and_odd_breaks_report_page_parity() {
        let tracker = base_tracker();
        let even = apply_section_break(
            &SectionBreakBlock {
                break_type: Some(SectionBreakType::EvenPage),
                ..empty_break()
            },
            &tracker,
        );
        assert_eq!(even.outcome.page_parity, Some(PageParity::Even));
        let odd = apply_section_break(
            &SectionBreakBlock {
                break_type: Some(SectionBreakType::OddPage),
                ..empty_break()
            },
            &tracker,
        );
        assert_eq!(odd.outcome.page_parity, Some(PageParity::Odd));
    }

    #[test]
    fn continuous_break_with_same_columns_is_a_no_op_outcome() {
        let tracker = base_tracker();
        let result = apply_section_break(&empty_break(), &tracker);
        assert!(!result.outcome.break_to_new_page);
        assert!(!result.outcome.open_column_region);
        assert_eq!(result.tracker.queued.columns, None);
    }

    #[test]
    fn continuous_break_with_column_change_opens_region() {
        let tracker = base_tracker();
        let columns = ColumnLayout {
            count: 2.0,
            gap: 20.0,
            equal_width: Some(true),
            separator: Some(true),
            columns: None,
        };
        let result = apply_section_break(
            &SectionBreakBlock {
                columns: Some(columns.clone()),
                ..empty_break()
            },
            &tracker,
        );
        assert!(!result.outcome.break_to_new_page);
        assert!(result.outcome.open_column_region);
        assert_eq!(result.tracker.queued.columns, Some(columns));
    }

    #[test]
    fn next_page_break_queues_authored_columns() {
        let tracker = base_tracker();
        let columns = ColumnLayout {
            count: 2.0,
            gap: 20.0,
            equal_width: Some(false),
            separator: Some(true),
            columns: None,
        };
        let result = apply_section_break(
            &SectionBreakBlock {
                break_type: Some(SectionBreakType::NextPage),
                columns: Some(columns.clone()),
                ..empty_break()
            },
            &tracker,
        );
        assert_eq!(result.tracker.queued.columns, Some(columns));
    }

    #[test]
    fn scheduled_margins_clamp_negatives_and_fold_onto_prior_schedule() {
        let tracker = base_tracker();
        let first = apply_section_break(
            &SectionBreakBlock {
                margins: Some(margins(-10.0, 20.0, 20.0, 30.0)),
                ..empty_break()
            },
            &tracker,
        );
        // negative distances clamp to 0 (Math.max(0, value))
        assert_eq!(
            first.tracker.queued.margins.as_ref().unwrap().top,
            Some(0.0)
        );
        assert_eq!(
            first.tracker.queued.margins.as_ref().unwrap().left,
            Some(30.0)
        );

        let second = apply_section_break(
            &SectionBreakBlock {
                margins: Some(PageMargins {
                    header: Some(12.0),
                    ..margins(5.0, 20.0, 20.0, 40.0)
                }),
                ..empty_break()
            },
            &first.tracker,
        );
        // later fields fold onto the earlier schedule; untouched keys survive
        let queued = second.tracker.queued.margins.unwrap();
        assert_eq!(queued.top, Some(5.0));
        assert_eq!(queued.left, Some(40.0));
        assert_eq!(queued.header, Some(12.0));
        assert_eq!(queued.footer, None);
    }

    #[test]
    fn promote_queued_geometry_folds_and_clears() {
        let tracker = base_tracker();
        let block = SectionBreakBlock {
            break_type: Some(SectionBreakType::NextPage),
            page_size: Some(Size {
                w: 1200.0,
                h: 700.0,
            }),
            orientation: Some("landscape".to_string()),
            margins: Some(margins(50.0, 96.0, 96.0, 96.0)),
            ..empty_break()
        };
        let queued = apply_section_break(&block, &tracker).tracker;
        let promoted = promote_queued_geometry(&queued);
        assert_eq!(
            promoted.in_force.page_size,
            Size {
                w: 1200.0,
                h: 700.0
            }
        );
        assert_eq!(promoted.in_force.orientation, Some("landscape".to_string()));
        // scheduled margins overlay the inForce ones; absent keys carry forward
        assert_eq!(promoted.in_force.margins.top, 50.0);
        assert_eq!(promoted.in_force.margins.bottom, 96.0);
        assert_eq!(promoted.in_force.margins.header, None);
        assert_eq!(promoted.in_force.columns, SINGLE_COLUMN);
        assert_eq!(promoted.queued, empty_queue());
    }

    #[test]
    fn resolve_next_values_overlay_queued_over_in_force() {
        let tracker = base_tracker();
        assert_eq!(
            resolve_next_page_size(&tracker),
            Size {
                w: 816.0,
                h: 1056.0
            }
        );
        assert_eq!(resolve_next_columns(&tracker), SINGLE_COLUMN);
        assert_eq!(resolve_next_margins(&tracker), tracker.in_force.margins);

        let block = SectionBreakBlock {
            break_type: Some(SectionBreakType::NextPage),
            page_size: Some(Size { w: 500.0, h: 500.0 }),
            margins: Some(margins(10.0, 96.0, 96.0, 96.0)),
            ..empty_break()
        };
        let queued = apply_section_break(&block, &tracker).tracker;
        assert_eq!(resolve_next_page_size(&queued), Size { w: 500.0, h: 500.0 });
        let next_margins = resolve_next_margins(&queued);
        assert_eq!(next_margins.top, 10.0);
        assert_eq!(next_margins.left, 96.0);
    }

    #[test]
    fn resolve_page_margins_defaults_and_header_footer() {
        let resolved = resolve_page_margins(None);
        assert_eq!(resolved.top, 96.0);
        assert_eq!(resolved.header, Some(96.0));
        assert_eq!(resolved.footer, Some(96.0));

        // header/footer default to the RESOLVED top/bottom body margins
        let resolved = resolve_page_margins(Some(&margins(10.0, 20.0, 30.0, 40.0)));
        assert_eq!(resolved.header, Some(10.0));
        assert_eq!(resolved.footer, Some(30.0));

        // an explicit 0 is honored, not replaced by a default
        let zero = resolve_page_margins(Some(&PageMargins {
            header: Some(0.0),
            ..margins(0.0, 96.0, 96.0, 96.0)
        }));
        assert_eq!(zero.top, 0.0);
        assert_eq!(zero.header, Some(0.0));
        assert_eq!(zero.footer, Some(96.0));
    }

    #[test]
    fn js_math_round_matches_js_semantics() {
        assert_eq!(js_math_round(0.5), 1.0);
        assert_eq!(js_math_round(-0.5), 0.0); // JS: -0, ties toward +infinity
        assert_eq!(js_math_round(-1.5), -1.0);
        assert_eq!(js_math_round(2.4), 2.0);
        assert_eq!(js_math_round(816.0), 816.0);
        // spec edge: closest double below 0.5 must not round up
        assert_eq!(js_math_round(0.49999999999999994), 0.0);
    }

    #[derive(Debug, PartialEq)]
    enum ColumnLanding {
        AlreadyFresh,
        SameRegion,
        NewPage,
    }

    #[derive(Debug, PartialEq)]
    enum Call {
        UpdatePageLayout {
            page_size: Option<Size>,
            margins_top: Option<f64>,
            apply_immediately: bool,
        },
        ForcePageBreak {
            new_page_number: u32,
        },
        InsertBlankPage {
            new_page_number: u32,
        },
        AdvanceColumn(ColumnLanding),
        UpdateColumns {
            count: f64,
            gap: f64,
        },
        QueueColumns {
            count: f64,
            gap: f64,
        },
    }

    struct MockPaginator {
        page_size: Size,
        margins_top: f64,
        page_number: u32,
        pending_page_size: Option<Size>,
        pending_margins_top: Option<f64>,
        columns: ColumnLayout,
        pending_columns: Option<ColumnLayout>,
        column_index: usize,
        pristine: bool,
        /// Geometry stamped on the current sheet, distinct from pending state.
        page: Size,
        page_margins_top: f64,
        calls: Vec<Call>,
    }

    const INITIAL_MARGIN_TOP: f64 = 96.0;

    impl MockPaginator {
        fn new(page_size: Size, page_number: u32) -> Self {
            MockPaginator {
                page: page_size.clone(),
                page_margins_top: INITIAL_MARGIN_TOP,
                page_size,
                margins_top: INITIAL_MARGIN_TOP,
                page_number,
                pending_page_size: None,
                pending_margins_top: None,
                columns: SINGLE_COLUMN,
                pending_columns: None,
                column_index: 0,
                pristine: false,
                calls: Vec::new(),
            }
        }

        fn with_columns(page_size: Size, page_number: u32, columns: ColumnLayout) -> Self {
            MockPaginator {
                columns,
                ..MockPaginator::new(page_size, page_number)
            }
        }

        fn stamp_page(&mut self) {
            self.page = self.page_size.clone();
            self.page_margins_top = self.margins_top;
        }

        fn open_page(&mut self) {
            if let Some(size) = self.pending_page_size.take() {
                self.page_size = size;
            }
            if let Some(top) = self.pending_margins_top.take() {
                self.margins_top = top;
            }
            if let Some(columns) = self.pending_columns.take() {
                self.columns = columns;
            }
            self.page_number += 1;
            self.column_index = 0;
            self.pristine = true;
            self.stamp_page();
        }
    }

    impl SectionBreakPaginator for MockPaginator {
        fn update_page_layout(
            &mut self,
            page_size: Option<&Size>,
            margins: Option<&PageMargins>,
            apply_immediately: bool,
        ) -> Result<(), LayoutError> {
            self.calls.push(Call::UpdatePageLayout {
                page_size: page_size.cloned(),
                margins_top: margins.map(|m| m.top),
                apply_immediately,
            });
            if apply_immediately {
                if let Some(size) = page_size {
                    self.page_size = size.clone();
                }
                if let Some(margins) = margins {
                    self.margins_top = margins.top;
                }
                self.pending_page_size = None;
                self.pending_margins_top = None;
                if self.pristine {
                    self.stamp_page();
                }
            } else {
                if let Some(size) = page_size {
                    self.pending_page_size = Some(size.clone());
                }
                if let Some(margins) = margins {
                    self.pending_margins_top = Some(margins.top);
                }
            }
            Ok(())
        }

        fn force_page_break(&mut self) -> u32 {
            if !self.pristine {
                self.open_page();
            }
            self.calls.push(Call::ForcePageBreak {
                new_page_number: self.page_number,
            });
            self.page_number
        }

        fn insert_blank_page(&mut self) -> u32 {
            self.open_page();
            self.calls.push(Call::InsertBlankPage {
                new_page_number: self.page_number,
            });
            self.page_number
        }

        fn advance_to_next_column(&mut self) -> bool {
            let landing = if self.pristine && self.column_index == 0 {
                ColumnLanding::AlreadyFresh
            } else if (self.column_index as f64) + 1.0 < self.columns.count {
                self.column_index += 1;
                self.pristine = false;
                ColumnLanding::SameRegion
            } else {
                self.open_page();
                ColumnLanding::NewPage
            };
            let opened_region = landing != ColumnLanding::SameRegion;
            self.calls.push(Call::AdvanceColumn(landing));
            opened_region
        }

        fn current_page_size(&mut self) -> Size {
            self.page.clone()
        }

        fn current_columns(&self) -> ColumnLayout {
            self.columns.clone()
        }

        fn update_columns(&mut self, columns: &ColumnLayout) {
            self.calls.push(Call::UpdateColumns {
                count: columns.count,
                gap: columns.gap,
            });
            self.pending_columns = None;
            self.columns = columns.clone();
        }

        fn queue_columns(&mut self, columns: &ColumnLayout) {
            self.calls.push(Call::QueueColumns {
                count: columns.count,
                gap: columns.gap,
            });
            self.pending_columns = Some(columns.clone());
        }
    }

    const PORTRAIT: Size = Size {
        w: 800.0,
        h: 1000.0,
    };
    const LANDSCAPE: Size = Size {
        w: 1200.0,
        h: 700.0,
    };

    fn config(page_size: Size, columns: Option<ColumnLayout>) -> SectionLayoutConfig {
        SectionLayoutConfig {
            page_size,
            margins: margins(50.0, 50.0, 50.0, 50.0),
            columns,
        }
    }

    #[test]
    fn continuous_break_same_size_defers_geometry_without_breaking() {
        let mut paginator = MockPaginator::new(PORTRAIT, 1);
        handle_section_break(
            &empty_break(),
            &mut paginator,
            &config(PORTRAIT, None),
            Some(SectionBreakType::Continuous),
        )
        .unwrap();
        assert_eq!(
            paginator.calls,
            vec![
                Call::UpdatePageLayout {
                    page_size: Some(PORTRAIT),
                    margins_top: Some(50.0),
                    apply_immediately: false,
                },
                Call::UpdateColumns {
                    count: 1.0,
                    gap: 0.0
                },
            ]
        );
        // current page still uses the old geometry
        assert_eq!(paginator.current_page_size(), PORTRAIT);
    }

    #[test]
    fn continuous_break_with_size_change_is_promoted_to_page_break() {
        let mut paginator = MockPaginator::new(PORTRAIT, 1);

        // sb1: next section is landscape
        handle_section_break(
            &empty_break(),
            &mut paginator,
            &config(LANDSCAPE, None),
            Some(SectionBreakType::Continuous),
        )
        .unwrap();
        assert_eq!(paginator.page_number, 2);
        assert_eq!(paginator.current_page_size(), LANDSCAPE);
        assert!(paginator.calls.contains(&Call::UpdatePageLayout {
            page_size: Some(LANDSCAPE),
            margins_top: Some(50.0),
            apply_immediately: true,
        }));

        // Give the second break content to break away from.
        paginator.pristine = false;
        handle_section_break(
            &empty_break(),
            &mut paginator,
            &config(PORTRAIT, None),
            Some(SectionBreakType::Continuous),
        )
        .unwrap();
        assert_eq!(paginator.page_number, 3);
        assert_eq!(paginator.current_page_size(), PORTRAIT);
    }

    // Sub-pixel size difference rounds equal (Math.round on both sides), so a
    // continuous break must NOT be promoted.
    #[test]
    fn continuous_break_rounds_sizes_before_comparing() {
        let mut paginator = MockPaginator::new(PORTRAIT, 1);
        handle_section_break(
            &empty_break(),
            &mut paginator,
            &config(Size { w: 800.4, h: 999.6 }, None),
            Some(SectionBreakType::Continuous),
        )
        .unwrap();
        assert_eq!(paginator.page_number, 1);
        assert!(matches!(
            paginator.calls[0],
            Call::UpdatePageLayout {
                apply_immediately: false,
                ..
            }
        ));
    }

    #[test]
    fn next_page_break_updates_layout_immediately_and_breaks_once() {
        let mut paginator = MockPaginator::new(PORTRAIT, 1);
        handle_section_break(
            &empty_break(),
            &mut paginator,
            &config(LANDSCAPE, None),
            Some(SectionBreakType::NextPage),
        )
        .unwrap();
        assert_eq!(
            paginator.calls,
            vec![
                Call::UpdatePageLayout {
                    page_size: Some(LANDSCAPE),
                    margins_top: Some(50.0),
                    apply_immediately: true,
                },
                Call::ForcePageBreak { new_page_number: 2 },
                Call::UpdateColumns {
                    count: 1.0,
                    gap: 0.0
                },
            ]
        );
    }

    // Absent w:type defaults to 'nextPage' (ECMA-376 §17.6.22).
    #[test]
    fn missing_break_type_defaults_to_next_page() {
        let mut paginator = MockPaginator::new(PORTRAIT, 1);
        handle_section_break(
            &empty_break(),
            &mut paginator,
            &config(PORTRAIT, None),
            None,
        )
        .unwrap();
        assert!(
            paginator
                .calls
                .contains(&Call::ForcePageBreak { new_page_number: 2 })
        );
    }

    #[test]
    fn even_page_break_adds_extra_page_when_landing_odd() {
        // page 2 → break lands on 3 (odd) → evenPage forces one more (4)
        let mut paginator = MockPaginator::new(PORTRAIT, 2);
        handle_section_break(
            &empty_break(),
            &mut paginator,
            &config(PORTRAIT, None),
            Some(SectionBreakType::EvenPage),
        )
        .unwrap();
        assert_eq!(paginator.page_number, 4);

        // page 1 → break lands on 2 (even) → no extra page
        let mut paginator = MockPaginator::new(PORTRAIT, 1);
        handle_section_break(
            &empty_break(),
            &mut paginator,
            &config(PORTRAIT, None),
            Some(SectionBreakType::EvenPage),
        )
        .unwrap();
        assert_eq!(paginator.page_number, 2);
    }

    #[test]
    fn odd_page_break_adds_extra_page_when_landing_even() {
        // page 1 → break lands on 2 (even) → oddPage forces one more (3)
        let mut paginator = MockPaginator::new(PORTRAIT, 1);
        handle_section_break(
            &empty_break(),
            &mut paginator,
            &config(PORTRAIT, None),
            Some(SectionBreakType::OddPage),
        )
        .unwrap();
        assert_eq!(paginator.page_number, 3);

        // page 2 → break lands on 3 (odd) → no extra page
        let mut paginator = MockPaginator::new(PORTRAIT, 2);
        handle_section_break(
            &empty_break(),
            &mut paginator,
            &config(PORTRAIT, None),
            Some(SectionBreakType::OddPage),
        )
        .unwrap();
        assert_eq!(paginator.page_number, 3);
    }

    #[test]
    fn continuous_break_applies_next_section_columns() {
        let mut paginator = MockPaginator::new(Size { w: 500.0, h: 500.0 }, 1);
        handle_section_break(
            &empty_break(),
            &mut paginator,
            &config(
                Size { w: 500.0, h: 500.0 },
                Some(ColumnLayout {
                    count: 2.0,
                    gap: 20.0,
                    equal_width: None,
                    separator: None,
                    columns: None,
                }),
            ),
            Some(SectionBreakType::Continuous),
        )
        .unwrap();
        assert_eq!(paginator.page_number, 1);
        assert_eq!(
            paginator.calls.last(),
            Some(&Call::UpdateColumns {
                count: 2.0,
                gap: 20.0
            })
        );
    }

    fn cols(count: f64, gap: f64) -> ColumnLayout {
        ColumnLayout {
            count,
            gap,
            equal_width: None,
            separator: None,
            columns: None,
        }
    }

    fn next_column_break<P: SectionBreakPaginator>(
        paginator: &mut P,
        next: SectionLayoutConfig,
    ) -> bool {
        handle_section_break(
            &empty_break(),
            paginator,
            &next,
            Some(SectionBreakType::NextColumn),
        )
        .unwrap()
    }

    #[test]
    fn next_column_break_stays_in_the_band_and_queues_the_new_columns() {
        let mut paginator = MockPaginator::with_columns(PORTRAIT, 1, cols(2.0, 24.0));
        let opened_region =
            next_column_break(&mut paginator, config(PORTRAIT, Some(cols(2.0, 36.0))));
        assert!(!opened_region);
        assert_eq!(paginator.page_number, 1);
        assert_eq!(paginator.column_index, 1);
        assert_eq!(paginator.columns.gap, 24.0);
        assert_eq!(
            paginator.calls,
            vec![
                Call::UpdatePageLayout {
                    page_size: Some(PORTRAIT),
                    margins_top: Some(50.0),
                    apply_immediately: false,
                },
                Call::AdvanceColumn(ColumnLanding::SameRegion),
                Call::QueueColumns {
                    count: 2.0,
                    gap: 36.0
                },
            ]
        );

        paginator.open_page();
        assert_eq!(paginator.columns.gap, 36.0);
    }

    #[test]
    fn next_column_break_promotes_a_column_count_change_to_a_page() {
        for next in [cols(3.0, 24.0), cols(1.0, 0.0)] {
            let mut paginator = MockPaginator::with_columns(PORTRAIT, 1, cols(2.0, 24.0));
            let opened_region =
                next_column_break(&mut paginator, config(PORTRAIT, Some(next.clone())));
            assert!(opened_region);
            assert_eq!(paginator.page_number, 2);
            assert_eq!(
                paginator.calls,
                vec![
                    Call::UpdatePageLayout {
                        page_size: Some(PORTRAIT),
                        margins_top: Some(50.0),
                        apply_immediately: true,
                    },
                    Call::ForcePageBreak { new_page_number: 2 },
                    Call::UpdateColumns {
                        count: next.count,
                        gap: next.gap
                    },
                ]
            );
        }
    }

    #[test]
    fn next_column_break_promotes_a_page_size_change_to_a_page() {
        let mut paginator = MockPaginator::with_columns(PORTRAIT, 1, cols(2.0, 24.0));
        let opened_region =
            next_column_break(&mut paginator, config(LANDSCAPE, Some(cols(2.0, 24.0))));
        assert!(opened_region);
        assert_eq!(paginator.page_number, 2);
        assert_eq!(paginator.current_page_size(), LANDSCAPE);
        assert!(
            paginator
                .calls
                .contains(&Call::ForcePageBreak { new_page_number: 2 })
        );
    }

    #[test]
    fn next_column_break_from_the_last_column_opens_a_page() {
        let mut paginator = MockPaginator::with_columns(PORTRAIT, 1, cols(2.0, 24.0));
        paginator.column_index = 1;
        let opened_region =
            next_column_break(&mut paginator, config(PORTRAIT, Some(cols(2.0, 24.0))));
        assert!(opened_region);
        assert_eq!(paginator.page_number, 2);
        assert_eq!(paginator.column_index, 0);
        assert_eq!(
            paginator.calls,
            vec![
                Call::UpdatePageLayout {
                    page_size: Some(PORTRAIT),
                    margins_top: Some(50.0),
                    apply_immediately: false,
                },
                Call::AdvanceColumn(ColumnLanding::NewPage),
                Call::UpdateColumns {
                    count: 2.0,
                    gap: 24.0
                },
            ]
        );
    }

    #[test]
    fn next_column_break_in_a_single_column_section_starts_a_page() {
        let mut paginator = MockPaginator::new(PORTRAIT, 1);
        let opened_region = next_column_break(&mut paginator, config(PORTRAIT, None));
        assert!(opened_region);
        assert_eq!(paginator.page_number, 2);
        assert!(
            paginator
                .calls
                .contains(&Call::AdvanceColumn(ColumnLanding::NewPage))
        );
    }

    #[test]
    fn next_column_break_on_a_pristine_page_adds_no_blank_page() {
        for columns in [cols(1.0, 0.0), cols(2.0, 24.0)] {
            let mut paginator = MockPaginator::with_columns(PORTRAIT, 1, columns.clone());
            paginator.pristine = true;
            let opened_region = next_column_break(&mut paginator, config(PORTRAIT, Some(columns)));
            assert!(opened_region);
            assert_eq!(paginator.page_number, 1);
            assert_eq!(paginator.column_index, 0);
            assert_eq!(
                paginator.calls[1],
                Call::AdvanceColumn(ColumnLanding::AlreadyFresh)
            );
        }
    }

    #[test]
    fn promoted_breaks_on_a_pristine_page_re_form_it() {
        let mut paginator = MockPaginator::new(PORTRAIT, 1);
        paginator.pristine = true;
        for (break_type, size) in [
            (SectionBreakType::Continuous, LANDSCAPE),
            (SectionBreakType::NextColumn, PORTRAIT),
        ] {
            handle_section_break(
                &empty_break(),
                &mut paginator,
                &config(size.clone(), None),
                Some(break_type),
            )
            .unwrap();
            assert_eq!(paginator.page_number, 1);
            assert_eq!(paginator.current_page_size(), size);
            assert_eq!(paginator.page_margins_top, 50.0);
        }
    }

    #[test]
    fn apply_next_column_break_advances_the_column_and_queues_columns() {
        let tracker = base_tracker();
        let columns = cols(2.0, 24.0);
        let result = apply_section_break(
            &SectionBreakBlock {
                break_type: Some(SectionBreakType::NextColumn),
                columns: Some(columns.clone()),
                ..empty_break()
            },
            &tracker,
        );
        assert!(result.outcome.advance_column);
        assert!(!result.outcome.break_to_new_page);
        assert!(!result.outcome.open_column_region);
        assert_eq!(result.tracker.queued.columns, Some(columns));
    }

    fn measured_paragraph_with_line_height(id: u32, lines: usize, line_height: f64) -> Value {
        json!({
            "block": {
                "kind": "paragraph",
                "id": id,
                "runs": [{"kind": "text", "text": "x"}],
                "attrs": {}
            },
            "measure": {
                "kind": "paragraph",
                "lines": (0..lines)
                    .map(|_| json!({
                        "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1,
                        "width": 200,
                        "ascent": line_height * 0.8,
                        "descent": line_height * 0.2,
                        "lineHeight": line_height
                    }))
                    .collect::<Vec<_>>(),
                "totalHeight": lines as f64 * line_height
            }
        })
    }

    fn measured_paragraph(id: u32, lines: usize) -> Value {
        measured_paragraph_with_line_height(id, lines, 24.0)
    }

    #[derive(Default)]
    struct Scenario {
        first: Option<ColumnLayout>,
        second: Option<ColumnLayout>,
        final_page_size: Option<Size>,
        head_lines: usize,
        tail_lines: usize,
    }

    fn layout_of(break_type: &str, scenario: Scenario) -> crate::types::Layout {
        let mut section_break = json!({"kind": "sectionBreak", "id": 1});
        if let Some(columns) = &scenario.first {
            section_break["columns"] = serde_json::to_value(columns).unwrap();
        }
        let mut options = json!({
            "pageSize": {"w": 816, "h": 1056},
            "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
            "bodyBreakType": break_type
        });
        if let Some(columns) = &scenario.second {
            options["columns"] = serde_json::to_value(columns).unwrap();
        }
        if let Some(size) = &scenario.final_page_size {
            options["finalPageSize"] = serde_json::to_value(size).unwrap();
        }
        let input = json!({
            "measured": [
                measured_paragraph(0, scenario.head_lines.max(1)),
                {"block": section_break, "measure": {"kind": "sectionBreak"}},
                measured_paragraph(2, scenario.tail_lines.max(1))
            ],
            "options": options
        });
        crate::compute_layout(&input.to_string()).unwrap()
    }

    fn placements(layout: &crate::types::Layout) -> Vec<(u32, f64, f64)> {
        layout
            .pages
            .iter()
            .flat_map(|page| {
                page.fragments.iter().filter_map(move |fragment| {
                    let crate::types::Fragment::Paragraph(paragraph) = fragment else {
                        return None;
                    };
                    Some((page.number, paragraph.x, paragraph.y))
                })
            })
            .collect()
    }

    fn parity_mismatch_layout(break_type: &str, leading: bool) -> crate::types::Layout {
        let mut measured = Vec::new();
        if !leading {
            measured.push(measured_paragraph(0, 1));
        }
        measured.push(json!({
            "block": {
                "kind": "sectionBreak",
                "id": 1,
                "columns": {"count": 2, "gap": 24}
            },
            "measure": {"kind": "sectionBreak"}
        }));
        measured.push(measured_paragraph(2, 1));
        let input = json!({
            "measured": measured,
            "options": {
                "pageSize": {"w": 816, "h": 1056},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
                "bodyBreakType": break_type,
                "columns": {"count": 3, "gap": 48},
                "finalPageSize": {"w": 1056, "h": 816},
                "finalMargins": {"top": 48, "right": 72, "bottom": 64, "left": 80}
            }
        });
        crate::compute_layout(&input.to_string()).unwrap()
    }

    fn page_spine(layout: &crate::types::Layout) -> Vec<(u32, f64, f64, f64, f64, usize, usize)> {
        layout
            .pages
            .iter()
            .map(|page| {
                let columns = page.columns.as_ref().unwrap_or(&SINGLE_COLUMN);
                (
                    page.number,
                    page.size.w,
                    page.size.h,
                    columns.count,
                    columns.gap,
                    page.fragments.len(),
                    page.region_section_index,
                )
            })
            .collect()
    }

    fn distinct_section_furniture() -> crate::regions::DocumentRegions {
        serde_json::from_value(json!({
            "sections": [
                {
                    "sectionId": "old",
                    "headerFooterRefs": {
                        "headerDefault": "header-old",
                        "footerDefault": "footer-old"
                    },
                    "pageBorders": {"source": "old"},
                    "watermark": {"source": "old"},
                    "pageNumbering": {"start": 9}
                },
                {
                    "sectionId": "new",
                    "headerFooterRefs": {
                        "headerDefault": "header-new",
                        "footerDefault": "footer-new"
                    },
                    "pageBorders": {"source": "new"},
                    "watermark": {"source": "new"},
                    "pageNumbering": {"start": 3}
                }
            ]
        }))
        .unwrap()
    }

    fn assert_section_furniture(
        page: &crate::types::Page,
        section_index: u64,
        section_id: &str,
        section_page_number: u64,
    ) {
        let refs = page.header_footer_refs.as_ref().unwrap();
        assert_eq!(page.section_index, Some(section_index));
        assert_eq!(page.section_id.as_deref(), Some(section_id));
        assert_eq!(
            refs.header_default.as_deref(),
            Some(format!("header-{section_id}").as_str())
        );
        assert_eq!(
            refs.footer_default.as_deref(),
            Some(format!("footer-{section_id}").as_str())
        );
        assert_eq!(
            page.page_borders.as_ref(),
            Some(&json!({"source": section_id}))
        );
        assert_eq!(
            page.watermark.as_ref(),
            Some(&json!({"source": section_id}))
        );
        assert_eq!(page.section_page_number, Some(section_page_number));
    }

    #[test]
    fn next_column_section_break_places_the_next_section_in_the_next_column() {
        let layout = layout_of(
            "nextColumn",
            Scenario {
                first: Some(cols(2.0, 24.0)),
                second: Some(cols(2.0, 24.0)),
                ..Scenario::default()
            },
        );
        assert_eq!(layout.pages.len(), 1);
        assert_eq!(placements(&layout), vec![(1, 96.0, 96.0), (1, 420.0, 96.0)]);
    }

    #[test]
    fn next_column_section_does_not_inherit_the_outgoing_balance_limit() {
        let mut measured = vec![
            measured_paragraph_with_line_height(0, 2, 100.0),
            json!({
                "block": {
                    "kind": "sectionBreak",
                    "id": 1,
                    "columns": {"count": 2, "gap": 20}
                },
                "measure": {"kind": "sectionBreak"}
            }),
        ];
        measured.extend((2..8).map(|id| measured_paragraph_with_line_height(id, 1, 100.0)));
        let input = json!({
            "measured": measured,
            "options": {
                "pageSize": {"w": 800, "h": 1000},
                "margins": {"top": 100, "right": 100, "bottom": 100, "left": 100},
                "bodyBreakType": "nextColumn",
                "columns": {"count": 2, "gap": 20}
            }
        });
        let layout = crate::compute_layout(&input.to_string()).unwrap();
        assert_eq!(layout.pages.len(), 1);
        assert_eq!(
            placements(&layout),
            vec![
                (1, 100.0, 100.0),
                (1, 410.0, 100.0),
                (1, 410.0, 200.0),
                (1, 410.0, 300.0),
                (1, 410.0, 400.0),
                (1, 410.0, 500.0),
                (1, 410.0, 600.0),
            ]
        );
    }

    #[test]
    fn later_next_column_section_does_not_inherit_the_outgoing_balance_limit() {
        let mut measured = vec![
            measured_paragraph_with_line_height(0, 1, 100.0),
            json!({
                "block": {
                    "kind": "sectionBreak",
                    "id": 1,
                    "columns": {"count": 2, "gap": 20}
                },
                "measure": {"kind": "sectionBreak"}
            }),
            measured_paragraph_with_line_height(2, 2, 100.0),
            json!({
                "block": {
                    "kind": "sectionBreak",
                    "id": 3,
                    "type": "nextPage",
                    "columns": {"count": 2, "gap": 20}
                },
                "measure": {"kind": "sectionBreak"}
            }),
        ];
        measured.extend((4..10).map(|id| measured_paragraph_with_line_height(id, 1, 100.0)));
        let input = json!({
            "measured": measured,
            "options": {
                "pageSize": {"w": 800, "h": 1000},
                "margins": {"top": 100, "right": 100, "bottom": 100, "left": 100},
                "bodyBreakType": "nextColumn",
                "columns": {"count": 2, "gap": 20}
            }
        });
        let layout = crate::compute_layout(&input.to_string()).unwrap();
        assert_eq!(layout.pages.len(), 2);
        assert_eq!(
            placements(&layout),
            vec![
                (1, 100.0, 100.0),
                (2, 100.0, 100.0),
                (2, 410.0, 100.0),
                (2, 410.0, 200.0),
                (2, 410.0, 300.0),
                (2, 410.0, 400.0),
                (2, 410.0, 500.0),
                (2, 410.0, 600.0),
            ]
        );
    }

    #[test]
    fn next_column_section_break_defers_a_gap_change_to_the_next_page() {
        let layout = layout_of(
            "nextColumn",
            Scenario {
                first: Some(cols(2.0, 24.0)),
                second: Some(cols(2.0, 120.0)),
                tail_lines: 74,
                ..Scenario::default()
            },
        );
        assert_eq!(layout.pages.len(), 2);
        assert_eq!(
            placements(&layout),
            vec![
                (1, 96.0, 96.0),
                (1, 420.0, 96.0),
                (2, 96.0, 96.0),
                (2, 468.0, 96.0),
            ]
        );
        assert_eq!(layout.pages[0].columns.as_ref().unwrap().gap, 24.0);
        assert_eq!(layout.pages[1].columns.as_ref().unwrap().gap, 120.0);
    }

    #[test]
    fn next_column_section_break_promoted_by_a_column_count_change() {
        for second in [cols(3.0, 24.0), cols(1.0, 0.0)] {
            let layout = layout_of(
                "nextColumn",
                Scenario {
                    first: Some(cols(2.0, 24.0)),
                    second: Some(second.clone()),
                    ..Scenario::default()
                },
            );
            assert_eq!(layout.pages.len(), 2);
            assert_eq!(placements(&layout), vec![(1, 96.0, 96.0), (2, 96.0, 96.0)]);
            assert_eq!(
                layout.pages[1].columns.as_ref().map_or(1.0, |c| c.count),
                second.count
            );
        }
    }

    #[test]
    fn next_column_section_break_promoted_by_a_page_size_change() {
        let landscape = Size {
            w: 1056.0,
            h: 816.0,
        };
        let layout = layout_of(
            "nextColumn",
            Scenario {
                first: Some(cols(2.0, 24.0)),
                second: Some(cols(2.0, 24.0)),
                final_page_size: Some(landscape.clone()),
                ..Scenario::default()
            },
        );
        assert_eq!(layout.pages.len(), 2);
        assert_eq!(placements(&layout), vec![(1, 96.0, 96.0), (2, 96.0, 96.0)]);
        assert_eq!(layout.pages[1].size, landscape);
    }

    #[test]
    fn next_column_section_break_in_a_single_column_document_starts_a_new_page() {
        let layout = layout_of("nextColumn", Scenario::default());
        assert_eq!(layout.pages.len(), 2);
        assert_eq!(placements(&layout), vec![(1, 96.0, 96.0), (2, 96.0, 96.0)]);
    }

    #[test]
    fn next_page_section_break_leaves_the_band_for_a_new_page() {
        let layout = layout_of(
            "nextPage",
            Scenario {
                first: Some(cols(2.0, 24.0)),
                second: Some(cols(2.0, 24.0)),
                ..Scenario::default()
            },
        );
        assert_eq!(layout.pages.len(), 2);
        assert_eq!(placements(&layout), vec![(1, 96.0, 96.0), (2, 96.0, 96.0)]);
    }

    #[test]
    fn leading_even_page_mismatch_preserves_the_filler_section() {
        let mut layout = parity_mismatch_layout("evenPage", true);
        assert_eq!(
            page_spine(&layout),
            vec![
                (1, 816.0, 1056.0, 2.0, 24.0, 0, 0),
                (2, 1056.0, 816.0, 3.0, 48.0, 1, 1),
            ]
        );

        crate::regions::apply_document_regions(&mut layout, &distinct_section_furniture());
        assert_section_furniture(&layout.pages[0], 0, "old", 9);
        assert_section_furniture(&layout.pages[1], 1, "new", 3);
    }

    #[test]
    fn odd_page_mismatch_preserves_the_filler_section() {
        let mut layout = parity_mismatch_layout("oddPage", false);
        assert_eq!(
            page_spine(&layout),
            vec![
                (1, 816.0, 1056.0, 2.0, 24.0, 1, 0),
                (2, 816.0, 1056.0, 2.0, 24.0, 0, 0),
                (3, 1056.0, 816.0, 3.0, 48.0, 1, 1),
            ]
        );

        crate::regions::apply_document_regions(&mut layout, &distinct_section_furniture());
        assert_section_furniture(&layout.pages[0], 0, "old", 9);
        assert_section_furniture(&layout.pages[1], 0, "old", 10);
        assert_section_furniture(&layout.pages[2], 1, "new", 3);
    }

    #[test]
    fn leading_next_column_section_break_adds_no_blank_page() {
        let input = json!({
            "measured": [
                {"block": {"kind": "sectionBreak", "id": 0, "columns": {"count": 2, "gap": 24}},
                 "measure": {"kind": "sectionBreak"}},
                measured_paragraph(1, 1)
            ],
            "options": {
                "pageSize": {"w": 816, "h": 1056},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
                "bodyBreakType": "nextColumn",
                "columns": {"count": 2, "gap": 24}
            }
        });
        let layout = crate::compute_layout(&input.to_string()).unwrap();
        assert_eq!(layout.pages.len(), 1);
        assert_eq!(placements(&layout), vec![(1, 96.0, 96.0)]);
    }

    #[test]
    fn promoted_break_on_a_pristine_page_reports_the_geometry_it_laid_out_to() {
        let landscape = Size {
            w: 1056.0,
            h: 816.0,
        };
        let input = json!({
            "measured": [
                {"block": {"kind": "sectionBreak", "id": 0, "columns": {"count": 2, "gap": 24}},
                 "measure": {"kind": "sectionBreak"}},
                measured_paragraph(1, 1)
            ],
            "options": {
                "pageSize": {"w": 816, "h": 1056},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
                "bodyBreakType": "nextColumn",
                "columns": {"count": 2, "gap": 24},
                "finalPageSize": {"w": 1056, "h": 816},
                "finalMargins": {"top": 48, "right": 96, "bottom": 96, "left": 96}
            }
        });
        let layout = crate::compute_layout(&input.to_string()).unwrap();
        assert_eq!(layout.pages.len(), 1);
        assert_eq!(layout.pages[0].size, landscape);
        assert_eq!(layout.pages[0].margins.top, 48.0);
        let crate::types::Fragment::Paragraph(paragraph) = &layout.pages[0].fragments[0] else {
            panic!("expected a paragraph fragment");
        };
        assert_eq!(paragraph.width, landscape.w - 192.0);
        assert_eq!(paragraph.y, 48.0);
    }

    #[test]
    fn restamped_pristine_page_uses_the_new_content_limit() {
        let input = json!({
            "measured": [
                {"block": {"kind": "sectionBreak", "id": 0},
                 "measure": {"kind": "sectionBreak"}},
                measured_paragraph(1, 30)
            ],
            "options": {
                "pageSize": {"w": 816, "h": 1056},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
                "bodyBreakType": "nextPage",
                "finalPageSize": {"w": 816, "h": 816},
                "finalMargins": {"top": 48, "right": 96, "bottom": 96, "left": 96}
            }
        });
        let layout = crate::compute_layout(&input.to_string()).unwrap();
        let fragments: Vec<_> = layout
            .pages
            .iter()
            .flat_map(|page| {
                page.fragments.iter().filter_map(move |fragment| {
                    let crate::types::Fragment::Paragraph(paragraph) = fragment else {
                        return None;
                    };
                    Some((
                        page.number,
                        paragraph.from_line,
                        paragraph.to_line,
                        paragraph.y,
                    ))
                })
            })
            .collect();

        assert_eq!(layout.pages.len(), 2);
        assert_eq!(fragments, vec![(1, 0, 28, 48.0), (2, 28, 30, 48.0)]);
    }

    #[test]
    fn restamped_pristine_page_uses_the_new_section_furniture() {
        let input = json!({
            "measured": [
                {"block": {"kind": "sectionBreak", "id": 0},
                 "measure": {"kind": "sectionBreak"}},
                measured_paragraph(1, 1)
            ],
            "options": {
                "pageSize": {"w": 816, "h": 1056},
                "margins": {"top": 96, "right": 96, "bottom": 96, "left": 96},
                "bodyBreakType": "nextPage",
                "finalPageSize": {"w": 1056, "h": 816},
                "finalMargins": {"top": 48, "right": 96, "bottom": 96, "left": 96}
            }
        });
        let mut layout = crate::compute_layout(&input.to_string()).unwrap();
        assert_eq!(layout.pages[0].region_section_index, 1);

        let regions: crate::regions::DocumentRegions = serde_json::from_value(json!({
            "sections": [
                {
                    "sectionId": "old",
                    "headerFooterRefs": {
                        "headerDefault": "header-old",
                        "footerDefault": "footer-old"
                    },
                    "pageBorders": {"source": "old"},
                    "watermark": {"source": "old"},
                    "pageNumbering": {"start": 9}
                },
                {
                    "sectionId": "new",
                    "headerFooterRefs": {
                        "headerDefault": "header-new",
                        "footerDefault": "footer-new"
                    },
                    "pageBorders": {"source": "new"},
                    "watermark": {"source": "new"},
                    "pageNumbering": {"start": 3}
                }
            ]
        }))
        .unwrap();
        crate::regions::apply_document_regions(&mut layout, &regions);

        let page = &layout.pages[0];
        let refs = page.header_footer_refs.as_ref().unwrap();
        assert_eq!(page.section_index, Some(1));
        assert_eq!(page.section_id.as_deref(), Some("new"));
        assert_eq!(refs.header_default.as_deref(), Some("header-new"));
        assert_eq!(refs.footer_default.as_deref(), Some("footer-new"));
        assert_eq!(page.page_borders, Some(json!({"source": "new"})));
        assert_eq!(page.watermark, Some(json!({"source": "new"})));
        assert_eq!(page.section_page_number, Some(3));
    }
}
