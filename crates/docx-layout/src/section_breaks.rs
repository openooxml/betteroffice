use crate::LayoutError;
use crate::prescan::{SectionLayoutConfig, default_columns};
use crate::types::{ColumnLayout, PageMargins, SectionBreakBlock, SectionBreakType, Size};

const SINGLE_COLUMN: ColumnLayout = ColumnLayout {
    count: 1.0,
    gap: 0.0,
    equal_width: None,
    separator: None,
};

#[derive(Clone, Debug, Default, PartialEq)]
pub struct PartialPageMargins {
    pub top: Option<f64>,
    pub right: Option<f64>,
    pub bottom: Option<f64>,
    pub left: Option<f64>,
    pub header: Option<f64>,
    pub footer: Option<f64>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct SectionGeometry {
    pub margins: PageMargins,
    pub page_size: Size,
    pub columns: ColumnLayout,
    pub orientation: Option<String>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct QueuedGeometry {
    pub margins: Option<PartialPageMargins>,
    pub page_size: Option<Size>,
    pub columns: Option<ColumnLayout>,
    pub orientation: Option<String>,
}

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

#[derive(Clone, Debug, PartialEq)]
pub struct SectionBreakOutcome {
    /// Move to a new page before continuing.
    pub break_to_new_page: bool,
    pub page_parity: Option<PageParity>,
    /// Begin a new column region on the current page (continuous column change).
    pub open_column_region: bool,
}

#[derive(Clone, Debug, PartialEq)]
pub struct ApplySectionBreakResult {
    pub outcome: SectionBreakOutcome,
    pub tracker: SectionLayoutTracker,
}

pub trait SectionBreakPaginator {
    fn update_page_layout(
        &mut self,
        page_size: Option<&Size>,
        margins: Option<&PageMargins>,
        apply_immediately: bool,
    ) -> Result<(), LayoutError>;
    fn force_page_break(&mut self) -> u32;
    /// Create a new page even when the current page is pristine.
    fn insert_blank_page(&mut self) -> u32;
    fn current_page_size(&mut self) -> Size;
    fn update_columns(&mut self, columns: &ColumnLayout);
}

fn round_ties_positive(x: f64) -> f64 {
    if !x.is_finite() || x == 0.0 {
        return x;
    }
    let f = x.floor();
    if x - f >= 0.5 { f + 1.0 } else { f }
}

fn max_propagating_nan(a: f64, b: f64) -> f64 {
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

fn schedule_margins(
    current: Option<&PartialPageMargins>,
    incoming: &PageMargins,
) -> PartialPageMargins {
    let mut merged = current.cloned().unwrap_or_default();
    merged.top = Some(max_propagating_nan(0.0, incoming.top));
    merged.right = Some(max_propagating_nan(0.0, incoming.right));
    merged.bottom = Some(max_propagating_nan(0.0, incoming.bottom));
    merged.left = Some(max_propagating_nan(0.0, incoming.left));
    if let Some(v) = incoming.header {
        merged.header = Some(max_propagating_nan(0.0, v));
    }
    if let Some(v) = incoming.footer {
        merged.footer = Some(max_propagating_nan(0.0, v));
    }
    merged
}

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

pub fn apply_section_break(
    block: &SectionBreakBlock,
    tracker: &SectionLayoutTracker,
) -> ApplySectionBreakResult {
    let mut updated = tracker.clone();
    let break_kind = block.break_type.unwrap_or(SectionBreakType::Continuous);

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

    let incoming_columns = SINGLE_COLUMN;

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
            },
            tracker: updated,
        };
    }

    ApplySectionBreakResult {
        outcome: SectionBreakOutcome {
            break_to_new_page: false,
            page_parity: None,
            open_column_region: false,
        },
        tracker: updated,
    }
}

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

pub fn resolve_next_page_size(tracker: &SectionLayoutTracker) -> Size {
    tracker
        .queued
        .page_size
        .clone()
        .unwrap_or_else(|| tracker.in_force.page_size.clone())
}

pub fn resolve_next_columns(tracker: &SectionLayoutTracker) -> ColumnLayout {
    tracker
        .queued
        .columns
        .clone()
        .unwrap_or_else(|| tracker.in_force.columns.clone())
}

const DEFAULT_MARGIN_PX: f64 = 96.0;

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

pub fn handle_section_break<P: SectionBreakPaginator>(
    _block: &SectionBreakBlock,
    paginator: &mut P,
    next_section_config: &SectionLayoutConfig,
    next_section_type: Option<SectionBreakType>,
) -> Result<(), LayoutError> {
    // ECMA-376 §17.6.22: w:type specifies how the NEXT section starts relative
    // to this one. Default is 'nextPage' when w:type is absent.
    let break_type = next_section_type.unwrap_or(SectionBreakType::NextPage);
    let page_size = Some(&next_section_config.page_size);
    let margins = Some(&next_section_config.margins);

    match break_type {
        SectionBreakType::NextPage => {
            paginator.update_page_layout(page_size, margins, true)?;
            paginator.force_page_break();
        }

        SectionBreakType::EvenPage => {
            paginator.update_page_layout(page_size, margins, true)?;
            let page_number = paginator.force_page_break();
            // If landed on odd page, add another page
            if !page_number.is_multiple_of(2) {
                paginator.insert_blank_page();
            }
        }

        SectionBreakType::OddPage => {
            paginator.update_page_layout(page_size, margins, true)?;
            let page_number = paginator.force_page_break();
            // If landed on even page, add another page
            if page_number.is_multiple_of(2) {
                paginator.insert_blank_page();
            }
        }

        SectionBreakType::Continuous => {
            let current_size = paginator.current_page_size();
            let next_size = &next_section_config.page_size;
            let page_size_changes = round_ties_positive(next_size.w)
                != round_ties_positive(current_size.w)
                || round_ties_positive(next_size.h) != round_ties_positive(current_size.h);
            if page_size_changes {
                paginator.update_page_layout(page_size, margins, true)?;
                paginator.force_page_break();
            } else {
                paginator.update_page_layout(page_size, margins, false)?;
            }
        }
    }

    // Update column layout for the next section
    let default_cols = default_columns();
    paginator.update_columns(
        next_section_config
            .columns
            .as_ref()
            .unwrap_or(&default_cols),
    );
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::BlockId;

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
        // Active geometry remains unchanged until promotion.
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
        // Active geometry has two columns; incoming geometry has one.
        // differs, so a continuous break opens a new column region.
        let tracker = create_section_layout_tracker(
            &margins(96.0, 96.0, 96.0, 96.0),
            &Size {
                w: 816.0,
                h: 1056.0,
            },
            Some(&ColumnLayout {
                count: 2.0,
                gap: 20.0,
                equal_width: None,
                separator: None,
            }),
        );
        let result = apply_section_break(&empty_break(), &tracker);
        assert!(!result.outcome.break_to_new_page);
        assert!(result.outcome.open_column_region);
        assert_eq!(result.tracker.queued.columns, Some(SINGLE_COLUMN));
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
        // Negative distances clamp to zero.
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
        // Scheduled margins preserve unspecified active values.
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
    fn round_ties_positive_handles_negative_halves() {
        assert_eq!(round_ties_positive(0.5), 1.0);
        assert_eq!(round_ties_positive(-0.5), 0.0);
        assert_eq!(round_ties_positive(-1.5), -1.0);
        assert_eq!(round_ties_positive(2.4), 2.0);
        assert_eq!(round_ties_positive(816.0), 816.0);
        // spec edge: closest double below 0.5 must not round up
        assert_eq!(round_ties_positive(0.49999999999999994), 0.0);
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
        UpdateColumns {
            count: f64,
            gap: f64,
        },
    }

    struct MockPaginator {
        page_size: Size,
        page_number: u32,
        pending_page_size: Option<Size>,
        calls: Vec<Call>,
    }

    impl MockPaginator {
        fn new(page_size: Size, page_number: u32) -> Self {
            MockPaginator {
                page_size,
                page_number,
                pending_page_size: None,
                calls: Vec::new(),
            }
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
                self.pending_page_size = None;
            } else if let Some(size) = page_size {
                self.pending_page_size = Some(size.clone());
            }
            Ok(())
        }

        fn force_page_break(&mut self) -> u32 {
            if let Some(size) = self.pending_page_size.take() {
                self.page_size = size;
            }
            self.page_number += 1;
            self.calls.push(Call::ForcePageBreak {
                new_page_number: self.page_number,
            });
            self.page_number
        }

        fn insert_blank_page(&mut self) -> u32 {
            self.page_number += 1;
            self.calls.push(Call::InsertBlankPage {
                new_page_number: self.page_number,
            });
            self.page_number
        }

        fn current_page_size(&mut self) -> Size {
            self.page_size.clone()
        }

        fn update_columns(&mut self, columns: &ColumnLayout) {
            self.calls.push(Call::UpdateColumns {
                count: columns.count,
                gap: columns.gap,
            });
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

        // sb2: back to portrait — promoted again
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

    // Equal rounded sizes do not promote a continuous break.
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
}
