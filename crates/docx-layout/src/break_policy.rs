//! Pre-placement break policy.

use crate::keep_together::paragraph_breaks_before_run;
use crate::types::LayoutBlock;

/// Why a paragraph opens a fresh page.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AuthoredBreak {
    /// `w:pageBreakBefore` (ECMA-376 §17.3.1.23) on the paragraph itself.
    PageBreakBefore,
    /// A `w:br w:type="page"` run opening the paragraph (§17.3.3.1).
    HardBreakRun,
}

impl AuthoredBreak {
    /// Word keeps the paragraph's space-before across a hard break run —
    /// measured against Word 16.113, independent of how full the page was.
    pub fn keeps_leading_spacing(self) -> bool {
        matches!(self, Self::HardBreakRun)
    }
}

/// Why a block forces a fresh page before it is placed, if it does.
pub fn breaks_before_block(block: &LayoutBlock) -> Option<AuthoredBreak> {
    let (property, run) = paragraph_breaks_before_run(block);
    if run {
        Some(AuthoredBreak::HardBreakRun)
    } else if property {
        Some(AuthoredBreak::PageBreakBefore)
    } else {
        None
    }
}

/// Geometry a keep-with-next group is weighed against at the page cursor.
/// `group_height` is the space the whole group (plus its follower's first
/// line) needs; `available_height` is what remains in the current column; and
/// `page_content_height` is the content height of a blank page/column.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeepWithNextFit {
    pub group_height: f64,
    pub available_height: f64,
    pub page_content_height: f64,
    pub page_has_content: bool,
}

/// Whether a keep-with-next group at the cursor must advance to a fresh page
/// before its head is placed.
///
/// Derivation from Word's w:keepNext behavior (§17.3.1.15 states the intent
/// but not the algorithm): a keepNext paragraph stays on the same page as the
/// START of its bound follower, so when the group cannot finish where the
/// cursor stands, the whole group moves to the next page. Each early return
/// below is a situation where moving is wrong:
///
/// - a group taller than an empty page has no intact placement anywhere;
///   advancing would re-fail on every subsequent page, so Word lets it split
///   at the cursor instead
/// - a group that finishes in the remaining space is already satisfied
/// - at the top of an empty page/column the cursor cannot retreat any further;
///   there is nothing above the group to detach from, and advancing would only
///   emit a blank page, so Word splits in place
pub fn keep_with_next_group_must_advance(fit: KeepWithNextFit) -> bool {
    let intact_placement_exists = fit.group_height <= fit.page_content_height;
    if !intact_placement_exists {
        return false;
    }

    let finishes_at_cursor = fit.group_height <= fit.available_height;
    if finishes_at_cursor {
        return false;
    }

    fit.page_has_content
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::{BlockId, ParagraphAttrs, ParagraphBlock};

    // minimal block stubs — the predicates only read the kind and a couple of attrs
    fn paragraph(attrs: Option<ParagraphAttrs>) -> LayoutBlock {
        LayoutBlock::Paragraph(ParagraphBlock {
            sdt_groups: None,
            id: BlockId::Num(0.0),
            para_id: None,
            runs: vec![],
            attrs,
            pm_start: None,
            pm_end: None,
        })
    }

    #[test]
    fn breaks_before_is_true_for_a_paragraph_with_page_break_before() {
        let attrs = ParagraphAttrs {
            page_break_before: Some(true),
            ..Default::default()
        };
        assert_eq!(
            breaks_before_block(&paragraph(Some(attrs))),
            Some(AuthoredBreak::PageBreakBefore)
        );
    }

    #[test]
    fn breaks_before_is_false_for_a_paragraph_without_page_break_before() {
        assert_eq!(breaks_before_block(&paragraph(None)), None);
        let attrs = ParagraphAttrs {
            page_break_before: Some(false),
            ..Default::default()
        };
        assert_eq!(breaks_before_block(&paragraph(Some(attrs))), None);
    }

    #[test]
    fn a_leading_hard_break_run_reports_a_hard_break() {
        let attrs = ParagraphAttrs {
            page_break_before_run: Some(true),
            ..Default::default()
        };
        assert_eq!(
            breaks_before_block(&paragraph(Some(attrs))),
            Some(AuthoredBreak::HardBreakRun)
        );
        assert!(AuthoredBreak::HardBreakRun.keeps_leading_spacing());
        assert!(!AuthoredBreak::PageBreakBefore.keeps_leading_spacing());
    }

    #[test]
    fn breaks_before_is_false_for_a_non_paragraph_block() {
        assert_eq!(breaks_before_block(&LayoutBlock::Unsupported), None);
    }

    #[test]
    fn advances_an_intact_group_off_a_straddled_boundary() {
        // fits a blank page, does not fit the remaining space, page has content
        assert!(keep_with_next_group_must_advance(KeepWithNextFit {
            group_height: 400.0,
            available_height: 200.0,
            page_content_height: 600.0,
            page_has_content: true,
        }));
    }

    #[test]
    fn lets_an_oversized_group_split_rather_than_loop_forever() {
        // taller than a whole page — the fit clause fails, so it is NOT advanced
        assert!(!keep_with_next_group_must_advance(KeepWithNextFit {
            group_height: 700.0,
            available_height: 200.0,
            page_content_height: 600.0,
            page_has_content: true,
        }));
    }

    #[test]
    fn stays_put_when_the_group_already_fits_the_remaining_space() {
        assert!(!keep_with_next_group_must_advance(KeepWithNextFit {
            group_height: 150.0,
            available_height: 200.0,
            page_content_height: 600.0,
            page_has_content: true,
        }));
    }

    #[test]
    fn does_not_advance_when_the_page_is_still_empty() {
        assert!(!keep_with_next_group_must_advance(KeepWithNextFit {
            group_height: 400.0,
            available_height: 200.0,
            page_content_height: 600.0,
            page_has_content: false,
        }));
    }

    #[test]
    fn treats_a_group_exactly_the_page_height_as_fitting_boundary() {
        // group_height == page_content_height satisfies the <= fit clause
        assert!(keep_with_next_group_must_advance(KeepWithNextFit {
            group_height: 600.0,
            available_height: 200.0,
            page_content_height: 600.0,
            page_has_content: true,
        }));
    }

    #[test]
    fn does_not_advance_when_the_group_exactly_fits_the_remaining_space_boundary() {
        // group_height == available_height fails the strict > clause
        assert!(!keep_with_next_group_must_advance(KeepWithNextFit {
            group_height: 200.0,
            available_height: 200.0,
            page_content_height: 600.0,
            page_has_content: true,
        }));
    }
}
