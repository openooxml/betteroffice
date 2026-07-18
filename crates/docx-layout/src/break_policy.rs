//! Port of `packages/core/src/layout/pagination/breakPolicy.ts`.
//!
//! Exported fns (1:1 with the TS module's exports):
//! - `breaks_before_block`               ← `breaksBeforeBlock(block)`
//! - `keep_with_next_group_must_advance` ← `keepWithNextGroupMustAdvance(fit)`
//! - `KeepWithNextFit`                   ← `KeepWithNextFit` (exported type)
//!
//! Pre-placement break policy — the decisions the placement walk consults
//! before a block is placed, expressed as small pure predicates. The place
//! loop calls these through the seams in `hooks.rs`.

use crate::keep_together::paragraph_breaks_before;
use crate::types::LayoutBlock;

/// Whether a block forces a fresh page before it is placed. Today only a
/// paragraph with `w:pageBreakBefore` (ECMA-376 §17.3.1.23) does.
pub fn breaks_before_block(block: &LayoutBlock) -> bool {
    paragraph_breaks_before(block)
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

// ---- tests (ported from breakPolicy.test.ts) ---------------------------------

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
        assert!(breaks_before_block(&paragraph(Some(attrs))));
    }

    #[test]
    fn breaks_before_is_false_for_a_paragraph_without_page_break_before() {
        assert!(!breaks_before_block(&paragraph(None)));
        let attrs = ParagraphAttrs {
            page_break_before: Some(false),
            ..Default::default()
        };
        assert!(!breaks_before_block(&paragraph(Some(attrs))));
    }

    #[test]
    fn breaks_before_is_false_for_a_non_paragraph_block() {
        assert!(!breaks_before_block(&LayoutBlock::Unsupported));
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
