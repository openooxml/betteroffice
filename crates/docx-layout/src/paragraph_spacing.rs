//! Port of `packages/core/src/layout/pagination/paragraphSpacing.ts`.
//!
//! Exported fns (1:1 with the TS module's exports):
//! - `get_spacing_before` ← `getSpacingBefore(block)`
//! - `get_spacing_after`  ← `getSpacingAfter(block)`

use crate::types::{ParagraphBlock, Run};

// mirrors paragraphSpacing.ts isEmptyParagraph
fn is_empty_paragraph(block: &ParagraphBlock) -> bool {
    if block.runs.is_empty() {
        return true;
    }
    if block.runs.len() != 1 {
        return false;
    }
    match &block.runs[0] {
        Run::Text(r) => r.text.is_empty(),
        _ => false,
    }
}

/// Word collapses style-inherited spacing on empty paragraphs (only direct
/// formatting survives). `spacingExplicit` tracks which side was set inline.
pub fn get_spacing_before(block: &ParagraphBlock) -> f64 {
    let value = block
        .attrs
        .as_ref()
        .and_then(|a| a.spacing.as_ref())
        .and_then(|s| s.before)
        .unwrap_or(0.0);
    let explicit = block
        .attrs
        .as_ref()
        .and_then(|a| a.spacing_explicit.as_ref())
        .and_then(|e| e.before)
        .unwrap_or(false);
    if is_empty_paragraph(block) && !explicit {
        return 0.0;
    }
    value
}

/// TS `getSpacingAfter`.
pub fn get_spacing_after(block: &ParagraphBlock) -> f64 {
    let value = block
        .attrs
        .as_ref()
        .and_then(|a| a.spacing.as_ref())
        .and_then(|s| s.after)
        .unwrap_or(0.0);
    let explicit = block
        .attrs
        .as_ref()
        .and_then(|a| a.spacing_explicit.as_ref())
        .and_then(|e| e.after)
        .unwrap_or(false);
    if is_empty_paragraph(block) && !explicit {
        return 0.0;
    }
    value
}
