//! Paragraph spacing resolution.

use crate::types::{LayoutBlock, ParagraphBlock, Run, ShapeBlock};

/// Resolve paragraph line units against the section pitch.
pub fn resolve_line_unit_spacing(block: &mut LayoutBlock, line_px: f64) {
    match block {
        LayoutBlock::Paragraph(paragraph) => resolve_paragraph_line_spacing(paragraph, line_px),
        LayoutBlock::Table(table) => {
            for row in &mut table.rows {
                for cell in &mut row.cells {
                    for block in &mut cell.blocks {
                        resolve_line_unit_spacing(block, line_px);
                    }
                }
            }
        }
        LayoutBlock::TextBox(text_box) => {
            for paragraph in &mut text_box.content {
                resolve_paragraph_line_spacing(paragraph, line_px);
            }
        }
        LayoutBlock::Shape(shape) => resolve_shape_line_spacing(shape, line_px),
        _ => {}
    }
}

fn resolve_shape_line_spacing(shape: &mut ShapeBlock, line_px: f64) {
    if let Some(paragraphs) = &mut shape.inner_text {
        for paragraph in paragraphs {
            resolve_paragraph_line_spacing(paragraph, line_px);
        }
    }
    for child in &mut shape.children {
        resolve_shape_line_spacing(child, line_px);
    }
}

fn resolve_paragraph_line_spacing(paragraph: &mut ParagraphBlock, line_px: f64) {
    let Some(spacing) = paragraph
        .attrs
        .as_mut()
        .and_then(|attrs| attrs.spacing.as_mut())
    else {
        return;
    };
    if let Some(lines) = spacing
        .before_lines
        .filter(|lines| lines.is_finite() && *lines > 0.0)
    {
        spacing.before = Some(lines * line_px / 100.0);
    }
    if let Some(lines) = spacing
        .after_lines
        .filter(|lines| lines.is_finite() && *lines > 0.0)
    {
        spacing.after = Some(lines * line_px / 100.0);
    }
}

pub(crate) fn is_empty_paragraph(block: &ParagraphBlock) -> bool {
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

/// Returns effective leading spacing.
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

/// Returns effective trailing spacing.
pub fn get_spacing_after(block: &ParagraphBlock) -> f64 {
    block
        .attrs
        .as_ref()
        .and_then(|a| a.spacing.as_ref())
        .and_then(|s| s.after)
        .unwrap_or(0.0)
}
