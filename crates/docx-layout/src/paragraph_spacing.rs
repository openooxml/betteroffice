//! Paragraph spacing resolution.

use crate::types::{LayoutBlock, MeasuredBlock, ParagraphBlock, Run, ShapeBlock};

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

pub(crate) fn contextual_spacing_pair(curr: &mut LayoutBlock, next: &mut LayoutBlock) {
    let LayoutBlock::Paragraph(c) = curr else {
        return;
    };
    let next_is_table = matches!(next, LayoutBlock::Table(_));
    let n = match next {
        LayoutBlock::Paragraph(paragraph) => paragraph,
        LayoutBlock::Table(table) if table.floating.is_none() && is_empty_paragraph(c) => {
            let Some(LayoutBlock::Paragraph(paragraph)) = table
                .rows
                .first_mut()
                .and_then(|row| row.cells.first_mut())
                .and_then(|cell| cell.blocks.first_mut())
            else {
                return;
            };
            paragraph
        }
        _ => return,
    };
    let same_style = c
        .attrs
        .as_ref()
        .and_then(|attrs| attrs.style_id.as_deref())
        .unwrap_or("")
        == n.attrs
            .as_ref()
            .and_then(|attrs| attrs.style_id.as_deref())
            .unwrap_or("");
    if !same_style {
        return;
    }
    if let Some(ca) = &mut c.attrs
        && ca.contextual_spacing.unwrap_or(false)
        && let Some(spacing) = &mut ca.spacing
    {
        spacing.after = Some(0.0);
    }
    if !next_is_table
        && let Some(na) = &mut n.attrs
        && na.contextual_spacing.unwrap_or(false)
        && let Some(spacing) = &mut na.spacing
    {
        spacing.before = Some(0.0);
    }
}

pub(crate) fn apply_contextual_spacing_blocks(blocks: &mut [LayoutBlock]) {
    for i in 0..blocks.len().saturating_sub(1) {
        let (head, tail) = blocks.split_at_mut(i + 1);
        contextual_spacing_pair(&mut head[i], &mut tail[0]);
    }
    for block in blocks.iter_mut() {
        if let LayoutBlock::Table(table) = block {
            for row in &mut table.rows {
                for cell in &mut row.cells {
                    apply_contextual_spacing_blocks(&mut cell.blocks);
                }
            }
        }
    }
}

pub(crate) fn apply_contextual_spacing_measured(measured: &mut [MeasuredBlock]) {
    for i in 0..measured.len().saturating_sub(1) {
        let (head, tail) = measured.split_at_mut(i + 1);
        contextual_spacing_pair(&mut head[i].block, &mut tail[0].block);
    }
    for mb in measured.iter_mut() {
        if let LayoutBlock::Table(table) = &mut mb.block {
            for row in &mut table.rows {
                for cell in &mut row.cells {
                    apply_contextual_spacing_blocks(&mut cell.blocks);
                }
            }
        }
    }
}
