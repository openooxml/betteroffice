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

/// Resolve the section's document-grid snap pitch onto every paragraph in
/// the block tree (recursing into tables, text boxes and shapes, mirroring
/// [`resolve_line_unit_spacing`]). `None` clears any pitch, so sections
/// without an activating grid measure exactly as before.
pub fn resolve_doc_grid_pitch(block: &mut LayoutBlock, pitch_px: Option<f64>) {
    match block {
        LayoutBlock::Paragraph(paragraph) => {
            if let Some(attrs) = paragraph.attrs.as_mut() {
                attrs.doc_grid_pitch_px = pitch_px;
            } else if pitch_px.is_some() {
                paragraph.attrs = Some(crate::types::ParagraphAttrs {
                    doc_grid_pitch_px: pitch_px,
                    ..crate::types::ParagraphAttrs::default()
                });
            }
        }
        LayoutBlock::Table(table) => {
            for row in &mut table.rows {
                for cell in &mut row.cells {
                    for block in &mut cell.blocks {
                        resolve_doc_grid_pitch(block, pitch_px);
                    }
                }
            }
        }
        LayoutBlock::TextBox(text_box) => {
            for paragraph in &mut text_box.content {
                if let Some(attrs) = paragraph.attrs.as_mut() {
                    attrs.doc_grid_pitch_px = pitch_px;
                } else if pitch_px.is_some() {
                    paragraph.attrs = Some(crate::types::ParagraphAttrs {
                        doc_grid_pitch_px: pitch_px,
                        ..crate::types::ParagraphAttrs::default()
                    });
                }
            }
        }
        LayoutBlock::Shape(shape) => resolve_shape_doc_grid_pitch(shape, pitch_px),
        _ => {}
    }
}

fn resolve_shape_doc_grid_pitch(shape: &mut ShapeBlock, pitch_px: Option<f64>) {
    if let Some(paragraphs) = &mut shape.inner_text {
        for paragraph in paragraphs {
            if let Some(attrs) = paragraph.attrs.as_mut() {
                attrs.doc_grid_pitch_px = pitch_px;
            } else if pitch_px.is_some() {
                paragraph.attrs = Some(crate::types::ParagraphAttrs {
                    doc_grid_pitch_px: pitch_px,
                    ..crate::types::ParagraphAttrs::default()
                });
            }
        }
    }
    for child in &mut shape.children {
        resolve_shape_doc_grid_pitch(child, pitch_px);
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

fn effective_style_id(block: &ParagraphBlock) -> &str {
    let attrs = block.attrs.as_ref();
    attrs
        .and_then(|attrs| {
            attrs
                .effective_style_id
                .as_deref()
                .filter(|style| !style.is_empty())
        })
        .or_else(|| {
            attrs.and_then(|attrs| attrs.style_id.as_deref().filter(|style| !style.is_empty()))
        })
        .unwrap_or("")
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
    let same_style = effective_style_id(c) == effective_style_id(n);
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

pub fn apply_contextual_spacing_blocks(blocks: &mut [LayoutBlock]) {
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::{BlockId, ParagraphAttrs, ParagraphBlock, ParagraphSpacing};

    fn paragraph(style_id: Option<&str>, effective_style_id: Option<&str>) -> LayoutBlock {
        LayoutBlock::Paragraph(ParagraphBlock {
            sdt_groups: None,
            id: BlockId::Str("p".to_owned()),
            para_id: None,
            runs: Vec::new(),
            attrs: Some(ParagraphAttrs {
                style_id: style_id.map(str::to_owned),
                effective_style_id: effective_style_id.map(str::to_owned),
                contextual_spacing: Some(true),
                spacing: Some(ParagraphSpacing {
                    before: Some(8.0),
                    after: Some(8.0),
                    ..ParagraphSpacing::default()
                }),
                ..ParagraphAttrs::default()
            }),
            pm_start: None,
            pm_end: None,
        })
    }

    fn after(block: &LayoutBlock) -> f64 {
        match block {
            LayoutBlock::Paragraph(paragraph) => paragraph
                .attrs
                .as_ref()
                .and_then(|attrs| attrs.spacing.as_ref())
                .and_then(|spacing| spacing.after)
                .unwrap_or(-1.0),
            _ => -1.0,
        }
    }

    #[test]
    fn effective_default_matches_explicit_both_orders() {
        for (first, second) in [
            (
                paragraph(None, Some("Normal")),
                paragraph(Some("Normal"), None),
            ),
            (
                paragraph(Some("Normal"), None),
                paragraph(None, Some("Normal")),
            ),
        ] {
            let mut blocks = vec![first, second];
            apply_contextual_spacing_blocks(&mut blocks);
            assert_eq!(after(&blocks[0]), 0.0);
        }
    }

    #[test]
    fn custom_default_keeps_explicit_normal_distinct() {
        let mut blocks = vec![
            paragraph(None, Some("BodyDefault")),
            paragraph(Some("Normal"), None),
        ];
        apply_contextual_spacing_blocks(&mut blocks);
        assert_eq!(after(&blocks[0]), 8.0);
        let mut matched = vec![
            paragraph(None, Some("BodyDefault")),
            paragraph(Some("BodyDefault"), None),
        ];
        apply_contextual_spacing_blocks(&mut matched);
        assert_eq!(after(&matched[0]), 0.0);
    }

    #[test]
    fn differing_style_keeps_gap() {
        let mut blocks = vec![
            paragraph(None, Some("Normal")),
            paragraph(Some("Different"), None),
        ];
        apply_contextual_spacing_blocks(&mut blocks);
        assert_eq!(after(&blocks[0]), 8.0);
    }
}
