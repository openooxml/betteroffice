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
    for run in &mut paragraph.runs {
        if let Run::Image(image) = run
            && let Some(shape) = &mut image.inline_shape
        {
            resolve_shape_line_spacing(shape, line_px);
        }
    }
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
        LayoutBlock::Paragraph(paragraph) => resolve_paragraph_doc_grid_pitch(paragraph, pitch_px),
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
                resolve_paragraph_doc_grid_pitch(paragraph, pitch_px);
            }
        }
        LayoutBlock::Shape(shape) => resolve_shape_doc_grid_pitch(shape, pitch_px),
        _ => {}
    }
}

fn resolve_shape_doc_grid_pitch(shape: &mut ShapeBlock, pitch_px: Option<f64>) {
    if let Some(paragraphs) = &mut shape.inner_text {
        for paragraph in paragraphs {
            resolve_paragraph_doc_grid_pitch(paragraph, pitch_px);
        }
    }
    for child in &mut shape.children {
        resolve_shape_doc_grid_pitch(child, pitch_px);
    }
}

fn resolve_paragraph_doc_grid_pitch(paragraph: &mut ParagraphBlock, pitch_px: Option<f64>) {
    if let Some(attrs) = paragraph.attrs.as_mut() {
        attrs.doc_grid_pitch_px = pitch_px;
    } else if pitch_px.is_some() {
        paragraph.attrs = Some(crate::types::ParagraphAttrs {
            doc_grid_pitch_px: pitch_px,
            ..crate::types::ParagraphAttrs::default()
        });
    }
    for run in &mut paragraph.runs {
        if let Run::Image(image) = run
            && let Some(shape) = &mut image.inline_shape
        {
            resolve_shape_doc_grid_pitch(shape, pitch_px);
        }
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

fn effective_num_id(block: &ParagraphBlock) -> Option<f64> {
    block
        .attrs
        .as_ref()?
        .num_pr
        .as_ref()?
        .num_id
        .filter(|id| id.is_finite() && *id > 0.0)
}

pub fn suppress_document_auto_before(block: &mut LayoutBlock) {
    if let LayoutBlock::Paragraph(paragraph) = block
        && let Some(spacing) = paragraph
            .attrs
            .as_mut()
            .and_then(|attrs| attrs.spacing.as_mut())
        && spacing.before_auto == Some(true)
    {
        spacing.before = Some(0.0);
    }
}

pub fn suppress_break_auto_before(block: &mut LayoutBlock, previous: Option<&LayoutBlock>) {
    let LayoutBlock::Paragraph(paragraph) = block else {
        return;
    };
    let breaks_before = paragraph.attrs.as_ref().is_some_and(|attrs| {
        attrs.page_break_before == Some(true) || attrs.page_break_before_run == Some(true)
    }) || matches!(previous, Some(LayoutBlock::PageBreak(_)));
    if !breaks_before {
        return;
    }
    if let Some(spacing) = paragraph
        .attrs
        .as_mut()
        .and_then(|attrs| attrs.spacing.as_mut())
        && spacing.before_auto == Some(true)
    {
        spacing.before = Some(0.0);
    }
}

pub fn suppress_list_auto_spacing(
    block: &mut LayoutBlock,
    previous: Option<&LayoutBlock>,
    next: Option<&LayoutBlock>,
) {
    let LayoutBlock::Paragraph(paragraph) = block else {
        return;
    };
    let Some(num_id) = effective_num_id(paragraph) else {
        return;
    };
    let same_list = |neighbor: Option<&LayoutBlock>| {
        matches!(neighbor, Some(LayoutBlock::Paragraph(other))
            if effective_num_id(other) == Some(num_id))
    };
    let Some(spacing) = paragraph
        .attrs
        .as_mut()
        .and_then(|attrs| attrs.spacing.as_mut())
    else {
        return;
    };
    if spacing.before_auto == Some(true) && same_list(previous) {
        spacing.before = Some(0.0);
    }
    if spacing.after_auto == Some(true) && same_list(next) {
        spacing.after = Some(0.0);
    }
}

/// The indices of the blocks that take part in the text flow; anchored
/// drawings are lowered ahead of their paragraph and sit outside it.
fn flow_indices<T>(items: &[T], block: impl Fn(&T) -> &LayoutBlock) -> Vec<usize> {
    (0..items.len())
        .filter(|&index| crate::header_footer::contributes_to_flow(block(&items[index])))
        .collect()
}

/// Splits `items` into the flow neighbours of `flow[k]` and the item itself.
fn flow_window<'a, T>(
    items: &'a mut [T],
    flow: &[usize],
    k: usize,
) -> (Option<&'a T>, &'a mut T, Option<&'a T>) {
    let index = flow[k];
    let (head, tail) = items.split_at_mut(index);
    let (current, rest) = tail.split_first_mut().unwrap();
    let previous = k.checked_sub(1).map(|j| &head[flow[j]]);
    let next = flow.get(k + 1).map(|&n| &rest[n - index - 1]);
    (previous, current, next)
}

pub fn apply_list_auto_spacing_blocks(blocks: &mut [LayoutBlock]) {
    let flow = flow_indices(blocks, |block| block);
    for k in 0..flow.len() {
        let (previous, current, next) = flow_window(blocks, &flow, k);
        suppress_list_auto_spacing(current, previous, next);
    }
    for block in blocks.iter_mut() {
        if let LayoutBlock::Table(table) = block {
            for row in &mut table.rows {
                for cell in &mut row.cells {
                    apply_list_auto_spacing_blocks(&mut cell.blocks);
                }
            }
        }
    }
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
    apply_list_auto_spacing_blocks(blocks);
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
    let flow = flow_indices(measured, |entry| &entry.block);
    if let Some(&first) = flow.first() {
        suppress_document_auto_before(&mut measured[first].block);
    }
    for k in 0..flow.len() {
        let (previous, current, next) = flow_window(measured, &flow, k);
        suppress_break_auto_before(&mut current.block, previous.map(|entry| &entry.block));
        suppress_list_auto_spacing(
            &mut current.block,
            previous.map(|entry| &entry.block),
            next.map(|entry| &entry.block),
        );
    }
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
    use serde_json::{Value, json};
    use std::collections::BTreeMap;

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
    fn inline_shapes_resolve_section_spacing_and_grid_recursively() {
        let fonts = crate::MeasureFonts::default();
        let _fonts = fonts.enter();
        let font = crate::register_measure_font_bytes(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
        let config = crate::measure_blocks::MeasurementConfig {
            font_chains: BTreeMap::from([("liberation sans|0|0".to_owned(), vec![font])]),
            defaults: json!({"fontFamily":"Liberation Sans","fontSize":12}),
            ..Default::default()
        };
        let paragraph = json!({
            "kind":"paragraph", "id":"inner", "runs":[{"kind":"text","text":"Box text"}],
            "attrs":{"spacing":{"beforeLines":100,"afterLines":200,"before":16,"after":32}}
        });
        let shape = |id: &str, inner_text: Vec<Value>, children: Vec<Value>| {
            json!({
                "id":id,"shapeType":"rect","geometryPath":[],"width":100,"height":100,
                "innerText":inner_text,"children":children
            })
        };
        let nested = shape("nested", vec![paragraph.clone()], vec![]);
        let carrier = json!({
            "kind":"paragraph","id":"carrier","runs":[{
                "kind":"image","src":"","width":100,"height":100,"inlineShape":nested
            }]
        });
        let child = shape("child", vec![paragraph.clone()], vec![]);
        let mut block: LayoutBlock = serde_json::from_value(json!({
            "kind":"paragraph","id":"outer","runs":[{
                "kind":"image","src":"","width":100,"height":100,
                "inlineShape":shape("box", vec![paragraph, carrier], vec![child])
            }]
        }))
        .unwrap();
        resolve_line_unit_spacing(&mut block, 32.0);
        resolve_doc_grid_pitch(&mut block, Some(32.0));
        crate::measure_blocks::measure_block(&mut block, 200.0, &config).unwrap();
        let value = serde_json::to_value(&block).unwrap();
        let shape = &value["runs"][0]["inlineShape"];
        assert_eq!(shape["innerMeasures"][0]["lines"][0]["lineHeight"], 32.0);
        for paragraph in [
            &shape["innerText"][0],
            &shape["innerText"][1]["runs"][0]["inlineShape"]["innerText"][0],
            &shape["children"][0]["innerText"][0],
        ] {
            assert_eq!(paragraph["attrs"]["spacing"]["before"], 32.0);
            assert_eq!(paragraph["attrs"]["spacing"]["after"], 64.0);
            assert_eq!(paragraph["attrs"]["docGridPitchPx"], 32.0);
        }
        resolve_doc_grid_pitch(&mut block, None);
        let value = serde_json::to_value(&block).unwrap();
        let shape = &value["runs"][0]["inlineShape"];
        for paragraph in [
            &shape["innerText"][0],
            &shape["innerText"][1]["runs"][0]["inlineShape"]["innerText"][0],
            &shape["children"][0]["innerText"][0],
        ] {
            assert!(paragraph["attrs"]["docGridPitchPx"].is_null());
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
    fn contextual_spacing_applied_again_changes_nothing() {
        let measured = |block| crate::types::MeasuredBlock {
            block,
            measure: crate::types::BlockExtent::Paragraph(crate::types::ParagraphExtent {
                lines: Vec::new(),
                total_height: 0.0,
            }),
        };
        let mut once = vec![
            measured(paragraph(Some("Normal"), None)),
            measured(paragraph(Some("Normal"), None)),
            measured(paragraph(Some("Other"), None)),
            measured(paragraph(Some("Other"), None)),
        ];
        apply_contextual_spacing_measured(&mut once);
        let mut twice = once.clone();
        apply_contextual_spacing_measured(&mut twice);
        assert_eq!(after(&once[0].block), 0.0);
        assert_eq!(
            serde_json::to_value(&once).unwrap(),
            serde_json::to_value(&twice).unwrap()
        );
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
