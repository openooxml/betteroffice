use docx_layout::types::{
    BlockExtent, LayoutBlock, MeasuredBlock, ParagraphBlock, ParagraphExtent, Run, ShapeExtent,
    TypesetRow,
};
use serde_json::{Value, json};

pub(super) fn pagination_signature(measured: &MeasuredBlock) -> Option<Value> {
    parts_signature(&measured.block, &measured.measure)
}

fn strip_positions(value: &mut Value) {
    match value {
        Value::Object(fields) => {
            for key in ["pmStart", "pmEnd", "docStart", "docEnd"] {
                fields.remove(key);
            }
            for value in fields.values_mut() {
                strip_positions(value);
            }
        }
        Value::Array(values) => values.iter_mut().for_each(strip_positions),
        _ => {}
    }
}

fn line_signature(line: &TypesetRow) -> Value {
    let TypesetRow {
        head_run: _,
        head_char: _,
        tail_run: _,
        tail_char: _,
        width: _,
        ascent,
        descent,
        line_height,
        synthetic_fallback,
        left_offset,
        right_offset,
        segments,
        float_skip_before,
        marker_tab_offset: _,
        run_advances: _,
        cluster_advances: _,
        bidi_slices: _,
    } = line;
    json!({
        "ascent": ascent,
        "descent": descent,
        "lineHeight": line_height,
        "syntheticFallback": synthetic_fallback,
        "leftOffset": left_offset,
        "rightOffset": right_offset,
        "floatSkipBefore": float_skip_before,
        "segments": segments.as_ref().map(|segments| segments.iter().map(|segment| {
            json!([segment.left_offset, segment.available_width])
        }).collect::<Vec<_>>()),
    })
}

fn boundary(run: Option<&Run>, offset: usize, end: bool) -> Option<f64> {
    let run = run?;
    if let Run::Text(text) = run {
        return text
            .pm_start
            .map(|start| start + offset.min(text.text.encode_utf16().count()) as f64)
            .or_else(|| end.then_some(text.pm_end).flatten());
    }
    if end {
        run.pm_end()
            .or_else(|| run.pm_start().map(|start| start + 1.0))
    } else {
        run.pm_start()
    }
}

fn reference_line(
    paragraph: &ParagraphBlock,
    extent: &ParagraphExtent,
    position: f64,
) -> Option<usize> {
    let mut matched = None;
    for (index, line) in extent.lines.iter().enumerate() {
        let from = boundary(paragraph.runs.get(line.head_run), line.head_char, false);
        let to = boundary(paragraph.runs.get(line.tail_run), line.tail_char, true);
        let start = if index == 0 {
            paragraph.pm_start.or(from)
        } else {
            from.or(paragraph.pm_start)
        }?;
        let mut end = if index + 1 == extent.lines.len() {
            paragraph.pm_end.or(to)
        } else {
            to.or(paragraph.pm_end)
        }?;
        if end <= start {
            end = start + 1.0;
        }
        if start >= 0.0 && position >= start && position < end {
            if matched.is_some() {
                return None;
            }
            matched = Some(index);
        }
    }
    matched
}

fn paragraph_signature(paragraph: &ParagraphBlock, extent: &ParagraphExtent) -> Option<Value> {
    if !extent.total_height.is_finite()
        || extent.lines.iter().any(|line| {
            !line.ascent.is_finite()
                || !line.descent.is_finite()
                || !line.line_height.is_finite()
                || line
                    .float_skip_before
                    .is_some_and(|height| !height.is_finite())
        })
    {
        return None;
    }
    let mut block = serde_json::to_value(paragraph).ok()?;
    strip_positions(&mut block);
    let runs = block["runs"].as_array_mut()?;
    for (value, run) in runs.iter_mut().zip(&paragraph.runs) {
        if matches!(run, Run::Text(_)) {
            value.as_object_mut()?.remove("text");
        }
    }
    let mut references = Vec::new();
    let mut objects = Vec::new();
    for (index, run) in paragraph.runs.iter().enumerate() {
        if let Run::Image(image) = run {
            objects.push((index, reference_line(paragraph, extent, image.pm_start?)?));
        }
    }
    for reference in
        docx_layout::footnotes::collect_note_refs(&[LayoutBlock::Paragraph(paragraph.clone())])
    {
        references.push((
            reference.map_id(),
            reference_line(paragraph, extent, reference.pm_pos)?,
        ));
    }
    Some(json!({
        "kind": "paragraph",
        "block": block,
        "spacingBefore": docx_layout::paragraph_spacing::get_spacing_before(paragraph),
        "spacingAfter": docx_layout::paragraph_spacing::get_spacing_after(paragraph),
        "empty": paragraph.runs.is_empty() || matches!(paragraph.runs.as_slice(),
            [Run::Text(text)] if text.text.is_empty()),
        "totalHeight": extent.total_height,
        "lines": extent.lines.iter().map(line_signature).collect::<Vec<_>>(),
        "references": references,
        "objects": objects,
    }))
}

fn parts_signature(block: &LayoutBlock, measure: &BlockExtent) -> Option<Value> {
    if let (LayoutBlock::Paragraph(paragraph), BlockExtent::Paragraph(extent)) = (block, measure) {
        return paragraph_signature(paragraph, extent);
    }
    let mut block_value = serde_json::to_value(block).ok()?;
    let mut extent_value = serde_json::to_value(measure).ok()?;
    strip_positions(&mut block_value);
    match (block, measure) {
        (LayoutBlock::Table(table), BlockExtent::Table(extent)) => {
            if table.rows.len() != extent.rows.len() {
                return None;
            }
            for (row_index, (row, row_extent)) in table.rows.iter().zip(&extent.rows).enumerate() {
                if row.cells.len() != row_extent.cells.len() {
                    return None;
                }
                for (cell_index, (cell, cell_extent)) in
                    row.cells.iter().zip(&row_extent.cells).enumerate()
                {
                    if cell.blocks.len() != cell_extent.blocks.len() {
                        return None;
                    }
                    let children = cell
                        .blocks
                        .iter()
                        .zip(&cell_extent.blocks)
                        .map(|(block, measure)| parts_signature(block, measure))
                        .collect::<Option<Vec<_>>>()?;
                    block_value["rows"][row_index]["cells"][cell_index]["blocks"] = json!(children);
                    extent_value["rows"][row_index]["cells"][cell_index]["blocks"] = Value::Null;
                }
            }
        }
        (LayoutBlock::TextBox(textbox), BlockExtent::TextBox(extent)) => {
            if textbox.content.len() != extent.inner_measures.len() {
                return None;
            }
            block_value["content"] = json!(
                textbox
                    .content
                    .iter()
                    .zip(&extent.inner_measures)
                    .map(|(block, measure)| paragraph_signature(block, measure))
                    .collect::<Option<Vec<_>>>()?
            );
            extent_value["innerMeasures"] = Value::Null;
        }
        (LayoutBlock::Shape(shape), BlockExtent::Shape(extent)) => {
            let blocks = shape.inner_text.as_deref().unwrap_or_default();
            let measures = extent.inner_measures.as_deref().unwrap_or_default();
            if blocks.len() != measures.len() {
                return None;
            }
            block_value["innerText"] = json!(
                blocks
                    .iter()
                    .zip(measures)
                    .map(|(block, measure)| paragraph_signature(block, measure))
                    .collect::<Option<Vec<_>>>()?
            );
            block_value.as_object_mut()?.remove("innerMeasures");
            extent_value["innerMeasures"] = Value::Null;
            block_value["children"] = json!(
                shape
                    .children
                    .iter()
                    .map(|child| {
                        parts_signature(
                            &LayoutBlock::Shape(child.clone()),
                            &BlockExtent::Shape(ShapeExtent {
                                width: child.width,
                                height: child.height,
                                inner_measures: child.inner_measures.clone(),
                            }),
                        )
                    })
                    .collect::<Option<Vec<_>>>()?
            );
        }
        (LayoutBlock::Image(_), BlockExtent::Image(_))
        | (LayoutBlock::Chart(_), BlockExtent::Chart(_))
        | (LayoutBlock::SectionBreak(_), BlockExtent::SectionBreak)
        | (LayoutBlock::PageBreak(_), BlockExtent::PageBreak)
        | (LayoutBlock::ColumnBreak(_), BlockExtent::ColumnBreak) => {}
        _ => return None,
    }
    Some(json!({"block": block_value, "measure": extent_value}))
}
