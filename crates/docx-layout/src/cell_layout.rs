//! Vertical table-cell layout with collapsed paragraph spacing.

use serde::Serialize;

use crate::types::{BlockExtent, FloatingTablePosition, LayoutBlock};

pub(crate) fn nested_table_float_offset(position: Option<&FloatingTablePosition>) -> Option<f64> {
    let position = position?;
    (position.vert_anchor.as_deref() == Some("text")
        && matches!(
            position.horz_anchor.as_deref(),
            None | Some("text" | "margin")
        )
        && position.tblp_x.is_none_or(f64::is_finite)
        && matches!(
            position.tblp_x_spec.as_deref(),
            None | Some("left" | "center" | "right")
        )
        && position.tblp_y_spec.is_none())
    .then_some(position.tblp_y)
    .flatten()
    .filter(|offset| offset.is_finite() && *offset >= 0.0)
}

/// Whether a cell drawing is an overlay the row must not grow for.
///
/// Word paints a `wrapNone` anchor (`behind` / `inFront`) over the cell and
/// reserves nothing for it, in the row height or in the flow below. A wrapping
/// anchor still displaces cell text and keeps its space.
pub(crate) fn cell_overlay_drawing(positioned: bool, wrap_type: Option<&str>) -> bool {
    positioned
        && !matches!(
            wrap_type,
            Some("square" | "tight" | "through" | "topAndBottom")
        )
}

/// Leading-edge shift for `w:tblInd` under pre-2013 compatibility.
///
/// Word 2013 changed what `w:tblInd` is measured to: in compatibilityMode <= 14
/// a left-aligned table's leading border sits at
/// `margin + tblInd - resolvedLeftCellMargin`, while in mode 15 it sits at
/// `margin + tblInd`. The resolved margin is the table-level `w:tblCellMar`
/// left (already cascaded table -> style -> default style by the seed).
/// Centered/right tables never shift; an absent mode defaults to 12 (shifts).
pub(crate) fn table_compat_leading_shift(
    justification: Option<&str>,
    compatibility_mode: Option<u8>,
    cell_margin_left: Option<f64>,
) -> f64 {
    if matches!(justification, Some("center" | "right")) {
        return 0.0;
    }
    if compatibility_mode.unwrap_or(12) > 14 {
        return 0.0;
    }
    match cell_margin_left {
        Some(margin) if margin.is_finite() && margin > 0.0 => margin,
        _ => 0.0,
    }
}

pub(crate) fn nested_table_horizontal_offset(
    position: Option<&FloatingTablePosition>,
    justification: Option<&str>,
    indent: Option<f64>,
    compatibility_mode: Option<u8>,
    cell_margin_left: Option<f64>,
    table_width: f64,
    content_width: f64,
) -> f64 {
    if let Some(position) = position
        && matches!(
            position.horz_anchor.as_deref(),
            None | Some("margin" | "text")
        )
    {
        match position.tblp_x_spec.as_deref() {
            Some("left") => return 0.0,
            Some("center") => return (content_width - table_width) / 2.0,
            Some("right") => return content_width - table_width,
            None => {
                if let Some(offset) = position.tblp_x.filter(|offset| offset.is_finite()) {
                    return offset;
                }
            }
            _ => {}
        }
    }
    match justification {
        Some("center") => ((content_width - table_width) / 2.0).max(0.0),
        Some("right") => (content_width - table_width).max(0.0),
        _ => {
            indent.unwrap_or(0.0).max(0.0)
                - table_compat_leading_shift(justification, compatibility_mode, cell_margin_left)
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CellContentLayout {
    /// Per block, the top y of each line (relative to `start_y`). Atomic/non-paragraph blocks → [].
    pub line_tops: Vec<Vec<f64>>,
    /// All line bottoms in document order, plus one entry per atomic block (its
    /// bottom) — the clean break points for the paginator.
    pub flat_bottoms: Vec<f64>,
    /// Total stacked height incl. the last block's trailing space-after.
    pub content_height: f64,
}

pub(crate) fn cell_vertical_offset(
    alignment: Option<&str>,
    cell_height: f64,
    measured_height: f64,
    content_height: f64,
    top_inset: f64,
    bottom_inset: f64,
) -> f64 {
    if measured_height >= cell_height - 0.5 {
        return 0.0;
    }
    let slack = (cell_height - top_inset - bottom_inset - content_height).max(0.0);
    match alignment {
        Some("center") => slack / 2.0,
        Some("bottom") => slack,
        _ => 0.0,
    }
}

/// Returns a measured block's stacked height.
fn extent_total_height(measure: &BlockExtent) -> Option<f64> {
    match measure {
        BlockExtent::Paragraph(p) => Some(p.total_height),
        BlockExtent::Table(t) => Some(t.total_height),
        BlockExtent::Image(image) => Some(image.height),
        BlockExtent::TextBox(text_box) => Some(text_box.height),
        BlockExtent::Shape(shape) => Some(shape.height),
        BlockExtent::Chart(chart) => Some(chart.height),
        _ => None,
    }
}

/// Compute the collapsed vertical layout of a cell's blocks starting at `start_y`.
pub fn layout_cell_content(
    blocks: Option<&[LayoutBlock]>,
    block_measures: Option<&[BlockExtent]>,
    start_y: f64,
) -> CellContentLayout {
    let mut line_tops: Vec<Vec<f64>> = Vec::new();
    let mut flat_bottoms: Vec<f64> = Vec::new();
    let mut y = start_y;
    let mut prev_after = 0.0f64;
    let mut float_bottom = start_y;
    let n = block_measures.map(|m| m.len()).unwrap_or(0);

    for i in 0..n {
        let measure = &block_measures.unwrap()[i];
        let block = blocks.and_then(|b| b.get(i));
        if let (Some(LayoutBlock::Paragraph(paragraph)), BlockExtent::Paragraph(para_measure)) =
            (block, measure)
        {
            let spacing = paragraph.attrs.as_ref().and_then(|a| a.spacing.as_ref());
            let before = spacing.and_then(|s| s.before).unwrap_or(0.0);
            y += prev_after.max(before);
            let mut tops: Vec<f64> = Vec::new();
            for line in &para_measure.lines {
                y += line.float_skip_before.unwrap_or(0.0);
                tops.push(y);
                y += line.line_height;
                flat_bottoms.push(y);
            }
            line_tops.push(tops);
            prev_after = spacing.and_then(|s| s.after).unwrap_or(0.0);
        } else if let (Some(LayoutBlock::Table(table)), BlockExtent::Table(extent)) =
            (block, measure)
            && let Some(offset) = nested_table_float_offset(table.floating.as_ref())
        {
            y += prev_after;
            let bottom = y + offset + extent.total_height;
            float_bottom = float_bottom.max(bottom);
            line_tops.push(Vec::new());
            flat_bottoms.push(bottom);
            prev_after = 0.0;
        } else if let Some(LayoutBlock::Shape(shape)) = block
            && cell_overlay_drawing(shape.position.is_some(), shape.wrap_type.as_deref())
        {
            line_tops.push(Vec::new());
        } else if let Some(total_height) = extent_total_height(measure) {
            // Nested table / non-paragraph: one atomic block (break only at its bottom).
            y += prev_after;
            y += total_height;
            line_tops.push(Vec::new());
            flat_bottoms.push(y);
            prev_after = 0.0;
        } else {
            line_tops.push(Vec::new());
        }
    }

    // The final block's trailing space-after becomes the cell's bottom padding.
    CellContentLayout {
        line_tops,
        flat_bottoms,
        content_height: (y + prev_after).max(float_bottom) - start_y,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    const LINE: f64 = 20.0;
    const SP: f64 = 8.0;

    #[test]
    fn nested_float_offsets_require_supported_finite_text_positions() {
        let mut position: FloatingTablePosition = serde_json::from_value(json!({
            "vertAnchor":"text", "horzAnchor":"margin", "tblpX":28, "tblpY":20
        }))
        .unwrap();
        assert_eq!(nested_table_float_offset(Some(&position)), Some(20.0));
        position.tblp_x = None;
        assert_eq!(nested_table_float_offset(Some(&position)), Some(20.0));
        for x in [Some(f64::NAN), Some(f64::INFINITY)] {
            position.tblp_x = x;
            assert_eq!(nested_table_float_offset(Some(&position)), None);
        }
        position.tblp_x = Some(28.0);
        for y in [None, Some(-1.0), Some(f64::NAN), Some(f64::INFINITY)] {
            position.tblp_y = y;
            assert_eq!(nested_table_float_offset(Some(&position)), None);
        }
        position.tblp_y = Some(20.0);
        for spec in ["left", "center", "right"] {
            position.tblp_x_spec = Some(spec.to_owned());
            assert_eq!(nested_table_float_offset(Some(&position)), Some(20.0));
        }
        for spec in ["inside", "outside", "unsupported"] {
            position.tblp_x_spec = Some(spec.to_owned());
            assert_eq!(nested_table_float_offset(Some(&position)), None);
        }
        position.tblp_x_spec = None;
        position.tblp_y_spec = Some("top".to_owned());
        assert_eq!(nested_table_float_offset(Some(&position)), None);
        position.tblp_y_spec = None;
        position.horz_anchor = Some("page".to_owned());
        assert_eq!(nested_table_float_offset(Some(&position)), None);
        position.horz_anchor = None;
        for anchor in [None, Some("page"), Some("margin")] {
            position.vert_anchor = anchor.map(str::to_owned);
            assert_eq!(nested_table_float_offset(Some(&position)), None);
        }
    }

    fn para(spacing: Option<(f64, f64)>) -> LayoutBlock {
        let attrs = spacing
            .map(|(before, after)| json!({ "spacing": { "before": before, "after": after } }));
        serde_json::from_value(json!({
            "kind": "paragraph",
            "id": 0,
            "runs": [],
            "attrs": attrs,
        }))
        .unwrap()
    }

    fn pm(lines: usize, spacing: Option<(f64, f64)>) -> BlockExtent {
        let (before, after) = spacing.unwrap_or((0.0, 0.0));
        let line = json!({
            "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 0,
            "width": 0.0, "ascent": 0.0, "descent": 0.0, "lineHeight": LINE,
        });
        serde_json::from_value(json!({
            "kind": "paragraph",
            "lines": vec![line; lines],
            "totalHeight": before + lines as f64 * LINE + after,
        }))
        .unwrap()
    }

    #[test]
    fn collapses_adjacent_paragraph_spacing_and_stacks_lines_from_each_block_top() {
        let sp = Some((SP, SP));
        let blocks = vec![para(sp), para(sp), para(sp)];
        let measures = vec![pm(1, sp), pm(1, sp), pm(1, sp)];

        let layout = layout_cell_content(Some(&blocks), Some(&measures), 0.0);

        // line tops: 8, then 8+20+max(8,8)=36, then 36+20+8=64
        assert_eq!(layout.line_tops[0], vec![SP]);
        assert_eq!(layout.line_tops[1], vec![SP + LINE + SP]);
        assert_eq!(layout.line_tops[2], vec![SP + LINE + SP + LINE + SP]);
        assert_eq!(
            layout.flat_bottoms,
            vec![
                SP + LINE,
                SP + LINE + SP + LINE,
                SP + LINE + SP + LINE + SP + LINE,
            ]
        );
        // content height includes the trailing after-spacing
        assert_eq!(
            layout.content_height,
            SP + LINE + SP + LINE + SP + LINE + SP
        );
    }

    #[test]
    fn treats_a_non_paragraph_nested_table_block_as_one_atomic_break_point() {
        let nested_table: LayoutBlock = serde_json::from_value(json!({
            "kind": "table",
            "id": 1,
            "rows": [],
        }))
        .unwrap();
        let table_measure: BlockExtent = serde_json::from_value(json!({
            "kind": "table",
            "rows": [],
            "columnWidths": [],
            "totalWidth": 0.0,
            "totalHeight": 50.0,
        }))
        .unwrap();
        let blocks = vec![para(Some((SP, SP))), nested_table];
        let measures = vec![pm(1, Some((SP, SP))), table_measure];
        let layout = layout_cell_content(Some(&blocks), Some(&measures), 0.0);
        // paragraph line bottom at 8+20=28; nested table atomic: gap = prevAfter(8) + 50
        assert_eq!(layout.line_tops[0], vec![SP]);
        assert_eq!(layout.line_tops[1], Vec::<f64>::new()); // atomic block has no per-line tops
        assert_eq!(layout.flat_bottoms, vec![SP + LINE, SP + LINE + SP + 50.0]);
        assert_eq!(layout.content_height, SP + LINE + SP + 50.0);
    }

    #[test]
    fn honors_start_y_and_multi_line_blocks() {
        let blocks = vec![para(None), para(None)];
        let measures = vec![pm(2, None), pm(1, None)];
        let layout = layout_cell_content(Some(&blocks), Some(&measures), 5.0);
        // block 0: tops 5, 25 (two lines); block 1: top 45
        assert_eq!(layout.line_tops[0], vec![5.0, 5.0 + LINE]);
        assert_eq!(layout.line_tops[1], vec![5.0 + 2.0 * LINE]);
        assert_eq!(layout.content_height, 3.0 * LINE);
    }

    #[test]
    fn includes_atomic_images_in_cell_height_and_break_boundaries() {
        let image: LayoutBlock = serde_json::from_value(json!({
            "kind": "image", "id": 1, "src": "rId1", "width": 50, "height": 40
        }))
        .unwrap();
        let image_measure: BlockExtent = serde_json::from_value(json!({
            "kind": "image", "width": 50, "height": 40
        }))
        .unwrap();
        let blocks = vec![para(Some((0.0, 5.0))), image, para(None)];
        let measures = vec![pm(1, Some((0.0, 5.0))), image_measure, pm(1, None)];
        let layout = layout_cell_content(Some(&blocks), Some(&measures), 2.0);
        assert_eq!(layout.line_tops, vec![vec![2.0], vec![], vec![67.0]]);
        assert_eq!(layout.flat_bottoms, vec![22.0, 67.0, 87.0]);
        assert_eq!(layout.content_height, 85.0);
    }

    fn twips_to_px(twips: f64) -> f64 {
        twips / 15.0
    }

    #[test]
    fn compat_mode_14_left_shifts_by_resolved_cell_margin() {
        let margin = twips_to_px(108.0);
        assert_eq!(
            table_compat_leading_shift(Some("left"), Some(14), Some(margin)),
            margin
        );
        assert_eq!(
            table_compat_leading_shift(None, Some(14), Some(margin)),
            margin
        );
        // 108 tw = 7.2 px at 96 DPI (11.25 px at 150 DPI capture scale).
        let offset = nested_table_horizontal_offset(
            None,
            Some("left"),
            Some(0.0),
            Some(14),
            Some(margin),
            200.0,
            500.0,
        );
        assert!((offset - (-margin)).abs() < 1e-9);
    }

    #[test]
    fn compat_mode_15_does_not_shift() {
        let margin = twips_to_px(108.0);
        assert_eq!(
            table_compat_leading_shift(Some("left"), Some(15), Some(margin)),
            0.0
        );
        let offset = nested_table_horizontal_offset(
            None,
            Some("left"),
            Some(0.0),
            Some(15),
            Some(margin),
            200.0,
            500.0,
        );
        assert_eq!(offset, 0.0);
    }

    #[test]
    fn centred_and_right_tables_never_shift() {
        let margin = twips_to_px(108.0);
        for justification in [Some("center"), Some("right")] {
            for mode in [Some(12), Some(14), Some(15), None] {
                assert_eq!(
                    table_compat_leading_shift(justification, mode, Some(margin)),
                    0.0
                );
            }
        }
        // Centred placement ignores the margin in either mode.
        assert_eq!(
            nested_table_horizontal_offset(
                None,
                Some("center"),
                Some(0.0),
                Some(14),
                Some(margin),
                200.0,
                500.0,
            ),
            150.0
        );
        assert_eq!(
            nested_table_horizontal_offset(
                None,
                Some("center"),
                Some(0.0),
                Some(15),
                Some(margin),
                200.0,
                500.0,
            ),
            150.0
        );
    }

    #[test]
    fn non_default_28tw_margin_shifts_by_its_own_value() {
        let small = twips_to_px(28.0);
        let large = twips_to_px(108.0);
        assert!((small - 1.8666666666666667).abs() < 1e-9);
        assert!((large - 7.2).abs() < 1e-9);
        assert_eq!(
            table_compat_leading_shift(Some("left"), Some(14), Some(small)),
            small
        );
        assert_ne!(
            table_compat_leading_shift(Some("left"), Some(14), Some(small)),
            large
        );
        let offset = nested_table_horizontal_offset(
            None,
            None,
            Some(10.0),
            Some(14),
            Some(small),
            200.0,
            500.0,
        );
        assert!((offset - (10.0 - small)).abs() < 1e-9);
    }

    #[test]
    fn absent_compat_mode_behaves_as_mode_12() {
        let margin = twips_to_px(108.0);
        assert_eq!(table_compat_leading_shift(None, None, Some(margin)), margin);
        assert_eq!(
            table_compat_leading_shift(None, Some(12), Some(margin)),
            margin
        );
        let offset =
            nested_table_horizontal_offset(None, None, Some(0.0), None, Some(margin), 200.0, 500.0);
        assert!((offset - (-margin)).abs() < 1e-9);
    }
}
