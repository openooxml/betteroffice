//! Port of `packages/core/src/layout/pagination/cellBlockLayout.ts`.
//!
//! Vertical layout of a table cell's stacked blocks — single source of truth
//! for where the lines of a cell's content sit. Adjacent paragraphs'
//! before/after spacing collapses to the larger of the two; a paragraph's
//! lines stack from its content top with no before/after between them.
//!
//! Exported fns (1:1 with the TS module):
//! - `layout_cell_content` ← `layoutCellContent(blocks, blockMeasures, startY)`
//!
//! Consumes the spine's types (`types.rs`) directly, same as the TS module
//! consumes `pagination/types.ts`.

use serde::Serialize;

use crate::types::{BlockExtent, LayoutBlock};

/// Mirror of the TS `CellContentLayout` result.
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

/// TS `'totalHeight' in measure && typeof measure.totalHeight === 'number'` —
/// of the extent union only paragraph and table extents carry a numeric
/// `totalHeight`.
fn extent_total_height(measure: &BlockExtent) -> Option<f64> {
    match measure {
        BlockExtent::Paragraph(p) => Some(p.total_height),
        BlockExtent::Table(t) => Some(t.total_height),
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
        } else if let Some(total_height) = extent_total_height(measure) {
            // Nested table / non-paragraph: one atomic block (break only at its bottom).
            y += prev_after + total_height;
            line_tops.push(Vec::new());
            flat_bottoms.push(y);
            prev_after = 0.0;
        } else {
            line_tops.push(Vec::new());
        }
    }

    // The painter renders the final block's trailing space-after as paddingBottom.
    CellContentLayout {
        line_tops,
        flat_bottoms,
        content_height: y - start_y + prev_after,
    }
}

// Ported from packages/core/src/layout/pagination/__tests__/cellBlockLayout.test.ts
// — every case preserved. Fixtures are built through serde like real inputs.
#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    const LINE: f64 = 20.0;
    const SP: f64 = 8.0;

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
        // content height includes the trailing after (painter paddingBottom)
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
}
