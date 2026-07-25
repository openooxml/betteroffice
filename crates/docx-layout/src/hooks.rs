//! FEATURE HOOKS — seams for the pagination features not yet ported.
//!
//! Each function here is a thin stub that mirrors a TS export being ported
//! CONCURRENTLY as a standalone module. A stub returns
//! `LayoutError::Unsupported` the moment its feature is actually engaged by
//! the input (and the neutral value otherwise), so paragraph-only documents
//! flow while uncovered features surface `Unsupported` to the caller.
//!
//! Integration contract: when the sibling module lands, replace the stub body
//! with a call into it — the call sites in `place.rs`/`prescan.rs` do not
//! change. Signatures mirror the TS exports modulo the `Result` wrapper that
//! carries the Unsupported signal (the real ports return `Ok(..)`
//! unconditionally).

use crate::LayoutError;
use crate::page_flow::Paginator;
use crate::prescan::SectionLayoutConfig;
use crate::table_row_break::{build_table_row_break_info, snap_row_break};
use crate::types::{
    Fragment, LayoutBlock, MeasuredBlock, SectionBreakBlock, SectionBreakType, TableBlock,
    TableExtent, TableFragment,
};
use crate::{break_policy, column_balancing, keep_together, section_breaks};

fn unsupported(feature: &str) -> LayoutError {
    LayoutError::Unsupported(format!("not ported yet: {feature}"))
}

// ---------------------------------------------------------------------------
// break / keep policy (breakPolicy.ts + keepTogether.ts) — WIRED
// ---------------------------------------------------------------------------

// the scan types live with their producer; re-exported so the spine's
// `prescan`/`place` keep importing them from the hooks seam
pub use crate::keep_together::{KeepWithNextGroup, KeepWithNextScan};

/// Wired seam for `breaksBeforeBlock(block: LayoutBlock): boolean` in
/// `packages/core/src/layout/pagination/breakPolicy.ts`.
pub fn breaks_before_block(block: &LayoutBlock) -> Result<bool, LayoutError> {
    Ok(break_policy::breaks_before_block(block))
}

/// Wired seam for `analyzeKeepWithNext(blocks: LayoutBlock[]): KeepWithNextScan`
/// in `packages/core/src/layout/pagination/keepTogether.ts`.
pub fn analyze_keep_with_next(measured: &[MeasuredBlock]) -> Result<KeepWithNextScan, LayoutError> {
    Ok(keep_together::analyze_keep_with_next(measured))
}

/// Wired seam for `measureKeepWithNextGroup(group, measured): number` in
/// `keepTogether.ts`.
pub fn measure_keep_with_next_group(
    group: &KeepWithNextGroup,
    measured: &[MeasuredBlock],
) -> Result<f64, LayoutError> {
    Ok(keep_together::measure_keep_with_next_group(group, measured))
}

/// Wired seam for `keepWithNextGroupMustAdvance(fit: KeepWithNextFit): boolean`
/// in `breakPolicy.ts`. The TS `KeepWithNextFit` fields arrive positionally.
pub fn keep_with_next_group_must_advance(
    group_height: f64,
    available_height: f64,
    page_content_height: f64,
    page_has_content: bool,
) -> Result<bool, LayoutError> {
    Ok(break_policy::keep_with_next_group_must_advance(
        break_policy::KeepWithNextFit {
            group_height,
            available_height,
            page_content_height,
            page_has_content,
        },
    ))
}

// ---------------------------------------------------------------------------
// section breaks + column balancing (sectionBreaks.ts + columnBalancing.ts) —
// WIRED
// ---------------------------------------------------------------------------

/// Wired seam for `handleSectionBreak(block, paginator, nextSectionConfig,
/// nextSectionType): void` in
/// `packages/core/src/layout/pagination/sectionBreaks.ts`. The paginator's
/// section-break slice is the `SectionBreakPaginator` impl in `page_flow.rs`.
pub fn handle_section_break(
    block: &SectionBreakBlock,
    paginator: &mut Paginator,
    next_section_config: &SectionLayoutConfig,
    next_section_type: Option<SectionBreakType>,
) -> Result<(), LayoutError> {
    section_breaks::handle_section_break(block, paginator, next_section_config, next_section_type)
}

/// Wired seam for `balanceTerminalContinuousTextColumns({ measured, paginator,
/// start, end }): void` in
/// `packages/core/src/layout/pagination/columnBalancing.ts`. The paginator's
/// balancing slice is the `ColumnBalancePaginator` impl in `page_flow.rs`.
pub fn balance_terminal_continuous_text_columns(
    measured: &[MeasuredBlock],
    paginator: &mut Paginator,
    start: usize,
    end: usize,
) -> Result<(), LayoutError> {
    column_balancing::balance_terminal_continuous_text_columns(measured, paginator, start, end);
    Ok(())
}

// ---------------------------------------------------------------------------
// table placement (index.ts layoutTable + tableWidthUtils.ts +
// cellBlockLayout.ts + tableRowBreak.ts)
// ---------------------------------------------------------------------------

/// `tallyHeaderRows` (index.ts) — length of the leading run of isHeader rows,
/// the band that repeats on continuation pages.
fn tally_header_rows(block: &TableBlock) -> usize {
    let mut count = 0usize;
    for row in &block.rows {
        if row.is_header.unwrap_or(false) {
            count += 1;
        } else {
            break;
        }
    }
    count
}

/// `getHeaderRowsHeight` (index.ts) — summed measured height of the leading
/// header rows, the overhead a continuation fragment pays to repeat them.
fn get_header_rows_height(measure: &TableExtent, header_row_count: usize) -> f64 {
    let mut height = 0.0f64;
    let mut i = 0usize;
    while i < header_row_count && i < measure.rows.len() {
        height += measure.rows[i].height;
        i += 1;
    }
    height
}

/// Port of `layoutTable(block, measure, paginator): void` in
/// `packages/core/src/layout/pagination/index.ts`, leaning on
/// `buildTableRowBreakInfo` / `snapRowBreak` (`tableRowBreak.ts`),
/// `layoutCellContent` (`cellBlockLayout.ts`) and the grid geometry in
/// `tableWidthUtils.ts`.
///
/// Rows are placed in order. A row that doesn't fit in the remaining space is
/// broken across the page boundary (Word's "allow row to break across pages")
/// at the deepest whole line that fits — the leftover continues on the next
/// page. The cursor into the table is `(row_index, consumed)` where `consumed`
/// is how many px of `row_index` were already placed on a previous fragment.
pub fn layout_table(
    block: &TableBlock,
    measure: &TableExtent,
    paginator: &mut Paginator,
) -> Result<(), LayoutError> {
    let rows = &measure.rows;
    if rows.is_empty() {
        return Ok(());
    }

    let header_row_count = tally_header_rows(block);
    let header_rows_height = get_header_rows_height(measure, header_row_count);
    let break_info = build_table_row_break_info(block, measure);

    let mut row_index = 0usize;
    let mut consumed = 0.0f64; // px of rows[row_index] already placed on a previous fragment

    while row_index < rows.len() {
        let state_idx = paginator.get_current();
        let is_first_fragment = row_index == 0 && consumed == 0.0;
        let column_capacity =
            paginator.state(state_idx).content_limit - paginator.state(state_idx).content_top;
        let row_remaining_at_start = rows[row_index].height - consumed;
        let row_cant_split = block
            .rows
            .get(row_index)
            .and_then(|row| row.cant_split)
            .unwrap_or(false);
        if row_cant_split
            && consumed == 0.0
            && row_remaining_at_start > paginator.get_available_height()
            && paginator.state(state_idx).pen_y != paginator.state(state_idx).content_top
        {
            paginator.ensure_fits(row_remaining_at_start);
            continue;
        }

        // Account for trailing spacing from the previous block that addFragment
        // will consume (only the first fragment butts against prior content).
        let pending_spacing = if is_first_fragment {
            paginator.state(state_idx).deferred_spacing
        } else {
            0.0
        };
        let first_safe_offset = break_info.break_offsets[row_index]
            .iter()
            .copied()
            .find(|offset| *offset > consumed);
        let minimum_body_slice = first_safe_offset
            .map(|offset| offset - consumed)
            .unwrap_or(row_remaining_at_start);
        let header_overhead = if !is_first_fragment
            && header_row_count > 0
            && header_rows_height + minimum_body_slice.max(0.0) <= column_capacity
        {
            header_rows_height
        } else {
            0.0
        };
        let available_height = paginator.get_available_height() - pending_spacing - header_overhead;

        let start_row = row_index;
        let clip_top = consumed;
        let mut used = 0.0f64;
        let mut cur = row_index;
        // Px of `cur` already placed on a previous fragment. Only the first row of
        // this fragment can carry one (the rest start at 0); `cur == row_end` holds
        // at the top of every iteration.
        let first_row_offset = consumed;
        let mut row_end = row_index; // exclusive
        let mut clip_bottom: Option<f64> = None;
        let mut last_row_partial = false;

        while cur < rows.len() {
            let row_height = rows[cur].height;
            let start_off = if cur == start_row {
                first_row_offset
            } else {
                0.0
            };
            let remaining = row_height - start_off;

            if used + remaining <= available_height {
                // The rest of this row fits whole.
                used += remaining;
                cur += 1;
                row_end = cur;
                continue;
            }

            // This row does not fully fit in the remaining space. Break it mid-content
            // at the deepest whole line that fits (Word's "allow row to break across
            // pages") — this keeps the row's other columns on the page where they
            // start and flows a tall vertically-merged cell across the boundary.
            // `w:cantSplit` rows (§17.4.6) never break.
            let budget = available_height - used;
            let cant_split = block
                .rows
                .get(cur)
                .and_then(|r| r.cant_split)
                .unwrap_or(false);
            let unavoidable_cant_split =
                cant_split && remaining > column_capacity - header_overhead;
            let placeable = if cant_split && !unavoidable_cant_split {
                0.0
            } else {
                snap_row_break(&break_info, cur, start_off, budget)
            };
            if placeable > 0.0 {
                // Break this row mid-content at a whole-line boundary.
                used += placeable;
                row_end = cur + 1;
                clip_bottom = Some(start_off + placeable);
                last_row_partial = true;
            } else if row_end > start_row {
                // Nothing of this row fits, but earlier rows did — end before it.
            } else {
                // Fresh fragment and not even one line fits: place the rest of the row
                // with overflow rather than loop forever (oversized-row guard).
                used += remaining;
                row_end = cur + 1;
            }
            break;
        }

        // `used` is the visible content window; repeated headers stack on top of it
        let fragment_height = header_overhead + used;
        let is_last_fragment = row_end == rows.len() && !last_row_partial;

        let mut desired_x = paginator.get_column_x(paginator.state(state_idx).column_index);
        if block.justification.as_deref() == Some("center") {
            desired_x += (paginator.column_width() - measure.total_width) / 2.0;
        } else if block.justification.as_deref() == Some("right") {
            desired_x += paginator.column_width() - measure.total_width;
        } else if let Some(indent) = block.indent
            && indent != 0.0
            && !indent.is_nan()
        {
            // TS `else if (block.indent)` — truthiness skips 0 / NaN
            desired_x += indent;
        }

        let fragment = Fragment::Table(TableFragment {
            block_id: block.id.clone(),
            x: desired_x,
            y: 0.0, // set by add_fragment
            width: measure.total_width,
            height: fragment_height,
            row_start: start_row,
            row_end,
            pm_start: block.pm_start,
            pm_end: block.pm_end,
            is_floating: None,
            carried_from_prev: Some(!is_first_fragment),
            carried_to_next: Some(!is_last_fragment),
            header_row_count: if !is_first_fragment && header_row_count > 0 {
                Some(header_row_count as f64)
            } else {
                None
            },
            clip_top: if clip_top > 0.0 { Some(clip_top) } else { None },
            clip_bottom,
        });

        paginator.add_fragment(fragment, fragment_height, 0.0, 0.0);

        // The TS re-asserts the justified/indented x after addFragment (which
        // wrote the plain column x); mirror by patching the fragment just
        // pushed onto the page addFragment landed on.
        let landed_idx = paginator.get_current();
        let page_index = paginator.state(landed_idx).page_index;
        if let Some(Fragment::Table(placed)) = paginator.pages[page_index].fragments.last_mut() {
            placed.x = desired_x;
        }

        // Advance the cursor. A partial last row resumes at its break point
        // (`clip_bottom`); otherwise we move past the rows just placed.
        if last_row_partial {
            row_index = row_end - 1;
            consumed = clip_bottom.unwrap_or(0.0);
        } else {
            row_index = row_end;
            consumed = 0.0;
        }

        // If content remains, advance to the next column/page so the next
        // iteration sees fresh space (the current page is exhausted).
        if row_index < rows.len() {
            let next_offset = break_info.break_offsets[row_index]
                .iter()
                .copied()
                .find(|offset| *offset > consumed);
            let next_slice = next_offset
                .map(|offset| offset - consumed)
                .unwrap_or(rows[row_index].height - consumed);
            let next_needed =
                if header_row_count > 0 && header_rows_height + next_slice <= column_capacity {
                    header_rows_height
                } else {
                    0.0
                } + next_slice;
            paginator.ensure_fits(next_needed);
        }
    }

    Ok(())
}

// ---------------------------------------------------------------------------
// floating objects (floatingObjects.ts)
// ---------------------------------------------------------------------------

/// HOOK for `layoutFloatingTable(block, measure, paginator, contentWidth):
/// void` in `packages/core/src/layout/pagination/index.ts`, which reads
/// `MIN_WRAP_SEGMENT_WIDTH` from
/// `packages/core/src/layout/pagination/floatingObjects.ts`.
pub fn layout_floating_table(
    block: &TableBlock,
    measure: &TableExtent,
    paginator: &mut Paginator,
    _content_width: f64,
) -> Result<(), LayoutError> {
    if block.rows.is_empty() || measure.rows.is_empty() {
        return Err(unsupported("floating table without measurable rows"));
    }
    let initial_state = paginator.get_current();
    let column_capacity =
        paginator.state(initial_state).content_limit - paginator.state(initial_state).content_top;
    if measure.total_height > column_capacity {
        return layout_table(block, measure, paginator);
    }

    let floating = block
        .floating
        .as_ref()
        .expect("floating table hook requires tblpPr");
    let has_explicit_y = floating.tblp_y.is_some()
        || floating
            .tblp_y_spec
            .as_deref()
            .is_some_and(|spec| spec != "inline");
    if !has_explicit_y {
        paginator.ensure_fits(measure.total_height);
    }

    let state_idx = paginator.get_current();
    let state = paginator.state(state_idx);
    let page = &paginator.pages[state.page_index];
    let column_x = paginator.get_column_x(state.column_index);
    let column_width = paginator.column_width();
    let horizontal = floating.horz_anchor.as_deref().unwrap_or("margin");
    let vertical = floating.vert_anchor.as_deref().unwrap_or("text");
    let h_start = match horizontal {
        "page" => 0.0,
        "text" => column_x,
        _ => page.margins.left,
    };
    let h_end = match horizontal {
        "page" => page.size.w,
        "text" => column_x + column_width,
        _ => page.size.w - page.margins.right,
    };
    let v_start = match vertical {
        "page" => 0.0,
        "text" => state.pen_y,
        _ => page.margins.top,
    };
    let v_end = if vertical == "page" {
        page.size.h
    } else {
        page.size.h - page.margins.bottom
    };
    let inside_is_start = page.number % 2 == 1;

    let mut x = h_start;
    if let Some(offset) = floating.tblp_x.filter(|value| value.is_finite()) {
        x = h_start + offset;
    } else if let Some(spec) = floating.tblp_x_spec.as_deref() {
        let start_aligned = spec == "left"
            || (spec == "inside" && inside_is_start)
            || (spec == "outside" && !inside_is_start);
        let end_aligned = spec == "right"
            || (spec == "inside" && !inside_is_start)
            || (spec == "outside" && inside_is_start);
        if spec == "center" {
            x = h_start + (h_end - h_start - measure.total_width) / 2.0;
        } else if end_aligned {
            x = h_end - measure.total_width;
        } else if start_aligned {
            x = h_start;
        }
    } else if block.justification.as_deref() == Some("center") {
        x = h_start + (h_end - h_start - measure.total_width) / 2.0;
    } else if block.justification.as_deref() == Some("right") {
        x = h_end - measure.total_width;
    }

    let mut y = state.pen_y;
    if let Some(offset) = floating.tblp_y.filter(|value| value.is_finite()) {
        y = v_start + offset;
    } else if let Some(spec) = floating.tblp_y_spec.as_deref()
        && spec != "inline"
    {
        y = match spec {
            "center" => v_start + (v_end - v_start - measure.total_height) / 2.0,
            "bottom" | "outside" => v_end - measure.total_height,
            _ => v_start,
        };
    }

    let fragment = Fragment::Table(TableFragment {
        block_id: block.id.clone(),
        x,
        y,
        width: measure.total_width,
        height: measure.total_height,
        row_start: 0,
        row_end: block.rows.len(),
        pm_start: block.pm_start,
        pm_end: block.pm_end,
        is_floating: Some(true),
        carried_from_prev: None,
        carried_to_next: None,
        header_row_count: None,
        clip_top: None,
        clip_bottom: None,
    });
    paginator.push_fragment_direct(fragment);

    let finite = |value: Option<f64>| value.filter(|v| v.is_finite()).unwrap_or(0.0);
    let exclusion_left = x - finite(floating.left_from_text);
    let exclusion_right = x + measure.total_width + finite(floating.right_from_text);
    let left_space = exclusion_left - column_x;
    let right_space = column_x + column_width - exclusion_right;
    if left_space < 24.0 && right_space < 24.0 {
        let advance_to = y + measure.total_height + finite(floating.bottom_from_text);
        if advance_to > paginator.state(state_idx).pen_y {
            paginator.set_pen_y(state_idx, advance_to);
        }
    }
    Ok(())
}
