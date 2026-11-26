//! rewriting stored formulas on row/column insert/delete: refs shift, ranges
//! clip, wholly deleted refs collapse to `#REF!`. runs before cells shift.

use std::cell::OnceCell;
use std::collections::{HashMap, HashSet, VecDeque};
use std::rc::Rc;

use xlsx_calc::lexer::{
    MAX_FORMULA_BYTES, ensure_unambiguous_spaced_calls, ensure_unambiguous_spaced_references,
    lex_with_range_names,
};
use xlsx_calc::parser::{Expr, parse_formula_for_rewrite};
use xlsx_calc::{ColumnRange, ParseError, RowRange, TokKind, Token, lex_for_rewrite};
use xlsx_model::addr::{MAX_COLS, MAX_ROWS, col_to_letters};
use xlsx_model::sheet_qualifier::{
    Qualifier, is_function_call, narrowed_span, next_char_len, parse_sheet_token, qualified_target,
    qualifier_token, quoted_endpoints, rename_sheets, sheet_names_equal,
    sheet_qualifiers_with_sheets, skip_string,
};
use xlsx_model::workbook::{MAX_DEFINED_NAME_BYTES, defined_names_bytes};
use xlsx_model::{
    AnchorCell, AnchorEditAs, CellRange, CellRef, ChartAnchor, ConditionalFormatRule,
    ConditionalFormatting, DefinedName, ErrorValue, SheetId, Workbook,
};

use crate::apply::OpError;
use crate::apply::move_axis;
use crate::apply::remap_ref;
use crate::op::{CellState, Op};

/// distinct parsed formulas kept in the memo before oldest-entry eviction.
const PARSE_MEMO_CAP: usize = 1024;

/// the outcome of remapping one reference or range under a structural op.
enum Remapped<T> {
    /// the op does not move this reference.
    Unchanged,
    /// the reference shifted (or a range clipped) to a new address.
    Moved(T),
    /// the reference (or the whole range) fell inside the deleted span.
    Deleted,
}

/// the sheet a structural op targets, or `None` for non-structural ops.
fn structural_target(op: &Op) -> Option<SheetId> {
    match *op {
        Op::InsertRows { sheet, .. }
        | Op::DeleteRows { sheet, .. }
        | Op::InsertCols { sheet, .. }
        | Op::DeleteCols { sheet, .. }
        | Op::MoveRows { sheet, .. }
        | Op::MoveCols { sheet, .. } => Some(sheet),
        _ => None,
    }
}

/// rewrite every workbook formula affected by `op`, in place, returning the
/// inverse `SetCell` ops that restore the rewritten formulas.
pub(crate) fn remap_formulas(wb: &mut Workbook, op: &Op) -> Result<Vec<Op>, OpError> {
    let Some(target) = structural_target(op) else {
        return Ok(Vec::new());
    };
    let names: HashMap<String, SheetId> = wb
        .sheets
        .iter()
        .enumerate()
        .map(|(i, s)| (s.name.to_lowercase(), SheetId(i as u32)))
        .collect();

    let target_name = wb
        .sheet(target)
        .map(|sheet| sheet.name.clone())
        .unwrap_or_default();
    let index = SheetIndex {
        names: &names,
        target,
        target_name: &target_name,
    };
    let mut restores: Vec<Op> = Vec::new();
    let mut edits: Vec<(SheetId, CellRef, String)> = Vec::new();
    let mut parsed: HashMap<&str, Rc<Expr>> = HashMap::new();
    let mut parsed_order: VecDeque<&str> = VecDeque::new();
    for (i, sheet) in wb.sheets.iter().enumerate() {
        let owner = SheetId(i as u32);
        let matches = |ref_sheet: &Option<String>| resolves_to_target(ref_sheet, owner, &index);
        for rule in sheet
            .conditional_formats
            .iter()
            .flat_map(|group| &group.rules)
        {
            if rule.formulas.iter().any(|source| {
                source.contains('!')
                    && sheet_qualifiers_with_sheets(source, |name| index.has_sheet(name))
                        .iter()
                        .any(|qualifier| {
                            qualifier.has_ambiguous_target(source)
                                && index.covers(&qualifier.first, &qualifier.last)
                        })
            }) {
                return Err(unrewritable_rule(rule, &sheet.name));
            }
        }
        for (cell, c) in sheet.iter_cells() {
            let Some(src) = &c.formula else {
                continue;
            };
            if has_unrewritable_binding(src, &index) {
                return Err(OpError::FormulaNotRewritable { sheet: owner, cell });
            }
            let expr = if let Some(expr) = parsed.get(src.as_str()) {
                Rc::clone(expr)
            } else {
                let expr = Rc::new(
                    parse_formula_for_rewrite(src.as_str())
                        .map_err(|_| OpError::FormulaNotRewritable { sheet: owner, cell })?,
                );
                if parsed.len() >= PARSE_MEMO_CAP
                    && let Some(oldest) = parsed_order.pop_front()
                {
                    parsed.remove(oldest);
                }
                parsed.insert(src.as_str(), Rc::clone(&expr));
                parsed_order.push_back(src.as_str());
                expr
            };
            if contains_table_reference(&expr) {
                return Err(OpError::FormulaNotRewritable { sheet: owner, cell });
            }
            let mut changed = false;
            let new_expr = transform(&expr, op, &matches, &mut changed);
            if changed {
                let formula = new_expr.to_formula();
                parse_formula_for_rewrite(&formula)
                    .map_err(|_| OpError::FormulaNotRewritable { sheet: owner, cell })?;
                restores.push(Op::SetCell {
                    sheet: owner,
                    at: cell,
                    cell: CellState::from(c),
                });
                edits.push((owner, cell, formula));
            }
        }
    }

    for (sheet, at, text) in edits {
        let s = wb.sheet_mut(sheet).expect("sheet exists during remap");
        if let Some(mut cell) = s.cell(at).cloned() {
            cell.formula = Some(text);
            s.set_cell(at, cell);
        }
    }
    Ok(restores)
}

#[doc(hidden)]
pub fn remap_proposal_formula(
    source: &str,
    owner: SheetId,
    sheet_names: &[String],
    op: &Op,
) -> Option<String> {
    ProposalSheets::new(sheet_names).remap(source, owner, op)
}

/// the sheet names many formulas remap against, indexed at most once.
#[doc(hidden)]
pub struct ProposalSheets<'a> {
    sheet_names: &'a [String],
    names: OnceCell<HashMap<String, SheetId>>,
}

impl<'a> ProposalSheets<'a> {
    pub fn new(sheet_names: &'a [String]) -> Self {
        Self {
            sheet_names,
            names: OnceCell::new(),
        }
    }

    /// [`remap_proposal_formula`] over these names.
    pub fn remap(&self, source: &str, owner: SheetId, op: &Op) -> Option<String> {
        let sheet_names = self.sheet_names;
        if let Op::RenameSheet { sheet, name } | Op::RestoreSheet { sheet, name, .. } = op {
            let renamed =
                scan_sheet_rename(source, sheet_names.get(sheet.0 as usize)?, name, |name| {
                    sheet_names
                        .iter()
                        .any(|sheet| sheet_names_equal(sheet, name))
                });
            if renamed.ambiguous || renamed.formula.len() > MAX_FORMULA_BYTES {
                return None;
            }
            return writable_rename(source, name, renamed.formula).ok();
        }
        let Some(target) = structural_target(op) else {
            return Some(source.to_string());
        };
        let names = self.names.get_or_init(|| {
            sheet_names
                .iter()
                .enumerate()
                .map(|(index, name)| (name.to_lowercase(), SheetId(index as u32)))
                .collect()
        });
        let index = SheetIndex {
            names,
            target,
            target_name: sheet_names.get(target.0 as usize)?,
        };
        if has_unrewritable_binding(source, &index) {
            return None;
        }
        let expr = parse_formula_for_rewrite(source).ok()?;
        if contains_table_reference(&expr) {
            return None;
        }
        let matches = |ref_sheet: &Option<String>| resolves_to_target(ref_sheet, owner, &index);
        let mut changed = false;
        let rewritten = transform(&expr, op, &matches, &mut changed);
        if !changed {
            return Some(source.to_string());
        }
        let formula = rewritten.to_formula();
        parse_formula_for_rewrite(&formula).ok()?;
        Some(formula)
    }
}

pub(crate) fn remap_hyperlink_locations(wb: &mut Workbook, op: &Op) -> Vec<Op> {
    let Some(target) = structural_target(op) else {
        return Vec::new();
    };
    let names: HashMap<String, SheetId> = wb
        .sheets
        .iter()
        .enumerate()
        .map(|(index, sheet)| (sheet.name.to_lowercase(), SheetId(index as u32)))
        .collect();
    let target_name = wb
        .sheet(target)
        .map(|sheet| sheet.name.clone())
        .unwrap_or_default();
    let sheets = SheetIndex {
        names: &names,
        target,
        target_name: &target_name,
    };
    let mut restores = Vec::new();
    let mut edits = Vec::new();
    for (index, sheet) in wb.sheets.iter().enumerate() {
        let owner = SheetId(index as u32);
        let matches = |ref_sheet: &Option<String>| resolves_to_target(ref_sheet, owner, &sheets);
        let mut hyperlinks = sheet.hyperlinks.clone();
        let mut changed_sheet = false;
        for hyperlink in &mut hyperlinks {
            let Some(location) = &hyperlink.location else {
                continue;
            };
            let prefixed = location.starts_with('#');
            let source = location.strip_prefix('#').unwrap_or(location);
            let Some((expr, suffix)) = split_location_reference(source) else {
                continue;
            };
            let mut changed = false;
            let rewritten = transform(&expr, op, &matches, &mut changed).to_formula();
            if changed {
                let rewritten = format!("{rewritten}{suffix}");
                hyperlink.location = Some(if prefixed && !rewritten.starts_with('#') {
                    format!("#{rewritten}")
                } else {
                    rewritten
                });
                changed_sheet = true;
            }
        }
        if changed_sheet {
            restores.push(Op::SetHyperlinks {
                sheet: owner,
                hyperlinks: sheet.hyperlinks.clone(),
            });
            edits.push((owner, hyperlinks));
        }
    }
    for (sheet, hyperlinks) in edits {
        wb.sheet_mut(sheet)
            .expect("sheet exists during hyperlink remap")
            .hyperlinks = hyperlinks;
    }
    restores
}

/// Rewrites defined names through a structural edit.
pub(crate) fn remap_defined_names(wb: &mut Workbook, op: &Op) -> Result<Option<Op>, OpError> {
    let Some(target) = structural_target(op) else {
        return Ok(None);
    };
    if wb.defined_names.is_empty() {
        return Ok(None);
    }
    let names: HashMap<String, SheetId> = wb
        .sheets
        .iter()
        .enumerate()
        .map(|(index, sheet)| (sheet.name.to_lowercase(), SheetId(index as u32)))
        .collect();
    let target_name = wb
        .sheet(target)
        .map(|sheet| sheet.name.clone())
        .unwrap_or_default();
    let sheets = SheetIndex {
        names: &names,
        target,
        target_name: &target_name,
    };
    let previous = wb.defined_names.clone();
    let defined_names = defined_name_set(&previous);
    let mut rewritten = Vec::with_capacity(previous.len());
    for defined in &previous {
        let mut updated = defined.clone();
        let scoped = defined.local_sheet == Some(target);
        let global_target =
            defined.local_sheet.is_none() && wb.sheets.len() == 1 && target == SheetId(0);
        let matches = |ref_sheet: &Option<String>| match ref_sheet {
            Some(name) => {
                let (first, last) = quoted_endpoints(name);
                sheets.names.contains_key(&first.to_lowercase())
                    && sheets.names.contains_key(&last.to_lowercase())
                    && sheets.covers(&first, &last)
            }
            None => scoped || global_target,
        };
        if has_unrewritable_binding(&defined.formula, &sheets) {
            return Err(OpError::DefinedNameNotRewritable {
                name: defined.name.clone(),
            });
        }
        let rewrite = rewrite_defined_name(
            &defined.formula,
            op,
            &matches,
            defined.local_sheet.is_none() && !global_target,
            &defined_names,
            &sheets,
        );
        match rewrite {
            DefinedNameRewrite::Unchanged => {}
            DefinedNameRewrite::Rewritten(formula) if formula.len() <= MAX_FORMULA_BYTES => {
                updated.formula = formula;
            }
            DefinedNameRewrite::Rewritten(_) | DefinedNameRewrite::AffectedUnsupported => {
                return Err(OpError::DefinedNameNotRewritable {
                    name: defined.name.clone(),
                });
            }
            DefinedNameRewrite::Ambiguous => {
                return Err(OpError::DefinedNameNotRewritable {
                    name: defined.name.clone(),
                });
            }
            DefinedNameRewrite::Unsupported
                if scoped || mentions_sheet(&defined.formula, &sheets) =>
            {
                return Err(OpError::DefinedNameNotRewritable {
                    name: defined.name.clone(),
                });
            }
            DefinedNameRewrite::Unsupported => {}
        }
        rewritten.push(updated);
    }
    if rewritten == previous {
        return Ok(None);
    }
    if defined_names_bytes(&rewritten) > MAX_DEFINED_NAME_BYTES {
        return Err(OpError::TooManyDefinedNames);
    }
    wb.defined_names = rewritten;
    Ok(Some(Op::SetDefinedNames {
        defined_names: previous,
    }))
}

/// Rewrites chart references through a row or column op, the same way defined
/// names are, and moves the anchors of charts on the edited sheet. A chart's
/// unqualified references resolve against the sheet it is anchored on. A
/// reference aimed at the edited sheet that the rewriter cannot express is
/// refused rather than left pointing at the pre-edit addresses.
pub(crate) fn remap_charts(wb: &mut Workbook, op: &Op) -> Result<Vec<Op>, OpError> {
    let Some(target) = structural_target(op) else {
        return Ok(Vec::new());
    };
    let names: HashMap<String, SheetId> = wb
        .sheets
        .iter()
        .enumerate()
        .map(|(index, sheet)| (sheet.name.to_lowercase(), SheetId(index as u32)))
        .collect();
    let target_name = wb
        .sheet(target)
        .map(|sheet| sheet.name.clone())
        .unwrap_or_default();

    let sheets = SheetIndex {
        names: &names,
        target,
        target_name: &target_name,
    };
    let defined_names = defined_name_set(&wb.defined_names);
    let mut restores = Vec::new();
    let mut edits = Vec::new();
    for (index, sheet) in wb.sheets.iter().enumerate() {
        if sheet.charts.is_empty() {
            continue;
        }
        let owner = SheetId(index as u32);
        let anchored_on_target = owner == target;
        let matches = |ref_sheet: &Option<String>| resolves_to_target(ref_sheet, owner, &sheets);
        let mut charts = sheet.charts.clone();
        let mut changed_sheet = false;
        for chart in &mut charts {
            for reference in &mut chart.refs {
                if reference.formula.trim().is_empty() {
                    continue;
                }
                if has_unrewritable_binding(&reference.formula, &sheets) {
                    return Err(OpError::ChartRefNotRewritable {
                        part: chart.part.clone(),
                    });
                }
                match rewrite_chart_ref(&reference.formula, op, &matches, &defined_names, &sheets) {
                    DefinedNameRewrite::Unchanged => {}
                    DefinedNameRewrite::Rewritten(formula)
                        if formula.len() <= MAX_FORMULA_BYTES =>
                    {
                        reference.formula = formula;
                        changed_sheet = true;
                    }
                    _ if binds_to_target(&reference.formula, &sheets, anchored_on_target) => {
                        return Err(OpError::ChartRefNotRewritable {
                            part: chart.part.clone(),
                        });
                    }
                    _ => {}
                }
            }
            if anchored_on_target {
                let moved = remap_anchor(chart.anchor, op).map_err(|()| {
                    OpError::ChartAnchorNotMovable {
                        part: chart.part.clone(),
                    }
                })?;
                if let Some(anchor) = moved {
                    chart.anchor = anchor;
                    changed_sheet = true;
                }
            }
        }
        if changed_sheet {
            restores.push(Op::SetCharts {
                sheet: owner,
                charts: sheet.charts.clone(),
            });
            edits.push((owner, charts));
        }
    }
    for (sheet, charts) in edits {
        wb.sheet_mut(sheet)
            .expect("sheet exists during chart remap")
            .charts = charts;
    }
    Ok(restores)
}

/// Whether an unrewritable chart reference points at the edited sheet: either
/// it names it, or it carries no qualifier at all and the chart sits there.
fn binds_to_target(source: &str, index: &SheetIndex<'_>, anchored_on_target: bool) -> bool {
    mentions_sheet(source, index) || (anchored_on_target && !source.contains('!'))
}

/// A chart `c:f` is a defined-name formula that may additionally be wrapped in
/// the parentheses Excel writes around a multi-area reference.
fn rewrite_chart_ref(
    source: &str,
    op: &Op,
    matches_target: &dyn Fn(&Option<String>) -> bool,
    names: &DefinedNameSet,
    sheets: &SheetIndex<'_>,
) -> DefinedNameRewrite {
    let trimmed = source.trim();
    let Some(inner) = paren_group(trimmed) else {
        return rewrite_defined_name(trimmed, op, matches_target, false, names, sheets);
    };
    match rewrite_defined_name(inner, op, matches_target, false, names, sheets) {
        DefinedNameRewrite::Rewritten(formula) => {
            DefinedNameRewrite::Rewritten(format!("({formula})"))
        }
        other => other,
    }
}

/// The inside of `(...)` when the whole reference is one parenthesized group,
/// rather than an expression that merely starts and ends with a bracket.
fn paren_group(source: &str) -> Option<&str> {
    let inner = source.strip_prefix('(')?.strip_suffix(')')?;
    let mut depth = 0_u32;
    for byte in inner.bytes() {
        match byte {
            b'(' => depth += 1,
            b')' => depth = depth.checked_sub(1)?,
            _ => {}
        }
    }
    (depth == 0).then_some(inner)
}

/// Whether every marker `op` must move stays on the grid. An insertion that
/// would push one off is refused in preflight rather than clamped, which would
/// silently resize an object whose `editAs` forbids resizing.
pub fn insertion_keeps_chart_anchor_on_grid(anchor: ChartAnchor, op: &Op) -> bool {
    remap_anchor(anchor, op).is_ok()
}

/// Moves a drawing anchor through a grid edit on its own sheet, honouring
/// `editAs`: a two-cell anchor moves and resizes with the grid, a one-cell
/// anchor moves without resizing, an absolute anchor does neither. `Ok(None)`
/// when the anchor does not move, `Err` when a marker it must move would leave
/// the grid.
fn remap_anchor(anchor: ChartAnchor, op: &Op) -> Result<Option<ChartAnchor>, ()> {
    #[derive(Clone, Copy)]
    enum Mode {
        Insert,
        Delete,
        Move(u32),
    }
    let (axis, at, count, mode) = match *op {
        Op::InsertRows { at, count, .. } => (Axis::Row, at, count, Mode::Insert),
        Op::DeleteRows { at, count, .. } => (Axis::Row, at, count, Mode::Delete),
        Op::InsertCols { at, count, .. } => (Axis::Col, at, count, Mode::Insert),
        Op::DeleteCols { at, count, .. } => (Axis::Col, at, count, Mode::Delete),
        Op::MoveRows {
            row, count, before, ..
        } => (Axis::Row, row, count, Mode::Move(before)),
        Op::MoveCols {
            col, count, before, ..
        } => (Axis::Col, col, count, Mode::Move(before)),
        _ => return Ok(None),
    };
    let limit = axis.limit();
    let shift = |index: u32| match mode {
        Mode::Insert => shift_anchor_index(index, at, count, true, limit),
        Mode::Delete => shift_anchor_index(index, at, count, false, limit),
        Mode::Move(before) => Some(move_axis(index, at, count, before)),
    };
    let moving = matches!(mode, Mode::Move(_));
    match anchor {
        ChartAnchor::Absolute { .. } => Ok(None),
        ChartAnchor::OneCell { from, extent } => {
            let moved = shift_anchor_cell(from, axis, &shift).ok_or(())?;
            Ok((moved != from).then_some(ChartAnchor::OneCell {
                from: moved,
                extent,
            }))
        }
        ChartAnchor::TwoCell { from, to, edit_as } => {
            if edit_as == AnchorEditAs::Absolute {
                return Ok(None);
            }
            let mut moved_from = shift_anchor_cell(from, axis, &shift).ok_or(())?;
            let moved_to = match edit_as {
                AnchorEditAs::TwoCell => {
                    let mut moved = shift_anchor_cell(to, axis, &shift).ok_or(())?;
                    if moving && anchor_index(moved_from, axis) > anchor_index(moved, axis) {
                        match axis {
                            Axis::Row => {
                                std::mem::swap(&mut moved_from.row, &mut moved.row);
                                std::mem::swap(&mut moved_from.row_off, &mut moved.row_off);
                            }
                            Axis::Col => {
                                std::mem::swap(&mut moved_from.col, &mut moved.col);
                                std::mem::swap(&mut moved_from.col_off, &mut moved.col_off);
                            }
                        }
                    }
                    moved
                }
                _ => translate_anchor_cell(
                    to,
                    axis,
                    i64::from(anchor_index(moved_from, axis)) - i64::from(anchor_index(from, axis)),
                    limit,
                )
                .ok_or(())?,
            };
            Ok(
                (moved_from != from || moved_to != to).then_some(ChartAnchor::TwoCell {
                    from: moved_from,
                    to: moved_to,
                    edit_as,
                }),
            )
        }
    }
}

/// An index the edit pushes along, or pulls back onto the deletion point when
/// the row or column it named is gone. `None` when an insertion would push it
/// past the last row or column.
fn shift_anchor_index(index: u32, at: u32, count: u32, inserting: bool, limit: u32) -> Option<u32> {
    let ceiling = limit.saturating_sub(1);
    if inserting {
        if index >= at {
            index.checked_add(count).filter(|moved| *moved <= ceiling)
        } else {
            Some(index)
        }
    } else if index >= at.saturating_add(count) {
        Some(index - count)
    } else if index >= at {
        Some(at)
    } else {
        Some(index)
    }
}

fn anchor_index(cell: AnchorCell, axis: Axis) -> u32 {
    match axis {
        Axis::Row => cell.row,
        Axis::Col => cell.col,
    }
}

fn shift_anchor_cell(
    cell: AnchorCell,
    axis: Axis,
    shift: &dyn Fn(u32) -> Option<u32>,
) -> Option<AnchorCell> {
    let mut moved = cell;
    match axis {
        Axis::Row => moved.row = shift(cell.row)?,
        Axis::Col => moved.col = shift(cell.col)?,
    }
    Some(moved)
}

fn translate_anchor_cell(
    cell: AnchorCell,
    axis: Axis,
    delta: i64,
    limit: u32,
) -> Option<AnchorCell> {
    let ceiling = i64::from(limit.saturating_sub(1));
    let moved = i64::from(anchor_index(cell, axis)) + delta;
    if !(0..=ceiling).contains(&moved) {
        return None;
    }
    let mut out = cell;
    match axis {
        Axis::Row => out.row = moved as u32,
        Axis::Col => out.col = moved as u32,
    }
    Some(out)
}

/// Rewrites the sheet qualifier in every chart reference on a rename. A chart
/// reference that names the old sheet as a bare token carries no `!` to bind
/// it, so it is left alone rather than guessed at.
pub(crate) fn rename_chart_refs(wb: &mut Workbook, old_name: &str, new_name: &str) -> Vec<Op> {
    if old_name == new_name {
        return Vec::new();
    }
    let mut restores = Vec::new();
    let mut edits = Vec::new();
    for (index, sheet) in wb.sheets.iter().enumerate() {
        if sheet.charts.is_empty() {
            continue;
        }
        let mut charts = sheet.charts.clone();
        for chart in &mut charts {
            for reference in &mut chart.refs {
                reference.formula =
                    scan_sheet_rename(&reference.formula, old_name, new_name, |name| {
                        workbook_has_sheet(wb, name)
                    })
                    .formula;
            }
        }
        if charts != sheet.charts {
            restores.push(Op::SetCharts {
                sheet: SheetId(index as u32),
                charts: sheet.charts.clone(),
            });
            edits.push((SheetId(index as u32), charts));
        }
    }
    for (sheet, charts) in edits {
        wb.sheet_mut(sheet)
            .expect("sheet exists during chart rename")
            .charts = charts;
    }
    restores
}

/// Collapses the chart references that name a sheet the workbook no longer
/// has. A 3-D reference whose endpoint was removed shrinks onto the sheets
/// that remain instead; one that loses every sheet it covered collapses. The
/// rest of a multi-area reference survives either way.
pub(crate) fn strand_chart_refs(
    wb: &mut Workbook,
    removed_name: &str,
    order_before: &[String],
) -> Vec<Op> {
    let mut restores = Vec::new();
    let mut edits = Vec::new();
    for (index, sheet) in wb.sheets.iter().enumerate() {
        if sheet.charts.is_empty() {
            continue;
        }
        let mut charts = sheet.charts.clone();
        for chart in &mut charts {
            for reference in &mut chart.refs {
                if let Some(dropped) =
                    drop_removed_sheet(&reference.formula, removed_name, order_before)
                {
                    reference.formula = dropped;
                }
            }
        }
        if charts != sheet.charts {
            restores.push(Op::SetCharts {
                sheet: SheetId(index as u32),
                charts: sheet.charts.clone(),
            });
            edits.push((SheetId(index as u32), charts));
        }
    }
    for (sheet, charts) in edits {
        wb.sheet_mut(sheet)
            .expect("sheet exists during chart sheet removal")
            .charts = charts;
    }
    restores
}

fn drop_removed_sheet(source: &str, removed_name: &str, order_before: &[String]) -> Option<String> {
    let trimmed = source.trim();
    let (inner, wrapped) = match paren_group(trimmed) {
        Some(inner) => (inner, true),
        None => (trimmed, false),
    };
    let components = split_union(inner)?;
    let mut changed = false;
    let mut rewritten = Vec::with_capacity(components.len());
    for component in &components {
        match rewrite_component_for_removal(component, removed_name, order_before) {
            ComponentRemoval::Unchanged => rewritten.push((*component).to_owned()),
            ComponentRemoval::Rewritten(text) => {
                changed = true;
                rewritten.push(text);
            }
            ComponentRemoval::Collapsed => {
                changed = true;
                rewritten.push(ErrorValue::Ref.as_str().to_owned());
            }
        }
    }
    if !changed {
        return None;
    }
    let rewritten = rewritten.join(",");
    Some(if wrapped {
        format!("({rewritten})")
    } else {
        rewritten
    })
}

enum ComponentRemoval {
    Unchanged,
    Rewritten(String),
    Collapsed,
}

/// What a sheet removal does to one component of a reference: nothing, a
/// narrowed 3-D span, or `#REF!` when nothing it named survives.
fn rewrite_component_for_removal(
    component: &str,
    removed_name: &str,
    order_before: &[String],
) -> ComponentRemoval {
    let mut replacements = Vec::new();
    let has_sheet = |name: &str| {
        order_before
            .iter()
            .any(|sheet| sheet_names_equal(sheet, name))
    };
    for qualifier in sheet_qualifiers_with_sheets(component, has_sheet) {
        if !qualifier.bound || qualifier.external {
            continue;
        }
        if !qualifier.is_span() {
            if sheet_names_equal(&qualifier.first, removed_name) {
                return ComponentRemoval::Collapsed;
            }
            continue;
        }
        if !sheet_names_equal(&qualifier.first, removed_name)
            && !sheet_names_equal(&qualifier.last, removed_name)
        {
            continue;
        }
        match narrowed_span(&qualifier, removed_name, order_before) {
            Some(Some(replacement)) => replacements.push((qualifier.span, replacement)),
            Some(None) => return ComponentRemoval::Collapsed,
            None => {}
        }
    }
    if replacements.is_empty() {
        return ComponentRemoval::Unchanged;
    }
    let mut out = String::with_capacity(component.len());
    let mut copied = 0;
    for (span, replacement) in replacements {
        out.push_str(&component[copied..span.start]);
        out.push_str(&replacement);
        copied = span.end;
    }
    out.push_str(&component[copied..]);
    ComponentRemoval::Rewritten(out)
}

/// What the defined-name rewriter made of one formula.
enum DefinedNameRewrite {
    Unchanged,
    Rewritten(String),
    Ambiguous,
    AffectedUnsupported,
    /// A component neither the formula parser nor the whole-axis reader
    /// accepts, so nothing in it can be moved.
    Unsupported,
}

/// Rewrites one defined-name formula: a union of components, each either a
/// whole-row/whole-column reference or an expression the formula parser reads.
fn rewrite_defined_name(
    source: &str,
    op: &Op,
    matches_target: &dyn Fn(&Option<String>) -> bool,
    global: bool,
    names: &DefinedNameSet,
    sheets: &SheetIndex<'_>,
) -> DefinedNameRewrite {
    let (prefix, source) = match source.strip_prefix('=') {
        Some(source) => ("=", source),
        None => ("", source),
    };
    if has_ambiguous_spaced_call(source) {
        return DefinedNameRewrite::Ambiguous;
    }
    let Some(components) = split_union(source) else {
        return DefinedNameRewrite::Unsupported;
    };
    let mut rewritten = Vec::with_capacity(components.len());
    let mut changed = false;
    for component in components {
        if let Some(axis) = AxisRange::parse(component) {
            if global && axis.sheet.is_none() && axis.axis.is_edited_by(op) {
                return DefinedNameRewrite::Ambiguous;
            }
            match axis.remapped(op, matches_target) {
                Remapped::Unchanged => rewritten.push(component.to_owned()),
                Remapped::Moved((start, end)) => {
                    changed = true;
                    rewritten.push(axis.to_formula(start, end));
                }
                Remapped::Deleted => {
                    changed = true;
                    rewritten.push(ErrorValue::Ref.as_str().to_owned());
                }
            }
            continue;
        }
        // Whole-column names still need the token rewriter's ambiguity checks.
        let Some(expr) = parse_formula_for_rewrite(component).ok().filter(|expr| {
            !contains_column_range(expr)
                && !contains_table_reference(expr)
                && !contains_range_join(expr)
                && !contains_spaced_cell_range(component)
        }) else {
            match rewrite_reference_tokens(component, op, matches_target, global, names, sheets) {
                DefinedNameRewrite::Unchanged => {
                    rewritten.push(component.to_owned());
                    continue;
                }
                DefinedNameRewrite::Rewritten(component) => {
                    changed = true;
                    rewritten.push(component);
                    continue;
                }
                DefinedNameRewrite::Unsupported
                    if global && mentions_unqualified_reference(component) =>
                {
                    return DefinedNameRewrite::Ambiguous;
                }
                refusal => return refusal,
            }
        };
        if global && contains_unqualified_reference(&expr) {
            return DefinedNameRewrite::Ambiguous;
        }
        let mut component_changed = false;
        let formula = transform(&expr, op, matches_target, &mut component_changed).to_formula();
        if !component_changed {
            rewritten.push(component.to_owned());
            continue;
        }
        if parse_formula_for_rewrite(&formula).is_err() {
            return DefinedNameRewrite::AffectedUnsupported;
        }
        changed = true;
        rewritten.push(formula);
    }
    if !changed {
        return DefinedNameRewrite::Unchanged;
    }
    DefinedNameRewrite::Rewritten(format!("{prefix}{}", rewritten.join(",")))
}

fn contains_spaced_cell_range(source: &str) -> bool {
    source.contains(':')
        && source.contains(char::is_whitespace)
        && lex_for_rewrite(source).is_ok_and(|tokens| {
            tokens.iter().any(|token| {
                let kind = match &token.kind {
                    TokKind::ThreeDRef { target, .. } => target.as_ref(),
                    kind => kind,
                };
                matches!(kind, TokKind::Range { .. })
                    && source[token.start..token.end].rsplit_once(':').is_some_and(
                        |(first, last)| {
                            first.ends_with(char::is_whitespace)
                                || last.starts_with(char::is_whitespace)
                        },
                    )
            })
        })
}

/// error literals, which name no cell and so survive any structural edit.
const ERROR_LITERALS: &[&str] = &[
    "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#SPILL!", "#CALC!",
];

/// The names a workbook defines, lowercased, as Excel matches them without
/// regard to case.
type DefinedNameSet = HashSet<String>;

fn defined_name_set(names: &[DefinedName]) -> DefinedNameSet {
    names
        .iter()
        .map(|defined| defined.name.to_lowercase())
        .collect()
}

/// One endpoint of a scanned reference: the text of its sheet qualifier, the
/// sheet that binds it, and the address after it.
struct Endpoint<'a> {
    qualifier: &'a str,
    sheet: Option<String>,
    address: &'a str,
    /// the `#` of a spill reference (`A1#`), which names whatever the formula
    /// at the address spilled and so moves with that address.
    spill: bool,
}

/// The far half of a range the address parser cannot read as one address,
/// because whitespace or a second qualifier sits inside it.
struct JoinedEndpoint<'a> {
    /// the range operator and any whitespace around it, kept so the rewrite
    /// prints the range the way the workbook wrote it.
    separator: &'a str,
    endpoint: Endpoint<'a>,
}

/// A reference the token scanner read: the span it occupies and its endpoints.
struct ReferenceToken<'a> {
    span: core::ops::Range<usize>,
    first: Endpoint<'a>,
    joined: Option<JoinedEndpoint<'a>>,
}

/// What sits at one position of a component the formula parser rejected.
enum Scanned<'a> {
    /// a token that names no cell; resume at this offset.
    Skip(usize),
    Reference(ReferenceToken<'a>),
    /// neither, so nothing can be said about what the component references.
    Unreadable,
}

/// One endpoint read at a position, and the offset just after it.
enum ScannedEndpoint<'a> {
    Skip(usize),
    Read(Endpoint<'a>, usize),
    Unreadable,
}

/// An address the scanner recognised.
enum Address<'a> {
    Cell(CellRef),
    Span(CellRange),
    Axis(AxisRange<'a>),
}

/// What one scanned token names.
enum Symbol<'a> {
    /// an address a structural edit moves.
    Address(Address<'a>),
    /// a defined name or a literal, neither of which the grid moves.
    Fixed,
    /// a token that could be either, or one the scanner does not know at all.
    Unknown,
}

/// Rewrites a component the formula parser cannot read by remapping its
/// references one token at a time. Whole-axis references, `Sheet!#REF!`, 3-D
/// qualifiers and the range operator applied to a call all defeat the lexer,
/// yet each is ordinary in a defined name, so the component is scanned with
/// the same tokenizer `mentions_sheet` uses and every address it holds is
/// moved by the same rules the parsed path applies. A token the scan cannot
/// pin down to an address or to something the grid never moves refuses the
/// edit, since a wrong rewrite is saved into the workbook.
fn rewrite_reference_tokens(
    source: &str,
    op: &Op,
    matches_target: &dyn Fn(&Option<String>) -> bool,
    global: bool,
    names: &DefinedNameSet,
    sheets: &SheetIndex<'_>,
) -> DefinedNameRewrite {
    let qualifiers = sheet_qualifiers_with_sheets(source, |name| sheets.has_sheet(name));
    let external: HashMap<usize, usize> = if source.contains('[') {
        lex_for_rewrite(source)
            .unwrap_or_default()
            .into_iter()
            .filter(|token| matches!(token.kind, TokKind::ExternalRef(_)))
            .map(|token| (token.start, token.end))
            .collect()
    } else {
        HashMap::new()
    };
    let bytes = source.as_bytes();
    let mut out = String::new();
    let mut copied = 0;
    let mut index = 0;
    while index < bytes.len() {
        if let Some(&end) = external.get(&index) {
            index = end;
            continue;
        }
        match bytes[index] {
            b'"' => {
                index = skip_string(source, index);
                continue;
            }
            b'{' => match skip_array_constant(source, index) {
                Some(end) => {
                    index = end;
                    continue;
                }
                None => return DefinedNameRewrite::Unsupported,
            },
            b'#' => match error_literal_end(source, index) {
                Some(end) => {
                    index = end;
                    continue;
                }
                None => return DefinedNameRewrite::Unsupported,
            },
            // Selectors and external text the lexer could not read.
            b'[' | b']' => return DefinedNameRewrite::Unsupported,
            _ => {}
        }
        let token = match scan_reference(source, index, &qualifiers) {
            Scanned::Skip(end) => {
                index = end;
                continue;
            }
            Scanned::Unreadable => return DefinedNameRewrite::Unsupported,
            Scanned::Reference(token) => token,
        };
        index = token.span.end;
        match remap_reference_token(&token, op, matches_target, global, names) {
            DefinedNameRewrite::Unchanged => {}
            DefinedNameRewrite::Rewritten(text) => {
                out.push_str(&source[copied..token.span.start]);
                out.push_str(&text);
                copied = token.span.end;
            }
            refusal => return refusal,
        }
    }
    if copied == 0 {
        return DefinedNameRewrite::Unchanged;
    }
    out.push_str(&source[copied..]);
    DefinedNameRewrite::Rewritten(out)
}

/// Reads the reference at `index`: one endpoint, or the two the range operator
/// joins when the address parser cannot read them as one address.
fn scan_reference<'a>(source: &'a str, index: usize, qualifiers: &[Qualifier]) -> Scanned<'a> {
    let (first, first_end) = match scan_endpoint(source, index, qualifiers) {
        ScannedEndpoint::Skip(end) => return Scanned::Skip(end),
        ScannedEndpoint::Unreadable => return Scanned::Unreadable,
        ScannedEndpoint::Read(endpoint, end) => (endpoint, end),
    };
    let Some(start) = join_start(source, first_end) else {
        return Scanned::Reference(ReferenceToken {
            span: index..first_end,
            first,
            joined: None,
        });
    };
    match scan_endpoint(source, start, qualifiers) {
        // the range operator applied to a call, which each side moves through
        // on its own.
        ScannedEndpoint::Skip(_) => Scanned::Reference(ReferenceToken {
            span: index..first_end,
            first,
            joined: None,
        }),
        ScannedEndpoint::Unreadable => Scanned::Unreadable,
        ScannedEndpoint::Read(endpoint, end) => Scanned::Reference(ReferenceToken {
            span: index..end,
            first,
            joined: Some(JoinedEndpoint {
                separator: &source[first_end..start],
                endpoint,
            }),
        }),
    }
}

/// Reads the endpoint at `index`: the implicit-intersection operator and the
/// sheet qualifier it carries, if any, and the address after them.
fn scan_endpoint<'a>(
    source: &'a str,
    index: usize,
    qualifiers: &[Qualifier],
) -> ScannedEndpoint<'a> {
    let bytes = source.as_bytes();
    // implicit intersection picks one cell out of the reference after it, so
    // it moves with that reference and falls with it.
    let token = index + usize::from(bytes.get(index) == Some(&b'@'));
    let Some(first) = parse_sheet_token(source, token) else {
        return ScannedEndpoint::Skip(index + next_char_len(source, index));
    };
    let qualified = qualifiers
        .binary_search_by_key(&token, |qualifier| qualifier.span.start)
        .ok()
        .map(|index| &qualifiers[index])
        .filter(|qualifier| qualifier.bound)
        .and_then(|qualifier| {
            let start = qualified_target(source, qualifier.span.end)?;
            let sheet = if qualifier.is_span() {
                format!("{}:{}", qualifier.first, qualifier.last)
            } else {
                qualifier.first.clone()
            };
            Some((start, sheet))
        });
    let (start, sheet) = match qualified {
        Some((start, sheet)) => (start, Some(sheet)),
        None if is_function_call(source, &first) => return ScannedEndpoint::Skip(first.end),
        None => (token, None),
    };
    if sheet.is_some() && bytes.get(start) == Some(&b'#') {
        // `Sheet!#REF!` is what Excel leaves behind; it names nothing to move.
        return match error_literal_end(source, start) {
            Some(end) => ScannedEndpoint::Skip(end),
            None => ScannedEndpoint::Unreadable,
        };
    }
    let Some(end) = address_end(source, start) else {
        return ScannedEndpoint::Unreadable;
    };
    let address = &source[start..end];
    let (address, spill) = match address.strip_suffix('#') {
        Some(anchor) => (anchor, true),
        None => (address, false),
    };
    ScannedEndpoint::Read(
        Endpoint {
            qualifier: &source[index..start],
            sheet,
            address,
            spill,
        },
        end,
    )
}

/// The end of the address at `start`: one endpoint, or two joined by `:` when
/// the address parser reads the pair as one address.
fn address_end(source: &str, start: usize) -> Option<usize> {
    let bytes = source.as_bytes();
    let first = parse_sheet_token(source, start)?;
    if bytes.get(first.end) != Some(&b':') {
        return Some(first.end);
    }
    let Some(second) = parse_sheet_token(source, first.end + 1) else {
        return Some(first.end);
    };
    if qualified_target(source, second.end).is_some() || is_function_call(source, &second) {
        return Some(first.end);
    }
    Some(second.end)
}

/// Where the endpoint on the far side of a `:` range operator starts, allowing
/// the whitespace Excel writes around it. `None` when no endpoint follows: the
/// address parser has already read the whole range, or an error literal marks
/// a far end that a past deletion stranded.
fn join_start(source: &str, end: usize) -> Option<usize> {
    let colon = skip_whitespace(source, end);
    if source.as_bytes().get(colon) != Some(&b':') {
        return None;
    }
    let start = skip_whitespace(source, colon + 1);
    error_literal_end(source, start).is_none().then_some(start)
}

fn skip_whitespace(source: &str, index: usize) -> usize {
    let mut index = index;
    while source[index..].starts_with(char::is_whitespace) {
        index += next_char_len(source, index);
    }
    index
}

/// The end of the error literal at `index`, or `None` when `#` starts
/// something the scanner does not know.
fn error_literal_end(source: &str, index: usize) -> Option<usize> {
    let rest = &source[index..];
    ERROR_LITERALS
        .iter()
        .find(|literal| {
            rest.get(..literal.len())
                .is_some_and(|head| head.eq_ignore_ascii_case(literal))
        })
        .map(|literal| index + literal.len())
}

fn classify_address(address: &str) -> Option<Address<'_>> {
    if let Ok(cell) = CellRef::parse_a1(address) {
        return Some(Address::Cell(cell));
    }
    if let Ok(span) = CellRange::parse_a1(address) {
        return Some(Address::Span(span));
    }
    AxisRange::parse(address).map(Address::Axis)
}

/// What a token names. A name Excel would also read as an address cannot be
/// told apart from that address, so it is neither moved nor left alone.
fn classify_symbol<'a>(address: &'a str, names: &DefinedNameSet) -> Symbol<'a> {
    let defined = names.contains(&address.to_lowercase());
    match classify_address(address) {
        Some(_) if defined => Symbol::Unknown,
        Some(address) => Symbol::Address(address),
        None if defined || is_literal(address) => Symbol::Fixed,
        None => Symbol::Unknown,
    }
}

/// Whether a token is a number or a boolean. Neither a defined name nor an
/// address may start with a digit, so a token that does and reads as no
/// address is a number.
fn is_literal(address: &str) -> bool {
    address.starts_with(|character: char| character.is_ascii_digit())
        || address.eq_ignore_ascii_case("TRUE")
        || address.eq_ignore_ascii_case("FALSE")
}

fn remap_reference_token(
    token: &ReferenceToken<'_>,
    op: &Op,
    matches_target: &dyn Fn(&Option<String>) -> bool,
    global: bool,
    names: &DefinedNameSet,
) -> DefinedNameRewrite {
    match &token.joined {
        Some(joined) => remap_joined(&token.first, joined, op, matches_target, global, names),
        None => remap_endpoint(&token.first, op, matches_target, global, names),
    }
}

/// Moves one scanned reference. A token that is a defined name rather than an
/// address stays put, as it does on the parsed path, and so does one on a
/// sheet the edit cannot reach. A deleted reference collapses to a bare
/// `#REF!` — the spelling the parsed path writes, rather than the
/// `Sheet!#REF!` Excel leaves behind — so both paths strand a name alike.
fn remap_endpoint(
    endpoint: &Endpoint<'_>,
    op: &Op,
    matches_target: &dyn Fn(&Option<String>) -> bool,
    global: bool,
    names: &DefinedNameSet,
) -> DefinedNameRewrite {
    if endpoint.sheet.is_some() && !matches_target(&endpoint.sheet) {
        return DefinedNameRewrite::Unchanged;
    }
    let address = match classify_symbol(endpoint.address, names) {
        Symbol::Address(address) => address,
        Symbol::Fixed if !endpoint.spill => return DefinedNameRewrite::Unchanged,
        _ => return DefinedNameRewrite::Unsupported,
    };
    if global && endpoint.sheet.is_none() {
        // a workbook name's unqualified reference binds to whichever sheet is
        // active, so only an edit that cannot reach it is safe to wave through.
        let reachable = match &address {
            Address::Axis(axis) => axis.axis.is_edited_by(op),
            _ => true,
        };
        if reachable {
            return DefinedNameRewrite::Ambiguous;
        }
    }
    if !matches_target(&endpoint.sheet) {
        return DefinedNameRewrite::Unchanged;
    }
    if endpoint.spill {
        let Address::Cell(cell) = address else {
            return DefinedNameRewrite::Unsupported;
        };
        return match remap_cell(cell, op) {
            Remapped::Unchanged => DefinedNameRewrite::Unchanged,
            Remapped::Moved(cell) => {
                DefinedNameRewrite::Rewritten(format!("{}{}#", endpoint.qualifier, cell.to_a1()))
            }
            Remapped::Deleted => DefinedNameRewrite::Rewritten(ErrorValue::Ref.as_str().to_owned()),
        };
    }
    let moved = match &address {
        Address::Cell(cell) => match remap_cell(*cell, op) {
            Remapped::Unchanged => return DefinedNameRewrite::Unchanged,
            Remapped::Moved(cell) => format!("{}{}", endpoint.qualifier, cell.to_a1()),
            Remapped::Deleted => ErrorValue::Ref.as_str().to_owned(),
        },
        Address::Span(span) => match remap_span(*span, op) {
            Remapped::Unchanged => return DefinedNameRewrite::Unchanged,
            Remapped::Moved(span) => format!("{}{}", endpoint.qualifier, span.to_a1()),
            Remapped::Deleted => ErrorValue::Ref.as_str().to_owned(),
        },
        Address::Axis(axis) => match axis.shifted(op) {
            Remapped::Unchanged => return DefinedNameRewrite::Unchanged,
            Remapped::Moved((start, end)) => {
                format!("{}{}", endpoint.qualifier, axis.to_formula(start, end))
            }
            Remapped::Deleted => ErrorValue::Ref.as_str().to_owned(),
        },
    };
    DefinedNameRewrite::Rewritten(moved)
}

/// Moves a range whose halves the address parser cannot read as one address.
/// Both endpoints are one span: clipping them apart would leave a deletion
/// that took the first of them stranded on `#REF!` while the second still
/// named a live cell. Endpoints on two different sheets, or either of them
/// naming anything but a cell, are forms the scanner cannot move.
fn remap_joined(
    first: &Endpoint<'_>,
    joined: &JoinedEndpoint<'_>,
    op: &Op,
    matches_target: &dyn Fn(&Option<String>) -> bool,
    global: bool,
    names: &DefinedNameSet,
) -> DefinedNameRewrite {
    let second = &joined.endpoint;
    if global && first.sheet.is_none() {
        // the near endpoint of a workbook name's range binds to whichever
        // sheet is active, wherever the far one points.
        return DefinedNameRewrite::Ambiguous;
    }
    if first.spill || second.spill || !endpoints_share_a_sheet(first, second) {
        return DefinedNameRewrite::Unsupported;
    }
    let (Symbol::Address(Address::Cell(start)), Symbol::Address(Address::Cell(end))) = (
        classify_symbol(first.address, names),
        classify_symbol(second.address, names),
    ) else {
        return DefinedNameRewrite::Unsupported;
    };
    if !matches_target(&first.sheet) {
        return DefinedNameRewrite::Unchanged;
    }
    match remap_span(CellRange::new(start, end), op) {
        Remapped::Unchanged => DefinedNameRewrite::Unchanged,
        Remapped::Moved(span) => DefinedNameRewrite::Rewritten(format!(
            "{}{}{}{}{}",
            first.qualifier,
            span.start.to_a1(),
            joined.separator,
            second.qualifier,
            span.end.to_a1()
        )),
        Remapped::Deleted => DefinedNameRewrite::Rewritten(ErrorValue::Ref.as_str().to_owned()),
    }
}

/// Whether both halves of a range name one sheet: a far endpoint carries the
/// near one's qualifier when it has none of its own.
fn endpoints_share_a_sheet(first: &Endpoint<'_>, second: &Endpoint<'_>) -> bool {
    match (&first.sheet, &second.sheet) {
        (_, None) => true,
        (None, Some(_)) => false,
        (Some(near), Some(far)) => sheet_names_equal(near, far),
    }
}

fn contains_unqualified_reference(expr: &Expr) -> bool {
    match expr {
        Expr::Ref { sheet: None, .. }
        | Expr::Range { sheet: None, .. }
        | Expr::ColumnRange { sheet: None, .. }
        | Expr::RowRange { sheet: None, .. } => true,
        Expr::Unary { expr, .. } | Expr::Percent(expr) => contains_unqualified_reference(expr),
        Expr::Binary { lhs, rhs, .. } => {
            contains_unqualified_reference(lhs) || contains_unqualified_reference(rhs)
        }
        Expr::RangeJoin { start, end } => {
            contains_unqualified_reference(start) || contains_unqualified_reference(end)
        }
        Expr::FuncCall { args, .. } => args.iter().any(contains_unqualified_reference),
        _ => false,
    }
}

/// A structural edit moves a table's rectangle, and the table part is not
/// remapped, so a formula reading one is left for the caller to refuse rather
/// than silently stranded on the pre-edit geometry.
fn contains_table_reference(expr: &Expr) -> bool {
    match expr {
        Expr::TableRef { .. } => true,
        Expr::Unary { expr, .. } | Expr::Percent(expr) => contains_table_reference(expr),
        Expr::Binary { lhs, rhs, .. } => {
            contains_table_reference(lhs) || contains_table_reference(rhs)
        }
        Expr::RangeJoin { start, end } => {
            contains_table_reference(start) || contains_table_reference(end)
        }
        Expr::FuncCall { args, .. } => args.iter().any(contains_table_reference),
        _ => false,
    }
}

/// A name written with the range operator keeps the token rewriter, whose
/// ambiguity checks the ast path does not reproduce — an unqualified end of a
/// workbook name binds to whichever sheet is active, so it cannot be moved.
fn contains_range_join(expr: &Expr) -> bool {
    match expr {
        Expr::RangeJoin { .. } => true,
        Expr::Unary { expr, .. } | Expr::Percent(expr) => contains_range_join(expr),
        Expr::Binary { lhs, rhs, .. } => contains_range_join(lhs) || contains_range_join(rhs),
        Expr::FuncCall { args, .. } => args.iter().any(contains_range_join),
        _ => false,
    }
}

fn contains_column_range(expr: &Expr) -> bool {
    match expr {
        Expr::ColumnRange { .. } | Expr::RowRange { .. } => true,
        Expr::Unary { expr, .. } | Expr::Percent(expr) => contains_column_range(expr),
        Expr::Binary { lhs, rhs, .. } => contains_column_range(lhs) || contains_column_range(rhs),
        Expr::RangeJoin { start, end } => {
            contains_column_range(start) || contains_column_range(end)
        }
        Expr::FuncCall { args, .. } => args.iter().any(contains_column_range),
        _ => false,
    }
}

fn mentions_unqualified_reference(source: &str) -> bool {
    let bytes = source.as_bytes();
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            index = skip_string(source, index);
            continue;
        }
        if bytes[index] == b'\'' {
            index = parse_sheet_token(source, index)
                .map(|token| token.end)
                .unwrap_or(source.len());
            continue;
        }
        if bytes[index] == b'[' {
            index += 1;
            while index < bytes.len() && bytes[index] != b']' {
                index += next_char_len(source, index);
            }
            index = (index + 1).min(bytes.len());
            continue;
        }
        if bytes[index] != b'$' && !bytes[index].is_ascii_alphanumeric() {
            index += next_char_len(source, index);
            continue;
        }
        let start = index;
        while index < bytes.len()
            && (bytes[index] == b'$'
                || bytes[index] == b':'
                || bytes[index].is_ascii_alphanumeric())
        {
            index += 1;
        }
        let candidate = &source[start..index];
        let qualified_before = source[..start].trim_end().ends_with('!');
        let qualified_after = source[index..].trim_start().starts_with('!');
        if !qualified_before
            && !qualified_after
            && (CellRef::parse_a1(candidate).is_ok()
                || CellRange::parse_a1(candidate).is_ok()
                || AxisRange::parse(candidate).is_some())
        {
            return true;
        }
    }
    false
}

/// Splits a formula on its top-level `,` union separators, keeping strings,
/// quoted sheet names, brackets and call arguments intact. `None` when those
/// do not balance.
fn split_union(source: &str) -> Option<Vec<&str>> {
    let bytes = source.as_bytes();
    let mut components = Vec::new();
    let mut start = 0;
    let mut index = 0;
    let mut depth = 0_u32;
    while index < bytes.len() {
        match bytes[index] {
            b'"' => index = skip_string(source, index),
            b'{' => index = skip_array_constant(source, index)?,
            b'\'' => index = parse_sheet_token(source, index)?.end,
            b'(' | b'[' => {
                depth += 1;
                index += 1;
            }
            b')' | b']' => {
                depth = depth.checked_sub(1)?;
                index += 1;
            }
            b',' if depth == 0 => {
                components.push(&source[start..index]);
                index += 1;
                start = index;
            }
            _ => index += next_char_len(source, index),
        }
    }
    if depth != 0 {
        return None;
    }
    components.push(&source[start..]);
    Some(components)
}

/// A whole-row or whole-column reference, including print titles.
struct AxisRange<'a> {
    qualifier: &'a str,
    sheet: Option<String>,
    axis: Axis,
    start: u32,
    end: u32,
    start_absolute: bool,
    end_absolute: bool,
}

impl<'a> AxisRange<'a> {
    fn parse(source: &'a str) -> Option<Self> {
        let source = source.trim();
        let (qualifier, sheet, rest) = match parse_sheet_token(source, 0) {
            Some(token) if source.as_bytes().get(token.end) == Some(&b'!') => (
                &source[..=token.end],
                Some(token.name),
                &source[token.end + 1..],
            ),
            _ => ("", None, source),
        };
        let (start_text, end_text) = rest.split_once(':')?;
        let (start_absolute, start_axis, start) = parse_axis_endpoint(start_text)?;
        let (end_absolute, end_axis, end) = parse_axis_endpoint(end_text)?;
        if start_axis != end_axis {
            return None;
        }
        Some(Self {
            qualifier,
            sheet,
            axis: start_axis,
            start: start.min(end),
            end: start.max(end),
            start_absolute,
            end_absolute,
        })
    }

    fn remapped(
        &self,
        op: &Op,
        matches_target: &dyn Fn(&Option<String>) -> bool,
    ) -> Remapped<(u32, u32)> {
        if !matches_target(&self.sheet) {
            return Remapped::Unchanged;
        }
        self.shifted(op)
    }

    /// The remap of an axis span already known to sit on the edited sheet.
    fn shifted(&self, op: &Op) -> Remapped<(u32, u32)> {
        if !self.axis.is_edited_by(op) {
            return Remapped::Unchanged;
        }
        let (start, end) = match *op {
            Op::DeleteRows { at, count, .. } | Op::DeleteCols { at, count, .. } => {
                match clip_interval(self.start, self.end, at, count) {
                    Some(interval) => interval,
                    None => return Remapped::Deleted,
                }
            }
            Op::InsertRows { at, count, .. } | Op::InsertCols { at, count, .. } => {
                let last = self.axis.limit() - 1;
                let shift = |index: u32| {
                    if index < at {
                        index
                    } else {
                        index.saturating_add(count).min(last)
                    }
                };
                (shift(self.start), shift(self.end))
            }
            Op::MoveRows {
                row: at,
                count,
                before,
                ..
            }
            | Op::MoveCols {
                col: at,
                count,
                before,
                ..
            } => {
                let start = move_axis(self.start, at, count, before);
                let end = move_axis(self.end, at, count, before);
                (start.min(end), start.max(end))
            }
            _ => return Remapped::Unchanged,
        };
        if (start, end) == (self.start, self.end) {
            return Remapped::Unchanged;
        }
        Remapped::Moved((start, end))
    }

    fn to_formula(&self, start: u32, end: u32) -> String {
        format!(
            "{}{}:{}",
            self.qualifier,
            self.axis.endpoint(start, self.start_absolute),
            self.axis.endpoint(end, self.end_absolute)
        )
    }
}

fn parse_axis_endpoint(source: &str) -> Option<(bool, Axis, u32)> {
    let absolute = source.starts_with('$');
    let text = source.strip_prefix('$').unwrap_or(source);
    if text.is_empty() {
        return None;
    }
    if text.bytes().all(|byte| byte.is_ascii_digit()) {
        let row: u32 = text.parse().ok()?;
        return (1..=MAX_ROWS)
            .contains(&row)
            .then_some((absolute, Axis::Row, row - 1));
    }
    if !text.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return None;
    }
    let cell = CellRef::parse_a1(&format!("{}1", text.to_ascii_uppercase())).ok()?;
    Some((absolute, Axis::Col, cell.col))
}

pub(crate) fn formula_has_sheet_span(source: &str, workbook: &Workbook) -> bool {
    sheet_qualifiers_with_sheets(source, |name| workbook_has_sheet(workbook, name))
        .iter()
        .any(|qualifier| {
            qualifier.bound && !qualifier.external && source[qualifier.span.clone()].contains(':')
        })
}

/// How a qualifier resolves against the workbook the edit targets.
struct SheetIndex<'a> {
    names: &'a HashMap<String, SheetId>,
    target: SheetId,
    target_name: &'a str,
}

impl SheetIndex<'_> {
    fn has_sheet(&self, name: &str) -> bool {
        self.names.contains_key(&name.to_lowercase())
    }

    /// Whether a qualifier covers the edited sheet. A span covers every sheet
    /// between its endpoints in workbook order; one whose endpoints this
    /// workbook does not hold falls back to naming them outright, so an
    /// unresolvable qualifier still binds rather than silently missing.
    fn covers(&self, first: &str, last: &str) -> bool {
        match (
            self.names.get(&first.to_lowercase()),
            self.names.get(&last.to_lowercase()),
        ) {
            (Some(start), Some(end)) => {
                (start.0.min(end.0)..=start.0.max(end.0)).contains(&self.target.0)
            }
            _ => {
                sheet_names_equal(first, self.target_name)
                    || sheet_names_equal(last, self.target_name)
            }
        }
    }
}

/// Whether a qualifier covering the edited sheet cannot be safely rewritten.
fn has_unrewritable_binding(source: &str, index: &SheetIndex<'_>) -> bool {
    sheet_qualifiers_with_sheets(source, |name| index.has_sheet(name))
        .iter()
        .any(|qualifier| {
            qualifier.bound
                && !qualifier.external
                && (qualifier.has_ambiguous_target(source)
                    || !(index.has_sheet(&qualifier.first) && index.has_sheet(&qualifier.last)))
                && index.covers(&qualifier.first, &qualifier.last)
        })
}

/// Whether an unrewritable formula names the edited sheet as a reference
/// qualifier, so leaving it untouched would strand it on the pre-edit
/// addresses.
fn mentions_sheet(source: &str, index: &SheetIndex<'_>) -> bool {
    sheet_qualifiers_with_sheets(source, |name| index.has_sheet(name))
        .iter()
        .any(|qualifier| {
            qualifier.bound
                && !qualifier.external
                && index.covers(&qualifier.first, &qualifier.last)
        })
}

pub(crate) fn remap_auto_filter(sheet: &mut xlsx_model::Sheet, id: SheetId, op: &Op) -> Option<Op> {
    let previous = sheet.auto_filter.clone();
    sheet.auto_filter = sheet.auto_filter.take().and_then(|mut filter| {
        let old_range = filter.range;
        filter.range = remap_hyperlink_range(old_range, op)?;
        filter.filter_columns = filter
            .filter_columns
            .iter()
            .filter_map(|&column| remap_filter_column(column, old_range, filter.range, op))
            .collect();
        if let Some(raw) = &filter.raw_xml
            && let Some((xml, has_sort_state)) =
                remap_auto_filter_xml(raw, old_range, filter.range, op)
        {
            filter.raw_xml = Some(xml);
            filter.has_sort_state = has_sort_state;
        }
        Some(filter)
    });
    (previous != sheet.auto_filter).then_some(Op::RestoreAutoFilter {
        sheet: id,
        filter: previous,
    })
}

fn remap_filter_column(column: u32, old: CellRange, new: CellRange, op: &Op) -> Option<u32> {
    if !matches!(
        op,
        Op::InsertCols { .. } | Op::DeleteCols { .. } | Op::MoveCols { .. }
    ) {
        return Some(column);
    }
    let absolute = old.start.col.checked_add(column)?;
    let moved = remap_ref(CellRef::new(0, absolute), op)?;
    // A saved `colId` is an offset from the region's first column, so a
    // criterion whose column leaves the region (a move can carry it past a
    // shrinking right edge) has no offset that names it. Drop it, as a delete
    // of the column does; `RestoreAutoFilter` puts it back on undo.
    if moved.col < new.start.col || moved.col > new.end.col {
        return None;
    }
    moved.col.checked_sub(new.start.col)
}

/// One `sortState` while the source is scanned: where it starts and how many
/// edits it owns, so a state that loses every key can be cut whole, the region
/// its keys must lie in, and how many keys it wrote and lost.
struct ScannedSortState {
    start: usize,
    first_edit: usize,
    remove: bool,
    region: Option<CellRange>,
    conditions: u32,
    dropped: u32,
}

impl ScannedSortState {
    fn new(start: usize, first_edit: usize) -> Self {
        Self {
            start,
            first_edit,
            remove: false,
            region: None,
            conditions: 0,
            dropped: 0,
        }
    }

    /// Whether the state has nothing left to sort by: a key that a delete
    /// stranded, or every key the region no longer holds.
    fn leaves_no_key(&self) -> bool {
        self.remove || (self.conditions > 0 && self.dropped == self.conditions)
    }
}

fn remap_auto_filter_xml(
    raw: &str,
    old: CellRange,
    new: CellRange,
    op: &Op,
) -> Option<(String, bool)> {
    use quick_xml::{Reader, events::Event};

    let mut reader = Reader::from_str(raw);
    let mut edits = Vec::new();
    let mut depth = 0;
    let mut sort_state: Option<ScannedSortState> = None;
    let mut has_sort_state = false;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().ok()?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let empty = matches!(event, Event::Empty(_));
                let name = element.local_name();
                let filter_column = depth == 1 && name.as_ref() == b"filterColumn";
                let sort = depth == 1 && name.as_ref() == b"sortState";
                let condition =
                    depth == 2 && sort_state.is_some() && name.as_ref() == b"sortCondition";
                if sort {
                    sort_state = Some(ScannedSortState::new(start, edits.len()));
                    has_sort_state = true;
                } else if condition {
                    sort_state.as_mut()?.conditions += 1;
                }
                let mut remove_column = false;
                let mut remove_condition = false;
                for attribute in element.attributes() {
                    let attribute = attribute.ok()?;
                    let value = attribute
                        .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                        .ok()?;
                    let value = value.as_ref();
                    let replacement = if filter_column && attribute.key.as_ref() == b"colId" {
                        let column = value.parse().ok()?;
                        match remap_filter_column(column, old, new, op) {
                            Some(moved) if moved == column => None,
                            Some(column) => Some(column.to_string()),
                            None => {
                                remove_column = true;
                                None
                            }
                        }
                    } else if depth == 0 && attribute.key.as_ref() == b"ref" {
                        (old != new).then(|| new.to_a1())
                    } else if attribute.key.as_ref() == b"ref" && (sort || condition) {
                        let range = CellRange::parse_a1(value).ok()?;
                        let state = sort_state.as_mut()?;
                        if condition
                            && matches!(op, Op::DeleteCols { .. })
                            && remap_ref(range.start, op).is_none()
                        {
                            state.remove = true;
                        }
                        match remap_hyperlink_range(range, op) {
                            Some(moved) => {
                                if sort {
                                    state.region = Some(moved);
                                }
                                // A sort key is a reference inside the sort
                                // region, like a criterion's `colId` is an
                                // offset inside the filter region: a move that
                                // shrinks the region can carry the key out of
                                // it, and the save would then name a column the
                                // region does not cover. Drop the key, and the
                                // whole state with it when no key is left.
                                if condition {
                                    let region = state.region.unwrap_or(new);
                                    if !(region.contains(moved.start) && region.contains(moved.end))
                                    {
                                        remove_condition = true;
                                        state.dropped += 1;
                                        continue;
                                    }
                                }
                                (moved != range).then(|| moved.to_a1())
                            }
                            None => {
                                if sort || condition {
                                    state.remove = true;
                                }
                                None
                            }
                        }
                    } else {
                        None
                    };
                    if let Some(replacement) = replacement
                        && replacement != value
                    {
                        let offset = start + 1 + attribute.value.as_ptr() as usize
                            - element.as_ref().as_ptr() as usize;
                        edits.push((offset..offset + attribute.value.len(), replacement));
                    }
                }
                if remove_column || remove_condition {
                    if !empty {
                        reader.read_to_end(element.name()).ok()?;
                    }
                    edits.push((start..reader.buffer_position() as usize, String::new()));
                    continue;
                }
                if empty {
                    if sort {
                        let state = sort_state.take()?;
                        if state.leaves_no_key() {
                            edits.truncate(state.first_edit);
                            edits.push((
                                state.start..reader.buffer_position() as usize,
                                String::new(),
                            ));
                            has_sort_state = false;
                        }
                    }
                } else {
                    depth += 1;
                }
            }
            Event::End(_) => {
                depth -= 1;
                if depth == 1
                    && let Some(state) = sort_state.take()
                    && state.leaves_no_key()
                {
                    edits.truncate(state.first_edit);
                    edits.push((
                        state.start..reader.buffer_position() as usize,
                        String::new(),
                    ));
                    has_sort_state = false;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let mut xml = raw.to_owned();
    for (span, replacement) in edits.into_iter().rev() {
        xml.replace_range(span, &replacement);
    }
    Some((xml, has_sort_state))
}

pub(crate) fn remap_hyperlink_range(range: CellRange, op: &Op) -> Option<CellRange> {
    match remap_span(range, op) {
        Remapped::Unchanged => Some(range),
        Remapped::Moved(range) => Some(range),
        Remapped::Deleted => None,
    }
}

pub(crate) fn rename_sheet_references(
    wb: &mut Workbook,
    old_name: &str,
    new_name: &str,
) -> Result<Vec<(SheetId, CellRef, CellState)>, OpError> {
    let mut restores = Vec::new();
    let mut edits = Vec::new();
    for (index, sheet) in wb.sheets.iter().enumerate() {
        let owner = SheetId(index as u32);
        for (cell, stored) in sheet.iter_cells() {
            let Some(source) = &stored.formula else {
                continue;
            };
            let rewritten = writable_rename(
                source,
                new_name,
                scan_sheet_rename(source, old_name, new_name, |name| {
                    workbook_has_sheet(wb, name)
                })
                .formula,
            )?;
            if rewritten != *source {
                if rewritten.len() > MAX_FORMULA_BYTES {
                    return Err(OpError::FormulaNotRewritable { sheet: owner, cell });
                }
                restores.push((owner, cell, CellState::from(stored)));
                edits.push((owner, cell, rewritten));
            }
        }
    }
    for (sheet, cell, formula) in edits {
        let stored = wb
            .sheet_mut(sheet)
            .and_then(|sheet| sheet.cell(cell).cloned());
        if let Some(mut stored) = stored {
            stored.formula = Some(formula);
            wb.sheet_mut(sheet)
                .expect("sheet exists")
                .set_cell(cell, stored);
        }
    }
    Ok(restores)
}

/// `text` with every structured reference to a table in `dropped` (`Table[...]`,
/// or the bare table name) replaced by `#REF!`, as Excel saves a reference to a
/// deleted table; `None` when nothing names one, or the text does not lex.
pub(crate) fn remove_table_references(
    text: &str,
    dropped: &[String],
    sheets: &[String],
) -> Option<String> {
    if dropped.is_empty() {
        return None;
    }
    let lower = text.to_lowercase();
    if !dropped
        .iter()
        .any(|table| lower.contains(&table.to_lowercase()))
    {
        return None;
    }
    let drops = |table: &str| dropped.iter().any(|name| sheet_names_equal(name, table));
    let tokens = lex_for_table_removal(text, dropped, sheets).ok()?;
    ensure_unambiguous_spaced_calls(text, &tokens).ok()?;
    let mut out = String::with_capacity(text.len());
    let mut copied = 0;
    for (at, token) in tokens.iter().enumerate() {
        let end = match &token.kind {
            TokKind::TableRef { table, .. } if drops(table) => token.end,
            TokKind::Ident(word)
                if drops(word)
                    && !matches!(
                        tokens.get(at + 1).map(|next| &next.kind),
                        Some(TokKind::LParen)
                    ) =>
            {
                token.end
            }
            _ => continue,
        };
        out.push_str(&text[copied..token.start]);
        out.push_str("#REF!");
        copied = end;
    }
    (copied > 0).then(|| {
        out.push_str(&text[copied..]);
        out
    })
}

fn lex_for_table_removal(
    text: &str,
    dropped: &[String],
    sheets: &[String],
) -> Result<Vec<Token>, ParseError> {
    lex_with_range_names(text, |name| {
        dropped.iter().any(|table| sheet_names_equal(table, name))
            && !sheets.iter().any(|sheet| sheet_names_equal(sheet, name))
    })
}

/// Whether `text` names a table in `dropped` in a form the rewrite cannot
/// read, so a deletion dropping that table would leave it behind.
pub(crate) fn names_unrewritable_table(text: &str, dropped: &[String], sheets: &[String]) -> bool {
    let lower = text.to_lowercase();
    dropped
        .iter()
        .any(|table| lower.contains(&table.to_lowercase()))
        && match lex_for_table_removal(text, dropped, sheets) {
            Err(_) => true,
            Ok(tokens) if ensure_unambiguous_spaced_calls(text, &tokens).is_err() => true,
            Ok(tokens) => {
                // Keep the deletion refusal for unquoted workbook filenames.
                tokens.iter().any(|token| {
                    matches!(token.kind, TokKind::ExternalRef(_))
                        && text[token.start..]
                            .strip_prefix('[')
                            .and_then(|source| source.split_once(']'))
                            .is_some_and(|(book, _)| {
                                !book.bytes().all(|byte| byte.is_ascii_digit())
                            })
                }) && remove_table_references(text, dropped, sheets).is_some()
            }
        }
}

/// Refuses a deletion that drops the tables in `dropped` when a cell formula
/// on another sheet, or a defined name, names one of them in text the
/// rewrite cannot read: it would keep naming a table that is gone.
pub(crate) fn ensure_table_references_rewritable(
    wb: &Workbook,
    removed: usize,
    dropped: &[String],
    sheets: &[String],
) -> Result<(), OpError> {
    if dropped.is_empty() {
        return Ok(());
    }
    for (index, sheet) in wb.sheets.iter().enumerate() {
        if index == removed {
            continue;
        }
        for (at, cell) in sheet.iter_cells() {
            if cell
                .formula
                .as_deref()
                .is_some_and(|formula| names_unrewritable_table(formula, dropped, sheets))
            {
                return Err(OpError::FormulaNotRewritable {
                    sheet: SheetId(index as u32),
                    cell: at,
                });
            }
        }
    }
    if let Some(defined) = wb
        .defined_names
        .iter()
        .find(|defined| names_unrewritable_table(&defined.formula, dropped, sheets))
    {
        return Err(OpError::DefinedNameNotRewritable {
            name: defined.name.clone(),
        });
    }
    Ok(())
}

/// The cell formulas on every sheet but `removed` that `lost` rewrites for
/// that sheet's deletion: the sheet's index before the removal, the cell, the
/// new formula and the cell as it was. Nothing changes here.
pub(crate) fn removed_sheet_cell_edits(
    wb: &Workbook,
    removed: usize,
    lost: &dyn Fn(&str) -> Option<String>,
) -> Result<Vec<(usize, CellRef, String, CellState)>, OpError> {
    let mut edits = Vec::new();
    for (index, sheet) in wb.sheets.iter().enumerate() {
        if index == removed {
            continue;
        }
        for (cell, stored) in sheet.iter_cells() {
            let Some(source) = &stored.formula else {
                continue;
            };
            let Some(rewritten) = lost(source) else {
                continue;
            };
            if rewritten.len() > MAX_FORMULA_BYTES {
                return Err(OpError::FormulaNotRewritable {
                    sheet: SheetId(index as u32),
                    cell,
                });
            }
            edits.push((index, cell, rewritten, CellState::from(stored)));
        }
    }
    Ok(edits)
}

/// Every sheet's conditional formats as a rename of `old_name` to `new_name`
/// leaves them: each rule formula naming the sheet rewritten as
/// [`rename_sheet_references`] rewrites cell formulas, and a rule kept as the
/// file wrote it patched to the same text. Only sheets with such a formula
/// are listed. Nothing changes here, so a refusal leaves the workbook as it
/// was.
pub(crate) fn renamed_conditional_formats(
    wb: &Workbook,
    old_name: &str,
    new_name: &str,
) -> Result<Vec<(SheetId, Vec<ConditionalFormatting>)>, OpError> {
    let rename = |source: &str| {
        scan_sheet_rename(source, old_name, new_name, |name| {
            workbook_has_sheet(wb, name)
        })
        .formula
    };
    rewritten_conditional_formats(
        wb,
        None,
        &|formula| writable_rename(formula, new_name, rename(formula)),
        &rename,
    )
}

/// Every sheet's conditional formats, but `skip`'s, with each rule formula
/// put through `rewrite`, and a rule kept as the file wrote it patched in
/// place with `patch` applied to its formula text, which has to agree with
/// the rule's formulas. Only sheets with a changed formula are listed, by
/// their index in `wb`. Nothing changes here.
pub(crate) fn rewritten_conditional_formats(
    wb: &Workbook,
    skip: Option<usize>,
    rewrite: &dyn Fn(&str) -> Result<String, OpError>,
    patch: &dyn Fn(&str) -> String,
) -> Result<Vec<(SheetId, Vec<ConditionalFormatting>)>, OpError> {
    let mut planned = Vec::new();
    for (index, sheet) in wb.sheets.iter().enumerate() {
        if skip == Some(index) {
            continue;
        }
        let mut formats = sheet.conditional_formats.clone();
        let mut changed = false;
        for rule in formats.iter_mut().flat_map(|group| group.rules.iter_mut()) {
            let mut renamed = Vec::with_capacity(rule.formulas.len());
            for formula in &rule.formulas {
                let rewritten = rewrite(formula)?;
                if rewritten.len() > MAX_FORMULA_BYTES {
                    return Err(unrewritable_rule(rule, &sheet.name));
                }
                renamed.push(rewritten);
            }
            if renamed == rule.formulas {
                continue;
            }
            if let Some(raw) = &rule.raw_xml {
                match rewrite_conditional_format_xml(raw, patch) {
                    Some((xml, authored)) if authored == rule.formulas => {
                        rule.raw_xml = Some(xml);
                    }
                    _ => return Err(unrewritable_rule(rule, &sheet.name)),
                }
            }
            rule.formulas = renamed;
            changed = true;
        }
        if changed {
            planned.push((SheetId(index as u32), formats));
        }
    }
    Ok(planned)
}

/// `escaped` with `also`, and every character XML 1.0 cannot carry literally,
/// written as decimal character references, as xlsx-parse writes formula text
/// and attribute values: a CR in element text, and a CR, LF or tab in an
/// attribute value, would otherwise read back normalized, and a control
/// character an authored formula brought in through a reference would be
/// invalid. A NUL, which no reference can carry, is written as it came.
fn char_references(escaped: &str, also: &[char]) -> String {
    let mut out = String::with_capacity(escaped.len());
    for ch in escaped.chars() {
        let xml_char = ch >= ' ' && ch != '\u{fffe}' && ch != '\u{ffff}';
        if ch != '\0' && (also.contains(&ch) || (!xml_char && !matches!(ch, '\t' | '\n'))) {
            out.push_str(&format!("&#{};", u32::from(ch)));
        } else {
            out.push(ch);
        }
    }
    out
}

fn unrewritable_rule(rule: &ConditionalFormatRule, sheet: &str) -> OpError {
    OpError::InvalidConditionalFormat(format!(
        "conditional format {} on sheet '{sheet}' cannot be rewritten for this sheet edit",
        rule.id
    ))
}

/// `raw`, a `cfRule` as the file wrote it, with `patch` applied to the text of
/// each `formula` element and each `cfvo` formula, plus the formula text it
/// held before, in document order. `None` when the fragment does not read.
fn rewrite_conditional_format_xml(
    raw: &str,
    patch: &dyn Fn(&str) -> String,
) -> Option<(String, Vec<String>)> {
    use quick_xml::escape::{escape, partial_escape};
    use quick_xml::{Reader, events::Event};

    let mut reader = Reader::from_str(raw);
    let mut edits = Vec::new();
    let mut authored = Vec::new();
    let mut formula: Option<(usize, String)> = None;
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event().ok()? {
            Event::Start(element) if element.local_name().as_ref() == b"formula" => {
                if formula.is_some() {
                    return None;
                }
                formula = Some((reader.buffer_position() as usize, String::new()));
            }
            Event::Text(text) if formula.is_some() => {
                formula.as_mut()?.1.push_str(&text.xml10_content().ok()?);
            }
            Event::CData(text) if formula.is_some() => {
                formula
                    .as_mut()?
                    .1
                    .push_str(std::str::from_utf8(&text.into_inner()).ok()?);
            }
            Event::GeneralRef(reference) if formula.is_some() => {
                let name = reference.decode().ok()?;
                let resolved = match name.as_ref() {
                    "lt" => '<',
                    "gt" => '>',
                    "amp" => '&',
                    "quot" => '"',
                    "apos" => '\'',
                    numeric => {
                        let code = numeric.strip_prefix('#')?;
                        let code = match code.strip_prefix('x') {
                            Some(hex) => u32::from_str_radix(hex, 16).ok()?,
                            None => code.parse().ok()?,
                        };
                        char::from_u32(code)?
                    }
                };
                formula.as_mut()?.1.push(resolved);
            }
            Event::End(element) if element.local_name().as_ref() == b"formula" => {
                let (content, text) = formula.take()?;
                let renamed = patch(&text);
                if renamed != text {
                    edits.push((
                        content..start,
                        char_references(&partial_escape(renamed.as_str()), &['\r']),
                    ));
                }
                authored.push(text);
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"cfvo" =>
            {
                let mut is_formula = false;
                let mut value = None;
                for attribute in element.attributes() {
                    let attribute = attribute.ok()?;
                    match attribute.key.as_ref() {
                        b"type" => {
                            is_formula = attribute
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .ok()?
                                == "formula";
                        }
                        b"val" => {
                            let text = attribute
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .ok()?
                                .into_owned();
                            let offset = start + 1 + attribute.value.as_ptr() as usize
                                - element.as_ref().as_ptr() as usize;
                            value = Some((offset..offset + attribute.value.len(), text));
                        }
                        _ => {}
                    }
                }
                if is_formula && let Some((span, text)) = value {
                    let renamed = patch(&text);
                    if renamed != text {
                        edits.push((
                            span,
                            char_references(&escape(renamed.as_str()), &['\r', '\n', '\t']),
                        ));
                    }
                    authored.push(text);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if formula.is_some() {
        return None;
    }
    let mut xml = raw.to_owned();
    for (span, replacement) in edits.into_iter().rev() {
        xml.replace_range(span, &replacement);
    }
    Some((xml, authored))
}

/// Splits a hyperlink location into its reference and any trailing `#` fragment.
fn split_location_reference(source: &str) -> Option<(Expr, &str)> {
    if let Ok(expr) = parse_formula_for_rewrite(source) {
        return Some((expr, ""));
    }
    let (reference, _) = source.rsplit_once('#')?;
    let expr = parse_formula_for_rewrite(reference).ok()?;
    Some((expr, &source[reference.len()..]))
}

/// rename the sheet inside an internal hyperlink location. locations are
/// commonly written `#Sheet!A1`; the leading `#` is not part of the reference
/// and must not reach the formula rewriter, but its presence is preserved.
pub(crate) fn rename_hyperlink_location(
    location: &str,
    old_name: &str,
    new_name: &str,
) -> Result<String, OpError> {
    let renamed = match location.strip_prefix('#') {
        Some(reference) => format!("#{}", rename_formula_sheet(reference, old_name, new_name)),
        None => rename_formula_sheet(location, old_name, new_name),
    };
    writable_rename(location, new_name, renamed)
}

/// `renamed`, the text `source` takes after a sheet rename, refused when the
/// rename wrote `new_name` into it and the name holds a character xml 1.0
/// cannot carry: formula text and hyperlink locations have no escape for one,
/// so it could only be written as invalid xml. Text that does not reference
/// the sheet is unchanged and never refused.
fn writable_rename(source: &str, new_name: &str, renamed: String) -> Result<String, OpError> {
    if renamed != source
        && !new_name.chars().all(|ch| {
            matches!(ch, '\t' | '\n' | '\r') || (ch >= ' ' && ch != '\u{fffe}' && ch != '\u{ffff}')
        })
    {
        return Err(OpError::SheetNameNotWritable);
    }
    Ok(renamed)
}

pub(crate) fn rename_formula_sheet(source: &str, old_name: &str, new_name: &str) -> String {
    if old_name == new_name {
        return source.to_owned();
    }
    rename_sheets(source, &[(old_name.to_owned(), new_name.to_owned())]).0
}

fn workbook_has_sheet(wb: &Workbook, name: &str) -> bool {
    wb.sheets
        .iter()
        .any(|sheet| sheet_names_equal(&sheet.name, name))
}

fn has_ambiguous_spaced_call(source: &str) -> bool {
    source.contains('(')
        && source.contains(char::is_whitespace)
        && lex_for_rewrite(source)
            .is_ok_and(|tokens| ensure_unambiguous_spaced_references(source, &tokens).is_err())
}

pub(crate) fn ensure_sheet_edit_references_rewritable(
    wb: &Workbook,
    name: &str,
    removal: Option<(usize, &[String])>,
) -> Result<(), OpError> {
    let blocked = |source: &str| {
        let spaced_call = OnceCell::new();
        source.contains('!')
            && sheet_qualifiers_with_sheets(source, |sheet| workbook_has_sheet(wb, sheet))
                .iter()
                .any(|qualifier| {
                    if !qualifier.bound || qualifier.external {
                        return false;
                    }
                    let affected = sheet_names_equal(&qualifier.first, name)
                        || sheet_names_equal(&qualifier.last, name)
                        || (qualifier.is_span()
                            && removal.is_some_and(|(_, order)| {
                                narrowed_span(qualifier, name, order).is_some()
                            }));
                    affected
                        && (qualifier.has_ambiguous_target(source)
                            || *spaced_call.get_or_init(|| has_ambiguous_spaced_call(source)))
                })
    };
    for (index, sheet) in wb.sheets.iter().enumerate() {
        if removal.is_some_and(|(removed, _)| removed == index) {
            continue;
        }
        for (cell, stored) in sheet.iter_cells() {
            if stored.formula.as_deref().is_some_and(blocked) {
                return Err(OpError::FormulaNotRewritable {
                    sheet: SheetId(index as u32),
                    cell,
                });
            }
        }
        for chart in &sheet.charts {
            if chart
                .refs
                .iter()
                .any(|reference| blocked(&reference.formula))
            {
                return Err(OpError::ChartRefNotRewritable {
                    part: chart.part.clone(),
                });
            }
        }
        for rule in sheet
            .conditional_formats
            .iter()
            .flat_map(|group| &group.rules)
        {
            if rule.formulas.iter().any(|formula| blocked(formula)) {
                return Err(unrewritable_rule(rule, &sheet.name));
            }
        }
    }
    for defined in &wb.defined_names {
        if !removal.is_some_and(|(index, _)| defined.local_sheet == Some(SheetId(index as u32)))
            && blocked(&defined.formula)
        {
            return Err(OpError::DefinedNameNotRewritable {
                name: defined.name.clone(),
            });
        }
    }
    Ok(())
}

/// A formula rewritten for a sheet rename, plus whether it still names the old
/// sheet in a position the rewriter could not resolve.
struct SheetRename {
    formula: String,
    ambiguous: bool,
}

/// Rewrites qualified sheet references and detects ambiguous bare names.
fn scan_sheet_rename(
    source: &str,
    old_name: &str,
    new_name: &str,
    has_sheet: impl Fn(&str) -> bool,
) -> SheetRename {
    if old_name == new_name {
        return SheetRename {
            formula: source.to_string(),
            ambiguous: false,
        };
    }
    let mut replacements = Vec::new();
    let mut ambiguous = false;
    for qualifier in sheet_qualifiers_with_sheets(source, has_sheet) {
        if qualifier.external {
            continue;
        }
        let renames_first = sheet_names_equal(&qualifier.first, old_name);
        let renames_last = sheet_names_equal(&qualifier.last, old_name);
        if !qualifier.bound {
            if renames_first && !qualifier.call {
                ambiguous = true;
            }
            continue;
        }
        if !renames_first && !renames_last {
            continue;
        }
        if qualifier.has_ambiguous_target(source) || has_ambiguous_spaced_call(source) {
            return SheetRename {
                formula: source.to_owned(),
                ambiguous: true,
            };
        }
        let first = if renames_first {
            new_name
        } else {
            &qualifier.first
        };
        let last = if renames_last {
            new_name
        } else {
            &qualifier.last
        };
        replacements.push((qualifier.span, qualifier_token(first, last)));
    }
    let formula = if replacements.is_empty() {
        source.to_string()
    } else {
        let mut output = String::with_capacity(source.len());
        let mut copied_until = 0;
        for (span, replacement) in replacements {
            output.push_str(&source[copied_until..span.start]);
            output.push_str(&replacement);
            copied_until = span.end;
        }
        output.push_str(&source[copied_until..]);
        output
    };
    SheetRename { formula, ambiguous }
}

/// Rewrites defined names through a sheet rename, dropping any whose formula
/// names the old sheet as a bare token, where the intent cannot be resolved.
/// Returns the rewritten list when it changed, leaving the workbook as is.
pub(crate) fn rename_defined_names(
    wb: &Workbook,
    old_name: &str,
    new_name: &str,
) -> Result<Option<Vec<DefinedName>>, OpError> {
    if old_name == new_name || wb.defined_names.is_empty() {
        return Ok(None);
    }
    let previous = &wb.defined_names;
    let mut rewritten: Vec<DefinedName> = Vec::with_capacity(previous.len());
    for defined in previous {
        let renamed = scan_sheet_rename(&defined.formula, old_name, new_name, |name| {
            workbook_has_sheet(wb, name)
        });
        if renamed.ambiguous {
            continue;
        }
        let mut defined = defined.clone();
        defined.formula = writable_rename(&defined.formula, new_name, renamed.formula)?;
        rewritten.push(defined);
    }
    if rewritten == *previous {
        return Ok(None);
    }
    if defined_names_bytes(&rewritten) > MAX_DEFINED_NAME_BYTES {
        return Err(OpError::TooManyDefinedNames);
    }
    Ok(Some(rewritten))
}

/// The end of the array constant at `start`, or `None` when its brace never
/// closes. Its cells are literals, so nothing inside names a cell.
fn skip_array_constant(source: &str, start: usize) -> Option<usize> {
    let bytes = source.as_bytes();
    let mut index = start + 1;
    while index < bytes.len() {
        match bytes[index] {
            b'}' => return Some(index + 1),
            b'"' => index = skip_string(source, index),
            _ => index += source[index..].chars().next().map_or(1, char::len_utf8),
        }
    }
    None
}

/// whether a reference read from `owner` points at the edited sheet.
/// unqualified refs bind to `owner`; a 3-D qualifier binds when the edited
/// sheet lies between its endpoints in workbook order.
fn resolves_to_target(ref_sheet: &Option<String>, owner: SheetId, index: &SheetIndex<'_>) -> bool {
    match ref_sheet {
        None => owner == index.target,
        Some(name) => {
            let (first, last) = quoted_endpoints(name);
            index.names.contains_key(&first.to_lowercase())
                && index.names.contains_key(&last.to_lowercase())
                && index.covers(&first, &last)
        }
    }
}

/// rebuild `expr`, remapping every reference to the edited sheet; sets
/// `changed` only when a reference actually moved.
fn transform(
    expr: &Expr,
    op: &Op,
    matches_target: &dyn Fn(&Option<String>) -> bool,
    changed: &mut bool,
) -> Expr {
    match expr {
        Expr::Ref { sheet, cell } if matches_target(sheet) => match remap_cell(*cell, op) {
            Remapped::Unchanged => expr.clone(),
            Remapped::Moved(new_cell) => {
                *changed = true;
                Expr::Ref {
                    sheet: sheet.clone(),
                    cell: new_cell,
                }
            }
            Remapped::Deleted => {
                *changed = true;
                Expr::Error(ErrorValue::Ref)
            }
        },
        Expr::Range { sheet, range } if matches_target(sheet) => match remap_span(*range, op) {
            Remapped::Unchanged => expr.clone(),
            Remapped::Moved(new_range) => {
                *changed = true;
                Expr::Range {
                    sheet: sheet.clone(),
                    range: new_range,
                }
            }
            Remapped::Deleted => {
                *changed = true;
                Expr::Error(ErrorValue::Ref)
            }
        },
        Expr::ColumnRange { sheet, range } if matches_target(sheet) => {
            match remap_columns(*range, op) {
                Remapped::Unchanged => expr.clone(),
                Remapped::Moved(range) => {
                    *changed = true;
                    Expr::ColumnRange {
                        sheet: sheet.clone(),
                        range,
                    }
                }
                Remapped::Deleted => {
                    *changed = true;
                    Expr::Error(ErrorValue::Ref)
                }
            }
        }
        Expr::RowRange { sheet, range } if matches_target(sheet) => match remap_rows(*range, op) {
            Remapped::Unchanged => expr.clone(),
            Remapped::Moved(range) => {
                *changed = true;
                Expr::RowRange {
                    sheet: sheet.clone(),
                    range,
                }
            }
            Remapped::Deleted => {
                *changed = true;
                Expr::Error(ErrorValue::Ref)
            }
        },
        Expr::Unary { op: u, expr: e } => Expr::Unary {
            op: *u,
            expr: Box::new(transform(e, op, matches_target, changed)),
        },
        Expr::Percent(e) => Expr::Percent(Box::new(transform(e, op, matches_target, changed))),
        Expr::Binary { op: b, lhs, rhs } => Expr::Binary {
            op: *b,
            lhs: Box::new(transform(lhs, op, matches_target, changed)),
            rhs: Box::new(transform(rhs, op, matches_target, changed)),
        },
        Expr::FuncCall { name, func, args } => Expr::FuncCall {
            name: name.clone(),
            func: *func,
            args: args
                .iter()
                .map(|a| transform(a, op, matches_target, changed))
                .collect(),
        },
        // two cell ends are one span: clipping them apart would strand the
        // near one on `#REF!` while the far one still named a live cell
        Expr::RangeJoin { start, end } => match joined_cells(start, end, matches_target) {
            Some((sheet, span)) => match remap_span(span, op) {
                Remapped::Unchanged => expr.clone(),
                Remapped::Moved(span) => {
                    *changed = true;
                    Expr::RangeJoin {
                        start: Box::new(Expr::Ref {
                            sheet: sheet.clone(),
                            cell: span.start,
                        }),
                        end: Box::new(Expr::Ref {
                            sheet,
                            cell: span.end,
                        }),
                    }
                }
                Remapped::Deleted => {
                    *changed = true;
                    Expr::Error(ErrorValue::Ref)
                }
            },
            None => Expr::RangeJoin {
                start: Box::new(transform(start, op, matches_target, changed)),
                end: Box::new(transform(end, op, matches_target, changed)),
            },
        },
        _ => expr.clone(),
    }
}

fn remap_columns(range: ColumnRange, op: &Op) -> Remapped<ColumnRange> {
    let axis = AxisRange {
        qualifier: "",
        sheet: None,
        axis: Axis::Col,
        start: range.start,
        end: range.end,
        start_absolute: range.abs_start,
        end_absolute: range.abs_end,
    };
    match axis.shifted(op) {
        Remapped::Unchanged => Remapped::Unchanged,
        Remapped::Moved((start, end)) => Remapped::Moved(ColumnRange {
            start,
            end,
            ..range
        }),
        Remapped::Deleted => Remapped::Deleted,
    }
}

fn remap_rows(range: RowRange, op: &Op) -> Remapped<RowRange> {
    let axis = AxisRange {
        qualifier: "",
        sheet: None,
        axis: Axis::Row,
        start: range.start,
        end: range.end,
        start_absolute: range.abs_start,
        end_absolute: range.abs_end,
    };
    match axis.shifted(op) {
        Remapped::Unchanged => Remapped::Unchanged,
        Remapped::Moved((start, end)) => Remapped::Moved(RowRange {
            start,
            end,
            ..range
        }),
        Remapped::Deleted => Remapped::Deleted,
    }
}

/// remap a single-cell reference through the op.
fn remap_cell(cell: CellRef, op: &Op) -> Remapped<CellRef> {
    match remap_ref(cell, op) {
        Some(new_cell) if new_cell == cell => Remapped::Unchanged,
        Some(new_cell) => Remapped::Moved(new_cell),
        None => Remapped::Deleted,
    }
}

/// remap a range: inserts shift both corners; deletes clip the span, collapsing
/// to `#REF!` only when the whole span is deleted.
/// a join whose two ends are plain cells on the sheet being edited, which a
/// structural edit moves as one span rather than as two references.
fn joined_cells(
    start: &Expr,
    end: &Expr,
    matches_target: &dyn Fn(&Option<String>) -> bool,
) -> Option<(Option<String>, CellRange)> {
    let (Expr::Ref { sheet: a, cell: s }, Expr::Ref { sheet: b, cell: e }) = (start, end) else {
        return None;
    };
    if a != b || !matches_target(a) {
        return None;
    }
    Some((a.clone(), CellRange::new(*s, *e)))
}

fn remap_span(range: CellRange, op: &Op) -> Remapped<CellRange> {
    match *op {
        Op::DeleteRows { at, count, .. } => clip_span(range, Axis::Row, at, count),
        Op::DeleteCols { at, count, .. } => clip_span(range, Axis::Col, at, count),
        Op::MoveRows { .. } | Op::MoveCols { .. } => move_span(range, op),
        // inserts only drop a corner on off-sheet-edge overflow
        _ => match (remap_ref(range.start, op), remap_ref(range.end, op)) {
            (Some(start), Some(end)) if start == range.start && end == range.end => {
                Remapped::Unchanged
            }
            (Some(start), Some(end)) => Remapped::Moved(CellRange { start, end }),
            _ => Remapped::Deleted,
        },
    }
}

/// A track move permutes the moved axis without dropping anything, so both
/// corners follow and only their order on that axis may swap. The other axis
/// stays with the corner it was written on.
fn move_span(range: CellRange, op: &Op) -> Remapped<CellRange> {
    let (Some(start), Some(end)) = (remap_ref(range.start, op), remap_ref(range.end, op)) else {
        return Remapped::Deleted;
    };
    let mut moved = CellRange { start, end };
    match *op {
        // the corners swap positions, so each absolute marker travels with the
        // coordinate it was written on (`$A1:B1` becomes `B1:$B1`)
        Op::MoveRows { .. } => {
            if moved.start.row > moved.end.row {
                std::mem::swap(&mut moved.start.row, &mut moved.end.row);
                std::mem::swap(&mut moved.start.abs_row, &mut moved.end.abs_row);
            }
        }
        _ => {
            if moved.start.col > moved.end.col {
                std::mem::swap(&mut moved.start.col, &mut moved.end.col);
                std::mem::swap(&mut moved.start.abs_col, &mut moved.end.abs_col);
            }
        }
    }
    if moved == range {
        Remapped::Unchanged
    } else {
        Remapped::Moved(moved)
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Axis {
    Row,
    Col,
}

impl Axis {
    fn is_edited_by(self, op: &Op) -> bool {
        matches!(
            (self, op),
            (
                Axis::Row,
                Op::InsertRows { .. } | Op::DeleteRows { .. } | Op::MoveRows { .. }
            ) | (
                Axis::Col,
                Op::InsertCols { .. } | Op::DeleteCols { .. } | Op::MoveCols { .. }
            )
        )
    }

    fn limit(self) -> u32 {
        match self {
            Axis::Row => MAX_ROWS,
            Axis::Col => MAX_COLS,
        }
    }

    fn endpoint(self, index: u32, absolute: bool) -> String {
        let anchor = if absolute { "$" } else { "" };
        match self {
            Axis::Row => format!("{anchor}{}", index + 1),
            Axis::Col => format!("{anchor}{}", col_to_letters(index)),
        }
    }
}

/// clip a range's span on one axis under a delete of `count` starting at `at`.
fn clip_span(range: CellRange, axis: Axis, at: u32, count: u32) -> Remapped<CellRange> {
    let (a, b) = match axis {
        Axis::Row => (range.start.row, range.end.row),
        Axis::Col => (range.start.col, range.end.col),
    };
    match clip_interval(a, b, at, count) {
        None => Remapped::Deleted,
        Some((na, nb)) if na == a && nb == b => Remapped::Unchanged,
        Some((na, nb)) => {
            let mut start = range.start;
            let mut end = range.end;
            match axis {
                Axis::Row => {
                    start.row = na;
                    end.row = nb;
                }
                Axis::Col => {
                    start.col = na;
                    end.col = nb;
                }
            }
            Remapped::Moved(CellRange { start, end })
        }
    }
}

/// clip the inclusive interval `[a, b]` under a delete of `count` indices at
/// `at`; `None` when it lies wholly inside the deleted span.
fn clip_interval(a: u32, b: u32, at: u32, count: u32) -> Option<(u32, u32)> {
    let end_del = at.saturating_add(count);
    if a >= at && b < end_del {
        return None;
    }
    let new_a = if a < at {
        a
    } else if a >= end_del {
        a - count
    } else {
        at
    };
    let new_b = if b < at {
        b
    } else if b >= end_del {
        b - count
    } else {
        at - 1
    };
    Some((new_a, new_b))
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{Cell, CellProvider, Hyperlink, Sheet};

    /// A structured reference to a dropped table, in either form and any
    /// case, becomes `#REF!`, also beside a reference into another workbook;
    /// other tables, string literals, functions, another workbook's names and
    /// a text that names no dropped table are left alone.
    #[test]
    fn structured_references_to_a_dropped_table_become_ref() {
        let dropped = ["Amounts".to_owned()];
        let sheets = [
            "Amounts".to_owned(),
            "Data".to_owned(),
            "Data two".to_owned(),
        ];
        for (source, expected) in [
            ("SUM(Amounts[Amt])", Some("SUM(#REF!)")),
            ("SUM(amounts[[#This Row],[Amt]])*2", Some("SUM(#REF!)*2")),
            ("Amounts[#All]", Some("#REF!")),
            ("ROWS(Amounts)", Some("ROWS(#REF!)")),
            (
                "SUM(Amounts[Amt])+SUM(Prices[Price])",
                Some("SUM(#REF!)+SUM(Prices[Price])"),
            ),
            ("SUM(Prices[Price])", None),
            ("\"Amounts[Amt]\"&A1", None),
            ("AmountsTotal+1", None),
            ("SUM(Data!A1)", None),
            ("[1]!B1+SUM(Amounts[Amt])", Some("[1]!B1+SUM(#REF!)")),
        ] {
            assert_eq!(
                remove_table_references(source, &dropped, &sheets).as_deref(),
                expected,
                "{source}"
            );
        }
        for (source, expected) in [
            (
                "[1]Data!$A$1+Amounts[[#This Row],[Amt]]",
                Some("[1]Data!$A$1+#REF!"),
            ),
            ("[1]!Amounts+1", None),
            ("SUM(Amounts:Data!A1)", None),
            ("SUM(Amounts:'Data two'!A1:B2)", None),
            ("SUM(Amounts:Data!Total)", Some("SUM(#REF!:Data!Total)")),
            ("[2]Amounts[Amt]", None),
            ("Data!Amounts", None),
        ] {
            assert_eq!(
                remove_table_references(source, &dropped, &sheets).as_deref(),
                expected,
                "{source}"
            );
        }
        assert!(names_unrewritable_table(
            "[Book.xlsx]Data!A1+SUM(Amounts[Amt])",
            &dropped,
            &sheets,
        ));
        assert!(!names_unrewritable_table(
            "[1]Data!A1+SUM(Amounts[Amt])",
            &dropped,
            &sheets,
        ));
        assert!(!names_unrewritable_table(
            "[Book.xlsx]Data!A1",
            &dropped,
            &sheets
        ));
        assert_eq!(
            remove_table_references("SUM(Amounts[Amt])", &[], &sheets),
            None
        );
    }

    #[test]
    fn a_deleted_3d_address_does_not_name_a_dropped_table() {
        let source = "SUM(Sales:Q2!#REF!)+SUM([1]Sales:Q2!#REF!)";
        let dropped = ["Sales".to_owned()];
        let sheets = ["Sales".to_owned(), "Q2".to_owned()];
        assert!(!names_unrewritable_table(source, &dropped, &sheets));
        assert_eq!(remove_table_references(source, &dropped, &sheets), None);
    }

    #[test]
    fn sheet_spans_with_r1c1_prefixes_survive_table_removal_and_follow_sheet_edits() {
        for first in ["RCost", "R2C3x"] {
            let sheets = [first, "Mid", "Sheet2"].map(str::to_owned);
            for target in ["A1", "#REF!"] {
                let source = format!("SUM({first}:Sheet2!{target})");
                assert_eq!(
                    remove_table_references(&source, &sheets[..1], &sheets),
                    None,
                );
                assert_eq!(
                    rename_formula_sheet(&source, "Sheet2", "New Sheet"),
                    format!("SUM('{first}:New Sheet'!{target})"),
                );
                assert_eq!(
                    rename_formula_sheet(&source, first, "Renamed"),
                    format!("SUM(Renamed:Sheet2!{target})"),
                );
                assert_eq!(
                    xlsx_model::sheet_qualifier::remove_sheet_references(
                        &source, "Sheet2", &sheets
                    ),
                    Some(format!("SUM({first}:Mid!{target})")),
                );
                assert_eq!(
                    xlsx_model::sheet_qualifier::remove_sheet_references(&source, first, &sheets),
                    Some(format!("SUM(Mid:Sheet2!{target})")),
                );
            }
        }
    }

    #[test]
    fn a_dropped_table_can_be_the_first_operand_of_a_qualified_range_join() {
        let sheets = ["Other".to_owned()];
        for (table, operand) in [
            ("Amounts", "Amounts:Other!B5"),
            ("Amounts", "Amounts:Other!A1:B5"),
            ("Amounts", "Amounts:Other!$A:$A"),
            ("Amounts", "Amounts:Other!1:3"),
            ("Amounts", "Amounts : Other!B5"),
            (r"\Table1", r"\Table1:Other!A1"),
            ("A1B", "A1B:Other!A1"),
            ("XFD1048577", "XFD1048577:Other!A1"),
        ] {
            let source = format!("IFERROR(SUM({operand}),0)");
            assert_eq!(
                remove_table_references(&source, &[table.to_owned()], &sheets),
                Some(source.replacen(table, "#REF!", 1)),
                "{source}"
            );
        }
    }

    #[test]
    fn a_short_dropped_table_name_before_if_or_let_is_rewritten() {
        for source in ["SUM(Amt:IF(C1,B1,B2))", "SUM(Amt:LET(x,B1,x))"] {
            assert_eq!(
                remove_table_references(source, &["Amt".to_owned()], &[]),
                Some(source.replace("Amt", "#REF!"))
            );
        }
    }

    #[test]
    fn parenthesized_intersections_keep_ranges_during_table_removal() {
        let dropped = ["Amt".to_owned()];
        for intersection in [
            "A:N (1:2)",
            "A:T (1:2)",
            "Amt:MAX (1:2)",
            "Amt:Tax (1:2)",
            "1:2 (A:B)",
            "A1:B2 (C1:D2)",
        ] {
            let source = format!("SUM({intersection})+Amt[Old]");
            let tokens = lex_for_rewrite(&source).unwrap();
            assert!(matches!(
                &tokens[2].kind,
                TokKind::ColumnRange { .. } | TokKind::RowRange { .. } | TokKind::Range { .. }
            ));
            assert_eq!(tokens[3].kind, TokKind::Intersection);
            assert!(!names_unrewritable_table(&source, &dropped, &[]));
            assert_eq!(
                remove_table_references(&source, &dropped, &[]),
                Some(format!("SUM({intersection})+#REF!")),
            );
        }
    }

    #[test]
    fn dropping_a_table_preserves_an_intersection_with_a_range_union() {
        let mut workbook = wb(&["Data", "Report"]);
        workbook.tables.push(xlsx_model::Table {
            name: "Amt".into(),
            sheet: SheetId(0),
            range: CellRange::parse_a1("A1:A2").unwrap(),
            header_rows: 1,
            totals_rows: 0,
            columns: vec!["Old".into()],
            style: Vec::new(),
        });
        workbook.defined_names = vec![defined("Input", "SUM(A:N (A1:B2,C1:D2))+Amt[Old]")];
        crate::apply(&mut workbook, &Op::RemoveSheet { index: 0 }).unwrap();
        assert!(workbook.tables.is_empty());
        assert_eq!(
            workbook.defined_names[0].formula,
            "SUM(A:N (A1:B2,C1:D2))+#REF!"
        );
    }

    #[test]
    fn sheets_can_be_renamed_and_removed_beside_range_unions_and_numeric_calls() {
        for (source, renamed, removed) in [
            (
                "SUM(Data!A:N (Data!1:2,Data!3:4))",
                "SUM('New Data'!A:N ('New Data'!1:2,'New Data'!3:4))",
                "SUM(#REF!:N (#REF!,#REF!))",
            ),
            (
                "SUM(LOG10 (1), Data!A1)",
                "SUM(LOG10 (1), 'New Data'!A1)",
                "SUM(LOG10 (1), #REF!)",
            ),
        ] {
            let mut workbook = wb(&["Data", "Report"]);
            set_formula(&mut workbook, SheetId(1), "B1", source);
            let mut edited = workbook.clone();
            crate::apply(
                &mut edited,
                &Op::RenameSheet {
                    sheet: SheetId(0),
                    name: "New Data".into(),
                },
            )
            .unwrap();
            assert_eq!(formula(&edited, SheetId(1), "B1").as_deref(), Some(renamed));

            crate::apply(&mut workbook, &Op::RemoveSheet { index: 0 }).unwrap();
            assert_eq!(
                formula(&workbook, SheetId(0), "B1").as_deref(),
                Some(removed)
            );
        }
    }

    #[test]
    fn unrelated_row_edits_distinguish_spaced_numeric_and_ambiguous_calls() {
        let source = "SUM(LOG10 (1), Data!A1)";
        let mut workbook = wb(&["Data", "Other"]);
        workbook.defined_names = vec![defined("Input", source)];
        crate::apply(&mut workbook, &insert_rows(1, 0, 1)).unwrap();
        assert_eq!(workbook.defined_names[0].formula, source);

        let source = "SUM(LOG10 (A1), Data!A1)";
        workbook.defined_names = vec![defined("Input", source)];
        assert_sheet_edit_refuses(
            &workbook,
            &insert_rows(1, 0, 1),
            &OpError::DefinedNameNotRewritable {
                name: "Input".into(),
            },
            source,
        );
    }

    #[test]
    fn ambiguous_spaced_calls_refuse_dropped_table_rewrites() {
        let dropped = ["Amt".to_owned()];
        for reference in [
            "LOG10 (1)",
            "LOG10 (A1)",
            "Amt:IF (C1,B1,B2)",
            "Amt:LET (x,B1,x)",
            "A:T (A1:B2,C1:D2)",
            "A:N (A1,B2)",
        ] {
            let source = format!("SUM({reference})+Amt[Old]");
            assert!(names_unrewritable_table(&source, &dropped, &[]), "{source}");
            assert_eq!(
                remove_table_references(&source, &dropped, &[]),
                None,
                "{source}"
            );
        }
    }

    #[test]
    fn external_table_selector_keeps_its_indexed_workbook_bang_requirement() {
        let dropped = ["Amounts".to_owned()];
        let invalid = "[2]Amounts[Amt]";
        let valid = "[2]!Amounts[Amt]";
        assert_eq!(
            lex_for_rewrite(invalid).unwrap_err().message,
            "expected an operator after the selector"
        );
        assert!(matches!(
            &lex_for_rewrite(valid).unwrap()[..],
            [xlsx_calc::Token { kind: TokKind::ExternalRef(target), .. }]
                if matches!(target.as_ref(), TokKind::TableRef { table, .. } if table == "Amounts")
        ));
        assert!(names_unrewritable_table(invalid, &dropped, &[]));
        assert!(!names_unrewritable_table(valid, &dropped, &[]));
        for source in [invalid, valid] {
            assert_eq!(remove_table_references(source, &dropped, &[]), None);
        }
    }

    #[test]
    fn mixed_quoted_range_join_keeps_same_named_table_and_sheet_untouched() {
        let source = "SUM(Amounts:'Data two'!A1:B2)";
        let sheets = ["Amounts", "Data two"].map(str::to_owned);
        let dropped = &sheets[..1];
        let tokens = lex_for_rewrite(source).unwrap();
        assert!(matches!(
            &tokens[2].kind,
            TokKind::ThreeDRef { first, last, .. } if first == "Amounts" && last == "Data two"
        ));
        assert!(!names_unrewritable_table(source, dropped, &sheets));
        assert_eq!(remove_table_references(source, dropped, &sheets), None);
    }

    #[test]
    fn mixed_quoted_span_removes_its_sheet_even_when_it_owns_a_same_named_table() {
        let source = "SUM(Amounts:'Data two'!A1:B2)";
        let report = SheetId(2);
        let mut workbook = wb(&["Amounts", "Data two", "Report"]);
        workbook.tables.push(xlsx_model::Table {
            name: "Amounts".into(),
            sheet: SheetId(0),
            range: CellRange::parse_a1("A1:A2").unwrap(),
            header_rows: 1,
            totals_rows: 0,
            columns: vec!["Amt".into()],
            style: Vec::new(),
        });
        set_formula(&mut workbook, report, "B1", source);
        workbook.defined_names = vec![defined("Span", source)];

        let mut renamed = workbook.clone();
        crate::apply(
            &mut renamed,
            &Op::RenameSheet {
                sheet: SheetId(0),
                name: "Totals".into(),
            },
        )
        .unwrap();
        assert_eq!(
            formula(&renamed, report, "B1").as_deref(),
            Some("SUM('Totals:Data two'!A1:B2)")
        );

        assert_sheet_edit_refuses(
            &workbook,
            &insert_rows(1, 0, 1),
            &OpError::FormulaNotRewritable {
                sheet: report,
                cell: r("B1"),
            },
            source,
        );

        crate::apply(&mut workbook, &Op::RemoveSheet { index: 0 }).unwrap();
        assert!(workbook.tables.is_empty());
        assert_eq!(
            formula(&workbook, SheetId(1), "B1").as_deref(),
            Some("SUM('Data two'!A1:B2)")
        );
        assert_eq!(workbook.defined_names[0].formula, "SUM('Data two'!A1:B2)");
    }

    #[test]
    fn mixed_quoted_span_endpoints_follow_sheet_renames_and_removals() {
        let source = "SUM(Jan:'End Sheet'!A1)";
        let report = SheetId(3);
        let mut workbook = wb(&["Jan", "Feb", "End Sheet", "Report"]);
        set_formula(&mut workbook, report, "B1", source);
        workbook.defined_names = vec![defined("Span", source)];
        for (op, owner, expected) in [
            (
                Op::RenameSheet {
                    sheet: SheetId(0),
                    name: "Start".into(),
                },
                report,
                "SUM('Start:End Sheet'!A1)",
            ),
            (
                Op::RemoveSheet { index: 0 },
                SheetId(2),
                "SUM('Feb:End Sheet'!A1)",
            ),
            (Op::RemoveSheet { index: 2 }, SheetId(2), "SUM(Jan:Feb!A1)"),
        ] {
            let mut edited = workbook.clone();
            crate::apply(&mut edited, &op).unwrap();
            assert_eq!(formula(&edited, owner, "B1").as_deref(), Some(expected));
            assert_eq!(edited.defined_names[0].formula, expected);
        }
    }

    #[test]
    fn mixed_quoted_spans_shift_names_and_charts_but_refuse_cell_remaps() {
        let source = "SUM(Jan:'End Sheet'!A1)";
        let report = SheetId(3);
        let mut workbook = wb(&["Jan", "Feb", "End Sheet", "Report"]);
        set_formula(&mut workbook, report, "B1", source);
        workbook.defined_names = vec![defined("Span", source)];
        charted(&mut workbook, report, &[source]);
        for target in [0, 1, 2] {
            let mut shifted = workbook.clone();
            let op = insert_rows(target, 0, 1);
            assert_sheet_edit_refuses(
                &workbook,
                &op,
                &OpError::FormulaNotRewritable {
                    sheet: report,
                    cell: r("B1"),
                },
                source,
            );
            remap_defined_names(&mut shifted, &op).unwrap();
            remap_charts(&mut shifted, &op).unwrap();
            assert_eq!(shifted.defined_names[0].formula, "SUM(Jan:'End Sheet'!A2)");
            assert_eq!(
                chart_formulas(&shifted, report),
                ["SUM(Jan:'End Sheet'!A2)"]
            );
        }
        assert!(formula_has_sheet_span(source, &workbook));
        workbook.sheets[3]
            .conditional_formats
            .push(ConditionalFormatting {
                sqref: "B1".into(),
                rules: vec![ConditionalFormatRule {
                    id: "span".into(),
                    rule_type: "expression".into(),
                    operator: None,
                    formulas: vec![source.into()],
                    priority: 1,
                    dxf_id: None,
                    format: Default::default(),
                    raw_xml: None,
                }],
            });
        assert_sheet_edit_refuses(
            &workbook,
            &insert_rows(1, 0, 1),
            &OpError::InvalidConditionalFormat(
                "structural edit would affect conditional format span on sheet 'Report' (sqref 'B1'); reference rewriting is unsupported".into(),
            ),
            source,
        );
    }

    #[test]
    fn local_qualified_and_external_selectors_do_not_block_table_removal() {
        let dropped = ["Prices".to_owned()];
        for prefix in [
            "[@Qty]",
            "SUM([@Col]:[@Col2])",
            "Sheet1!Table1[Col]",
            "'Sheet 1'!Table1[Col]",
            "[1]Sheet1!Prices[Amt]",
            "'[1]Sheet 1'!Prices[Amt]",
            r"C:\dir\[Book.xlsx]Sheet1!A1",
        ] {
            let source = format!("{prefix}+Prices[Amt]");
            assert!(
                !names_unrewritable_table(&source, &dropped, &[]),
                "{source}"
            );
            assert_eq!(
                remove_table_references(&source, &dropped, &[]),
                Some(format!("{prefix}+#REF!"))
            );
        }
        let literal = r#"[@Notes]&"Prices""#;
        assert!(!names_unrewritable_table(literal, &dropped, &[]));
        assert_eq!(remove_table_references(literal, &dropped, &[]), None);
    }

    #[test]
    fn removing_a_sheet_replaces_the_entire_qualified_table_reference() {
        let mut workbook = wb(&["Data", "Report"]);
        workbook.tables.push(xlsx_model::Table {
            name: "Prices".into(),
            sheet: SheetId(0),
            range: CellRange::parse_a1("A1:A2").unwrap(),
            header_rows: 1,
            totals_rows: 0,
            columns: vec!["Amt".into()],
            style: Vec::new(),
        });
        set_formula(
            &mut workbook,
            SheetId(1),
            "B1",
            "SUM(Data!Prices[Amt])+SUM(Prices:Report!A1)",
        );
        crate::apply(&mut workbook, &Op::RemoveSheet { index: 0 }).unwrap();
        assert_eq!(
            formula(&workbook, SheetId(0), "B1").as_deref(),
            Some("SUM(#REF!)+SUM(#REF!:Report!A1)")
        );
    }

    #[test]
    fn an_intersection_after_a_numeric_local_column_can_drop_its_table() {
        let dropped = ["Prices".to_owned()];
        for (source, expected) in [
            ("SUM([2024] Prices[Amt])", "SUM([2024] #REF!)"),
            ("SUM([1] Prices)", "SUM([1] #REF!)"),
        ] {
            assert_eq!(
                remove_table_references(source, &dropped, &[]).as_deref(),
                Some(expected)
            );
        }
    }

    #[test]
    fn dropping_a_table_rewrites_joins_with_missing_or_quoted_sheets() {
        let dropped = ["Amounts".to_owned()];
        let sheets = ["Amounts".to_owned(), "Data two".to_owned()];
        for (source, expected, sheet_names) in [
            (
                "SUM(Amounts:Missing!B5)",
                "SUM(#REF!:Missing!B5)",
                &sheets[1..],
            ),
            (
                "SUM(Amounts:'Data two'!B5)",
                "SUM(#REF!:'Data two'!B5)",
                &sheets[1..],
            ),
            (
                "SUM(Amounts:'Data two'!A1:B2)",
                "SUM(#REF!:'Data two'!A1:B2)",
                &sheets[1..],
            ),
            (
                "SUM(Amounts : 'Data two'!A1:B2)",
                "SUM(#REF! : 'Data two'!A1:B2)",
                &sheets[..],
            ),
            (
                "[1]!#REF!+SUM(Amounts[Amt])",
                "[1]!#REF!+SUM(#REF!)",
                &sheets[..],
            ),
        ] {
            assert!(!names_unrewritable_table(source, &dropped, sheet_names));
            assert_eq!(
                remove_table_references(source, &dropped, sheet_names).as_deref(),
                Some(expected),
                "{source}"
            );
        }
    }

    fn r(a1: &str) -> CellRef {
        CellRef::parse_a1(a1).unwrap()
    }

    /// workbook with the named sheets; caller populates cells.
    fn wb(names: &[&str]) -> Workbook {
        let mut wb = Workbook::default();
        for n in names {
            wb.sheets.push(Sheet::new(*n));
        }
        wb
    }

    fn set_formula(wb: &mut Workbook, sheet: SheetId, at: &str, f: &str) {
        wb.sheet_mut(sheet).unwrap().set_cell(
            r(at),
            Cell {
                formula: Some(f.to_string()),
                ..Default::default()
            },
        );
    }

    fn formula(wb: &Workbook, sheet: SheetId, at: &str) -> Option<String> {
        wb.formula(sheet, r(at)).map(str::to_string)
    }

    /// a join of two cells is one span: deleting the row its near end sits on
    /// clips the span rather than stranding that end on `#REF!` while the far
    /// end still names a live cell.
    #[test]
    fn a_deletion_inside_a_joined_range_clips_it_as_one_span() {
        let mut workbook = wb(&["Data"]);
        // written with a space the lexer reads this as a join rather than as
        // one range token
        set_formula(&mut workbook, SheetId(0), "D1", "SUM(A1: B4)");
        remap_formulas(
            &mut workbook,
            &Op::DeleteRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(
            formula(&workbook, SheetId(0), "D1").as_deref(),
            Some("SUM(A1:B3)")
        );
    }

    /// Moving one column of a joined range past the other swaps the corners;
    /// each `$` marker belongs to the coordinate it was written on, so it
    /// moves with that corner instead of staying on the endpoint. A and B
    /// trade places, and the range still covers the same two cells.
    #[test]
    fn a_span_whose_corners_swap_keeps_their_absolute_markers() {
        let mut workbook = wb(&["Data"]);
        set_formula(&mut workbook, SheetId(0), "E1", "SUM($A1:B1)");
        remap_formulas(
            &mut workbook,
            &Op::MoveCols {
                sheet: SheetId(0),
                col: 0,
                count: 1,
                before: 2,
            },
        )
        .unwrap();
        assert_eq!(
            formula(&workbook, SheetId(0), "E1").as_deref(),
            Some("SUM(A1:$B1)")
        );
    }

    #[test]
    fn whole_column_formulas_survive_row_edits_and_follow_column_edits() {
        for (op, expected) in [
            (
                Op::InsertRows {
                    sheet: SheetId(0),
                    at: 0,
                    count: 1,
                },
                "$S:V",
            ),
            (
                Op::DeleteRows {
                    sheet: SheetId(0),
                    at: 0,
                    count: 100,
                },
                "$S:V",
            ),
            (
                Op::InsertCols {
                    sheet: SheetId(0),
                    at: 0,
                    count: 1,
                },
                "$T:W",
            ),
            (
                Op::InsertCols {
                    sheet: SheetId(0),
                    at: 19,
                    count: 1,
                },
                "$S:W",
            ),
            (
                Op::DeleteCols {
                    sheet: SheetId(0),
                    at: 19,
                    count: 1,
                },
                "$S:U",
            ),
            (
                Op::DeleteCols {
                    sheet: SheetId(0),
                    at: 18,
                    count: 4,
                },
                "#REF!",
            ),
        ] {
            let mut workbook = wb(&["Data", "Other"]);
            set_formula(&mut workbook, SheetId(0), "A1", "VLOOKUP(2,$S:V,4,FALSE)");
            set_formula(&mut workbook, SheetId(1), "A1", "VLOOKUP(2,$S:V,4,FALSE)");
            set_formula(
                &mut workbook,
                SheetId(1),
                "A2",
                "VLOOKUP(2,Data!$S:V,4,FALSE)",
            );
            remap_formulas(&mut workbook, &op).unwrap();
            assert_eq!(
                formula(&workbook, SheetId(0), "A1").unwrap(),
                format!("VLOOKUP(2,{expected},4,FALSE)")
            );
            assert_eq!(
                formula(&workbook, SheetId(1), "A1").unwrap(),
                "VLOOKUP(2,$S:V,4,FALSE)"
            );
            let qualifier = if expected == "#REF!" { "" } else { "Data!" };
            assert_eq!(
                formula(&workbook, SheetId(1), "A2").unwrap(),
                format!("VLOOKUP(2,{qualifier}{expected},4,FALSE)")
            );
        }
    }

    fn charted(wb: &mut Workbook, sheet: SheetId, formulas: &[&str]) {
        wb.sheet_mut(sheet).unwrap().charts = vec![xlsx_model::SheetChart {
            raw_xml: None,
            part: "xl/charts/chart1.xml".to_owned(),
            drawing: "xl/drawings/drawing1.xml".to_owned(),
            anchor_index: 0,
            anchor: ChartAnchor::TwoCell {
                from: AnchorCell {
                    col: 2,
                    col_off: 0,
                    row: 4,
                    row_off: 0,
                },
                to: AnchorCell {
                    col: 8,
                    col_off: 0,
                    row: 19,
                    row_off: 0,
                },
                edit_as: AnchorEditAs::TwoCell,
            },
            refs: formulas
                .iter()
                .map(|formula| xlsx_model::ChartRef {
                    kind: xlsx_model::ChartRefKind::Values,
                    formula: (*formula).to_owned(),
                })
                .collect(),
        }];
    }

    fn chart_formulas(wb: &Workbook, sheet: SheetId) -> Vec<String> {
        wb.sheet(sheet).unwrap().charts[0]
            .refs
            .iter()
            .map(|reference| reference.formula.clone())
            .collect()
    }

    #[test]
    fn chart_references_shift_clip_and_collapse_like_cell_formulas() {
        let mut w = wb(&["Data", "Other"]);
        charted(
            &mut w,
            SheetId(1),
            &[
                "Data!$A$5",
                "Data!$A$1:$A$10",
                "Data!$A$6:$A$7",
                "Other!$A$5",
                "Data!$A:$A",
            ],
        );
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 5,
            count: 3,
        };
        let inverse = remap_charts(&mut w, &op).unwrap();
        assert_eq!(
            chart_formulas(&w, SheetId(1)),
            [
                "Data!$A$5",
                "Data!$A$1:$A$7",
                "#REF!",
                "Other!$A$5",
                "Data!$A:$A"
            ]
        );
        assert!(matches!(inverse.as_slice(), [Op::SetCharts { .. }]));
    }

    #[test]
    fn multi_area_chart_references_keep_their_parentheses() {
        let mut w = wb(&["Data"]);
        charted(
            &mut w,
            SheetId(0),
            &["(Data!$A$5:$A$6,Data!$C$5:$C$6)", "(Data!$A$1)"],
        );
        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 2,
        };
        remap_charts(&mut w, &op).unwrap();
        assert_eq!(
            chart_formulas(&w, SheetId(0)),
            ["(Data!$A$7:$A$8,Data!$C$7:$C$8)", "(Data!$A$3)"]
        );
    }

    #[test]
    fn unrewritable_chart_reference_aimed_at_the_edited_sheet_refuses_the_edit() {
        let mut w = wb(&["Data"]);
        charted(&mut w, SheetId(0), &["SUM(Data!Table1[Amount])"]);
        let original = w.clone();
        let error = remap_charts(
            &mut w,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap_err();
        assert_eq!(
            error,
            OpError::ChartRefNotRewritable {
                part: "xl/charts/chart1.xml".to_owned()
            }
        );
        assert_eq!(w, original, "a refusal must leave the workbook untouched");
    }

    #[test]
    fn unrewritable_chart_reference_aimed_elsewhere_survives_the_edit() {
        let mut w = wb(&["Data", "Other"]);
        charted(
            &mut w,
            SheetId(1),
            &[
                "OFFSET(Other!$A$1,0,0,COUNTA(Other!$A:$A),1)",
                "[Book.xlsx]Sheet1!$A$1",
                "",
            ],
        );
        remap_charts(
            &mut w,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(
            chart_formulas(&w, SheetId(1)),
            [
                "OFFSET(Other!$A$1,0,0,COUNTA(Other!$A:$A),1)",
                "[Book.xlsx]Sheet1!$A$1",
                ""
            ]
        );
    }

    #[test]
    fn chart_anchors_move_only_on_their_own_sheet() {
        let mut w = wb(&["Data", "Other"]);
        charted(&mut w, SheetId(1), &["Data!$A$1"]);
        remap_charts(
            &mut w,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 4,
            },
        )
        .unwrap();
        let ChartAnchor::TwoCell { from, .. } = w.sheets[1].charts[0].anchor else {
            panic!("two-cell anchor");
        };
        assert_eq!(from.row, 4, "an edit on another sheet cannot move it");
    }

    /// Moving the column a two-cell anchor's corner sits in past the other
    /// corner reverses the rectangle, so the two corners trade places. A
    /// corner is a cell *and* its offset, and the offset is measured against
    /// the column it was written in: swapping only the indices leaves the
    /// offset in a column of a different width, which can put the corners out
    /// of order and make the chart unrenderable.
    #[test]
    fn a_reversed_two_cell_anchor_swaps_each_corner_with_its_offset() {
        const COL_OFF: i64 = 80 * 9_525; // 80 px in EMU, inside a 20-character B
        let mut w = wb(&["Sheet1"]);
        w.sheets[0].charts = vec![xlsx_model::SheetChart {
            raw_xml: None,
            part: "xl/charts/chart1.xml".to_owned(),
            drawing: "xl/drawings/drawing1.xml".to_owned(),
            anchor_index: 0,
            anchor: ChartAnchor::TwoCell {
                from: AnchorCell {
                    col: 1,
                    col_off: COL_OFF,
                    row: 0,
                    row_off: 0,
                },
                to: AnchorCell {
                    col: 2,
                    col_off: 0,
                    row: 1,
                    row_off: 0,
                },
                edit_as: AnchorEditAs::TwoCell,
            },
            refs: Vec::new(),
        }];
        remap_charts(
            &mut w,
            &Op::MoveCols {
                sheet: SheetId(0),
                col: 1,
                count: 1,
                before: 4,
            },
        )
        .unwrap();
        let ChartAnchor::TwoCell { from, to, .. } = w.sheets[0].charts[0].anchor else {
            panic!("two-cell anchor");
        };
        assert_eq!(
            from,
            AnchorCell {
                col: 1,
                col_off: 0,
                row: 0,
                row_off: 0,
            },
            "the original right corner now leads, carrying its own offset"
        );
        assert_eq!(
            to,
            AnchorCell {
                col: 3,
                col_off: COL_OFF,
                row: 1,
                row_off: 0,
            },
            "the original left corner now trails, with the 80 px still inside it"
        );
    }

    #[test]
    fn chart_and_name_spans_follow_edits_on_any_sheet_they_cover() {
        for target in 0..3 {
            let mut workbook = wb(&["Jan", "Feb", "Mar", "Report"]);
            charted(&mut workbook, SheetId(3), &["'Jan:Mar'!$A$1:$A$5"]);
            workbook.defined_names = vec![defined("Span", "'Jan:Mar'!$A$1:$A$5")];
            let op = Op::InsertRows {
                sheet: SheetId(target),
                at: 0,
                count: 2,
            };
            remap_defined_names(&mut workbook, &op).unwrap();
            remap_charts(&mut workbook, &op).unwrap();
            assert_eq!(
                chart_formulas(&workbook, SheetId(3)),
                ["'Jan:Mar'!$A$3:$A$7"],
                "sheet {target}"
            );
            assert_eq!(
                workbook.defined_names[0].formula, "'Jan:Mar'!$A$3:$A$7",
                "sheet {target}"
            );
        }
    }

    #[test]
    fn proposal_remaps_refuse_local_and_external_sheet_spans() {
        let names = ["Jan", "Feb", "Mar", "Report"].map(str::to_owned);
        let op = insert_rows(1, 0, 1);
        for source in ["SUM(Jan:Mar!$A$1:$A$5)", "Feb!A1+'[1]Jan:Mar'!B1"] {
            assert_eq!(
                remap_proposal_formula(source, SheetId(3), &names, &op),
                None
            );
        }
    }

    #[test]
    fn proposal_remaps_refuse_unquoted_whole_row_sheet_spans() {
        let names = [
            "Sheet1", "Sheet2", "Sheet3", "Q1_Sales", "Q2_Sales", "Report",
        ]
        .map(str::to_owned);
        for (source, target) in [
            ("SUM(Sheet1:Sheet3!1:2)", 1),
            ("SUM(Sheet1:Sheet2!$1:$5)", 0),
            ("Sheet1:Sheet3!1:1", 2),
            ("SUM(Q1_Sales:Q2_Sales!1:10)", 3),
        ] {
            let op = insert_rows(target, 0, 1);
            assert_eq!(
                remap_proposal_formula(source, SheetId(5), &names, &op),
                None,
                "{source}"
            );
        }
    }

    #[test]
    fn let_range_joins_remap_unquoted_reference_shaped_sheet_names() {
        for sheet in ["R2C3", "A1", "TRUE"] {
            let names = [sheet, "Report"].map(str::to_owned);
            let source = format!("_xlfn.LET(_xlpm.c, {sheet}!A1, SUM(_xlpm.c:{sheet}!A2))");
            let expected = format!("_xlfn.LET(_xlpm.c,'{sheet}'!A2,SUM(_xlpm.c : '{sheet}'!A3))");
            assert_eq!(
                remap_proposal_formula(&source, SheetId(1), &names, &insert_rows(0, 0, 1)),
                Some(expected),
                "{source}"
            );
        }
    }

    #[test]
    fn row_insertion_rewrites_formulas_unsupported_for_evaluation() {
        let mut workbook = wb(&["B2", "Report"]);
        let owner = SheetId(1);
        let cases = [
            ("A1", "B2 !A1", "'B2'!A2"),
            ("A2", "SUM(B2!A1:B2!C3)", "SUM('B2'!A2:'B2'!C4)"),
            ("A3", "B2!A1:LOG10(1)", "'B2'!A2:LOG10(1)"),
            (
                "A4",
                r"'C:\Books\Data.xlsx'!A1+B2!A1",
                r"'C:\Books\Data.xlsx'!A1+'B2'!A2",
            ),
        ];
        for (address, source, _) in cases {
            set_formula(&mut workbook, owner, address, source);
        }
        remap_formulas(&mut workbook, &insert_rows(0, 0, 1)).unwrap();
        for (address, _, expected) in cases {
            assert_eq!(formula(&workbook, owner, address).as_deref(), Some(expected));
        }
    }

    #[test]
    fn structural_edits_move_local_references_beside_quoted_external_ones() {
        let mut workbook = wb(&["Data", "Report"]);
        set_formula(&mut workbook, SheetId(1), "A1", "'[1]Data'!B1");
        set_formula(&mut workbook, SheetId(1), "B1", "'[1]Data'!B1+Data!A1");
        workbook.defined_names = vec![defined("Mixed", "SUM([1]Data!B1,Data ! A1 : A2)")];
        let op = insert_rows(0, 0, 1);
        remap_formulas(&mut workbook, &op).unwrap();
        remap_defined_names(&mut workbook, &op).unwrap();
        assert_eq!(
            formula(&workbook, SheetId(1), "A1").as_deref(),
            Some("'[1]Data'!B1")
        );
        assert_eq!(
            formula(&workbook, SheetId(1), "B1").as_deref(),
            Some("'[1]Data'!B1+Data!A2")
        );
        assert_eq!(
            workbook.defined_names[0].formula,
            "SUM([1]Data!B1,Data ! A2 : A3)"
        );
    }

    #[test]
    fn structural_edits_refuse_external_sheet_spans_atomically() {
        let mut workbook = wb(&["Data", "Report"]);
        let owner = SheetId(1);
        let op = insert_rows(0, 0, 1);
        set_formula(&mut workbook, owner, "A1", "Data!A1");
        for source in [
            "'[1]Sheet1:Sheet2'!A1+Data!A1",
            "'[1]Sheet1':'Sheet2'!A1+Data!A1",
            "'[Book.xlsx]Sheet1':'Sheet2'!A1+Data!A1",
            "'[1]Sheet1':Sheet2!A1+Data!A1",
        ] {
            set_formula(&mut workbook, owner, "B1", source);
            assert_sheet_edit_refuses(
                &workbook,
                &op,
                &OpError::FormulaNotRewritable {
                    sheet: owner,
                    cell: r("B1"),
                },
                source,
            );
        }
    }

    /// A span whose endpoints this workbook does not hold cannot be resolved,
    /// so an edit that might move it is refused rather than left stale.
    #[test]
    fn an_unresolvable_span_refuses_an_edit_on_a_sheet_it_names() {
        let mut workbook = wb(&["Jan", "Report"]);
        charted(&mut workbook, SheetId(1), &["'Jan:Missing'!$A$1:$A$5"]);
        let error = remap_charts(
            &mut workbook,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 2,
            },
        )
        .unwrap_err();
        assert!(
            matches!(error, OpError::ChartRefNotRewritable { .. }),
            "{error:?}"
        );
    }

    /// Renaming an endpoint rewrites the span as Excel writes it, quoting the
    /// whole qualifier rather than one half of it.
    #[test]
    fn renaming_an_endpoint_rewrites_the_whole_span() {
        let mut workbook = wb(&["Jan", "Feb", "Mar"]);
        charted(
            &mut workbook,
            SheetId(2),
            &["'Jan:Mar'!$A$1", "Jan:Mar!$B$1"],
        );
        rename_chart_refs(&mut workbook, "Jan", "New Year");
        assert_eq!(
            chart_formulas(&workbook, SheetId(2)),
            ["'New Year:Mar'!$A$1", "'New Year:Mar'!$B$1"]
        );

        let mut simple = wb(&["Jan", "Feb", "Mar"]);
        charted(&mut simple, SheetId(2), &["Jan:Mar!$A$1"]);
        rename_chart_refs(&mut simple, "Mar", "Dec");
        assert_eq!(chart_formulas(&simple, SheetId(2)), ["Jan:Dec!$A$1"]);
    }

    /// Removing an endpoint narrows the span onto the sheets that remain;
    /// removing an interior sheet leaves it alone; removing every sheet it
    /// covered collapses it.
    #[test]
    fn removing_a_sheet_narrows_the_spans_that_covered_it() {
        let order = ["Jan", "Feb", "Mar", "Report"].map(str::to_owned);
        for (removed, expected) in [
            ("Jan", "Feb:Mar!$A$1"),
            ("Feb", "Jan:Mar!$A$1"),
            ("Mar", "Jan:Feb!$A$1"),
        ] {
            let mut workbook = wb(&["Jan", "Feb", "Mar", "Report"]);
            charted(&mut workbook, SheetId(3), &["Jan:Mar!$A$1"]);
            let index = order.iter().position(|name| name == removed).unwrap();
            workbook.sheets.remove(index);
            strand_chart_refs(&mut workbook, removed, &order);
            assert_eq!(
                chart_formulas(&workbook, SheetId(2)),
                [expected],
                "removing {removed}"
            );
        }

        let single = ["Jan".to_owned(), "Report".to_owned()];
        let mut workbook = wb(&["Jan", "Report"]);
        charted(&mut workbook, SheetId(1), &["Jan:Jan!$A$1"]);
        workbook.sheets.remove(0);
        strand_chart_refs(&mut workbook, "Jan", &single);
        assert_eq!(chart_formulas(&workbook, SheetId(0)), ["#REF!"]);
    }

    #[test]
    fn removing_a_sheet_collapses_the_chart_references_into_it() {
        let mut w = wb(&["Data", "Report"]);
        charted(
            &mut w,
            SheetId(1),
            &["Data!$A$1:$A$4", "(Data!$A$1,Report!$B$1)", "Report!$C$1"],
        );
        w.sheets.remove(0);
        let inverse = strand_chart_refs(&mut w, "Data", &["Data".to_owned(), "Report".to_owned()]);
        assert_eq!(
            chart_formulas(&w, SheetId(0)),
            ["#REF!", "(#REF!,Report!$B$1)", "Report!$C$1"]
        );
        assert!(matches!(inverse.as_slice(), [Op::SetCharts { .. }]));
    }

    #[test]
    fn shifts_single_cell_ref_on_insert() {
        let mut w = wb(&["Sheet1"]);
        set_formula(&mut w, SheetId(0), "B1", "A5+1");
        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 2,
            count: 3,
        };
        let inv = remap_formulas(&mut w, &op).unwrap();
        assert_eq!(formula(&w, SheetId(0), "B1").as_deref(), Some("A8+1"));
        assert_eq!(inv.len(), 1);
    }

    #[test]
    fn clips_range_on_delete() {
        let mut w = wb(&["Sheet1"]);
        set_formula(&mut w, SheetId(0), "C1", "SUM(A1:A10)");
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        };
        remap_formulas(&mut w, &op).unwrap();
        assert_eq!(formula(&w, SheetId(0), "C1").as_deref(), Some("SUM(A1:A9)"));
    }

    #[test]
    fn deleted_ref_becomes_ref_error() {
        let mut w = wb(&["Sheet1"]);
        set_formula(&mut w, SheetId(0), "B1", "A5*2");
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 4,
            count: 1,
        };
        remap_formulas(&mut w, &op).unwrap();
        assert_eq!(formula(&w, SheetId(0), "B1").as_deref(), Some("#REF!*2"));
    }

    #[test]
    fn fully_deleted_range_becomes_ref_error() {
        let mut w = wb(&["Sheet1"]);
        set_formula(&mut w, SheetId(0), "B1", "SUM(A5:A7)");
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 4,
            count: 3,
        };
        remap_formulas(&mut w, &op).unwrap();
        assert_eq!(formula(&w, SheetId(0), "B1").as_deref(), Some("SUM(#REF!)"));
    }

    #[test]
    fn preserves_dollar_anchors() {
        let mut w = wb(&["Sheet1"]);
        set_formula(&mut w, SheetId(0), "B1", "$A$5+1");
        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 2,
        };
        remap_formulas(&mut w, &op).unwrap();
        assert_eq!(formula(&w, SheetId(0), "B1").as_deref(), Some("$A$7+1"));
    }

    #[test]
    fn remaps_cross_sheet_ref_to_edited_sheet() {
        let mut w = wb(&["Sheet1", "Data"]);
        set_formula(&mut w, SheetId(1), "A1", "Sheet1!A5+1");
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        };
        remap_formulas(&mut w, &op).unwrap();
        assert_eq!(
            formula(&w, SheetId(1), "A1").as_deref(),
            Some("Sheet1!A4+1")
        );
    }

    #[test]
    fn structural_rewrite_keeps_cell_like_sheet_name_quoted() {
        let mut workbook = wb(&["A1", "Formula"]);
        set_formula(&mut workbook, SheetId(1), "A1", "'A1'!A1");
        remap_formulas(
            &mut workbook,
            &Op::InsertRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(
            formula(&workbook, SheetId(1), "A1").as_deref(),
            Some("'A1'!A2")
        );
    }

    #[test]
    fn leaves_other_sheet_refs_untouched() {
        let mut w = wb(&["Sheet1", "Data"]);
        set_formula(&mut w, SheetId(1), "A1", "A5+1");
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        };
        let inv = remap_formulas(&mut w, &op).unwrap();
        assert_eq!(formula(&w, SheetId(1), "A1").as_deref(), Some("A5+1"));
        assert!(inv.is_empty(), "no formula changed, no inverse");
    }

    #[test]
    fn unparseable_formula_rejects_structural_rewrite() {
        let mut w = wb(&["Sheet1"]);
        set_formula(&mut w, SheetId(0), "B1", "SUM(");
        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        };
        let error = remap_formulas(&mut w, &op).unwrap_err();
        assert_eq!(formula(&w, SheetId(0), "B1").as_deref(), Some("SUM("));
        assert_eq!(
            error,
            OpError::FormulaNotRewritable {
                sheet: SheetId(0),
                cell: r("B1")
            }
        );
    }

    #[test]
    fn rename_preserves_formula_source_and_unsupported_syntax() {
        let source = " SUM( Sheet1!A:A , \"Sheet1!A1\", 'SHEET1'!B2 ) ";
        assert_eq!(
            rename_formula_sheet(source, "Sheet1", "New Name"),
            " SUM( 'New Name'!A:A , \"Sheet1!A1\", 'New Name'!B2 ) "
        );
    }

    #[test]
    fn rename_escapes_quotes_in_new_sheet_name() {
        assert_eq!(
            rename_formula_sheet("Sheet1!A1", "Sheet1", "Owner's Data"),
            "'Owner''s Data'!A1"
        );
    }

    #[test]
    fn rename_handles_3d_refs_without_touching_external_or_structured_refs() {
        let source = "Sheet1:Sheet3!A1+[Book.xlsx]Sheet1!A1+Table1[Sheet1!Column]+Sheet1!A1";
        assert_eq!(
            rename_formula_sheet(source, "Sheet1", "Renamed"),
            "Renamed:Sheet3!A1+[Book.xlsx]Sheet1!A1+Table1[Sheet1!Column]+Renamed!A1"
        );
    }

    #[test]
    fn sheet_rename_updates_spaced_and_joined_qualifiers_in_formulas_and_names() {
        let source = "Tax ! A1+SUM(Amt:Tax!Total)+SUM(A1B:Tax!A1)+'[Book.xlsx]Data:Tax'!A1";
        let expected = "'New Tax' ! A1+SUM(Amt:'New Tax'!Total)+SUM(A1B:'New Tax'!A1)+'[Book.xlsx]Data:Tax'!A1";
        let mut workbook = wb(&["Tax", "Data"]);
        set_formula(&mut workbook, SheetId(1), "A1", source);
        workbook.defined_names = vec![defined("Test", source)];
        crate::apply(
            &mut workbook,
            &Op::RenameSheet {
                sheet: SheetId(0),
                name: "New Tax".into(),
            },
        )
        .unwrap();
        assert_eq!(
            formula(&workbook, SheetId(1), "A1").as_deref(),
            Some(expected)
        );
        assert_eq!(workbook.defined_names[0].formula, expected);
    }

    #[test]
    fn conditional_format_rename_guard_recognizes_only_local_sheet_spans() {
        let mut workbook = wb(&["Data", "Report"]);
        workbook.sheets[1]
            .conditional_formats
            .push(ConditionalFormatting {
                sqref: "A1".into(),
                rules: vec![ConditionalFormatRule {
                    id: "rule".into(),
                    rule_type: "expression".into(),
                    operator: None,
                    formulas: vec!["SUM(A1B:Data!B5)+'[1]Data:Report'!C1".into()],
                    priority: 1,
                    dxf_id: None,
                    format: Default::default(),
                    raw_xml: None,
                }],
            });
        let op = Op::RenameSheet {
            sheet: SheetId(0),
            name: "New Data".into(),
        };
        assert!(
            crate::conditional_formats::ensure_conditional_formats_allow_edit(&workbook, &op)
                .is_ok()
        );
        workbook.sheets[1].conditional_formats[0].rules[0].formulas =
            vec!["SUM(Data:Report!B5)".into()];
        assert!(
            crate::conditional_formats::ensure_conditional_formats_allow_edit(&workbook, &op)
                .is_err()
        );
    }

    #[test]
    fn rename_rejects_formula_growth_past_the_length_cap() {
        let mut workbook = wb(&["S", "Formula"]);
        set_formula(&mut workbook, SheetId(1), "A1", "S!A1");
        let original = workbook.clone();
        let error = rename_sheet_references(&mut workbook, "S", &"x".repeat(MAX_FORMULA_BYTES))
            .unwrap_err();
        assert!(matches!(error, OpError::FormulaNotRewritable { .. }));
        assert_eq!(workbook, original);
    }

    fn assert_sheet_edit_refuses(workbook: &Workbook, op: &Op, expected: &OpError, source: &str) {
        let mut edited = workbook.clone();
        assert_eq!(
            &crate::apply(&mut edited, op).unwrap_err(),
            expected,
            "{source}"
        );
        assert_eq!(&edited, workbook, "{source}");
    }

    #[test]
    fn qualified_local_selector_intersections_refuse_sheet_edits() {
        let owner = SheetId(1);
        let target = SheetId(0);
        let mut workbook = wb(&["Data", "Report"]);
        let error = OpError::FormulaNotRewritable {
            sheet: owner,
            cell: r("B1"),
        };
        for source in [
            "SUM(Data!$B$2:$B$4 [@Status])",
            "SUM(Data!A1:B2 [Col])",
            "SUM(Data!A:B [@Status])",
            "SUM(Data!1:2 [@Status])",
        ] {
            set_formula(&mut workbook, owner, "B1", source);
            for op in [
                Op::RemoveSheet { index: 0 },
                Op::RenameSheet {
                    sheet: target,
                    name: "New Data".into(),
                },
                insert_rows(0, 0, 1),
            ] {
                assert_sheet_edit_refuses(&workbook, &op, &error, source);
            }
        }
    }

    #[test]
    fn sheet_span_intersections_and_chained_colons_refuse_sheet_edits() {
        let owner = SheetId(3);
        let mut workbook = wb(&["Jan", "Feb", "Mar", "Report"]);
        let error = OpError::FormulaNotRewritable {
            sheet: owner,
            cell: r("B1"),
        };
        for source in [
            "SUM(Jan:Mar!A:B [@Status])",
            "SUM(Jan:Mar!1:2 [@Status])",
            "SUM(Jan:Mar!A:B (1:2))",
            "SUM(Jan:Mar!1:2 (A:B))",
            "SUM(Jan:Mar!A:B:C)",
            "SUM(Jan:Mar!1:2:3)",
            "SUM(Jan:Mar!A1 (A1:B2))",
            "SUM('Jan:Mar'!A:B:C)",
            "SUM(Jan:Mar!A1:B2 (1:2))+Feb!A1",
        ] {
            set_formula(&mut workbook, owner, "B1", source);
            for endpoint in [0, 2] {
                for op in [
                    Op::RemoveSheet { index: endpoint },
                    Op::RenameSheet {
                        sheet: SheetId(endpoint as u32),
                        name: "New Month".into(),
                    },
                    insert_rows(endpoint as u32, 0, 1),
                ] {
                    assert_sheet_edit_refuses(&workbook, &op, &error, source);
                }
            }
            assert_sheet_edit_refuses(&workbook, &insert_rows(1, 0, 1), &error, source);
            assert_sheet_edit_refuses(&workbook, &Op::RemoveSheet { index: 1 }, &error, source);
        }
    }

    #[test]
    fn ambiguous_sheet_span_bindings_refuse_edits_in_names_charts_and_rules() {
        let source = "SUM(Jan:Mar!A:B:C)";
        let report = SheetId(3);
        let mut workbook = wb(&["Jan", "Feb", "Mar", "Report"]);
        let ops = [
            Op::RemoveSheet { index: 2 },
            Op::RenameSheet {
                sheet: SheetId(2),
                name: "New March".into(),
            },
            insert_rows(1, 0, 1),
        ];
        workbook.defined_names = vec![defined("Span", source)];
        for op in &ops {
            let error = OpError::DefinedNameNotRewritable {
                name: "Span".into(),
            };
            assert_sheet_edit_refuses(&workbook, op, &error, source);
        }
        workbook.defined_names.clear();
        charted(&mut workbook, report, &[source]);
        for op in &ops {
            let error = OpError::ChartRefNotRewritable {
                part: "xl/charts/chart1.xml".into(),
            };
            assert_sheet_edit_refuses(&workbook, op, &error, source);
        }
        workbook.sheets[3].charts.clear();
        workbook.sheets[3]
            .conditional_formats
            .push(ConditionalFormatting {
                sqref: "A1".into(),
                rules: vec![ConditionalFormatRule {
                    id: "rule".into(),
                    rule_type: "expression".into(),
                    operator: None,
                    formulas: vec![source.into()],
                    priority: 1,
                    dxf_id: None,
                    format: Default::default(),
                    raw_xml: Some(format!("<cfRule><formula>{source}</formula></cfRule>")),
                }],
            });
        for op in &ops {
            let error = unrewritable_rule(
                &workbook.sheets[3].conditional_formats[0].rules[0],
                "Report",
            );
            assert_sheet_edit_refuses(&workbook, op, &error, source);
        }
    }

    #[test]
    fn rename_quotes_cell_like_names_and_matches_unicode_sources() {
        assert_eq!(rename_formula_sheet("S!A1", "S", "A1"), "'A1'!A1");
        assert_eq!(
            rename_formula_sheet("École1!A1", "École1", "Classe"),
            "Classe!A1"
        );
        assert_eq!(
            rename_formula_sheet("Sheet😀!A1", "Sheet😀", "Renamed"),
            "Renamed!A1"
        );
        assert_eq!(rename_formula_sheet("S!A1", "S", "R4C"), "'R4C'!A1");
    }

    fn defined(name: &str, formula: &str) -> DefinedName {
        DefinedName {
            name: name.to_owned(),
            formula: formula.to_owned(),
            local_sheet: None,
            hidden: false,
        }
    }

    /// Row and column ops rewrote cell formulas but left defined names on
    /// their pre-edit addresses.
    #[test]
    fn row_insertion_shifts_defined_name_formulas() {
        let mut workbook = wb(&["Data", "Other"]);
        workbook.defined_names = vec![
            defined("Global", "Data!$A$5"),
            defined("Elsewhere", "Other!$A$5"),
            DefinedName {
                local_sheet: Some(SheetId(0)),
                ..defined("Scoped", "$A$5")
            },
        ];
        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 2,
        };

        let inverse = remap_defined_names(&mut workbook, &op).unwrap();
        assert_eq!(workbook.defined_names[0].formula, "Data!$A$7");
        assert_eq!(workbook.defined_names[1].formula, "Other!$A$5");
        assert_eq!(workbook.defined_names[2].formula, "$A$7");
        assert!(matches!(inverse, Some(Op::SetDefinedNames { .. })));
    }

    #[test]
    fn one_sheet_workbook_names_rewrite_unqualified_references() {
        let mut workbook = wb(&["Data"]);
        workbook.defined_names = vec![defined("Input", "$A$5"), defined("Qualified", "=Data!$B$5")];

        remap_defined_names(&mut workbook, &insert_rows(0, 0, 2)).unwrap();

        assert_eq!(workbook.defined_names[0].formula, "$A$7");
        assert_eq!(workbook.defined_names[1].formula, "=Data!$B$7");
    }

    #[test]
    fn row_deletion_collapses_defined_names_onto_ref_errors() {
        let mut workbook = wb(&["Data"]);
        workbook.defined_names = vec![defined("Global", "Data!$A$1")];
        let op = Op::DeleteRows {
            sheet: SheetId(0),
            at: 0,
            count: 1,
        };

        remap_defined_names(&mut workbook, &op).unwrap();
        assert_eq!(workbook.defined_names[0].formula, "#REF!");
    }

    fn insert_rows(sheet: u32, at: u32, count: u32) -> Op {
        Op::InsertRows {
            sheet: SheetId(sheet),
            at,
            count,
        }
    }

    fn delete_rows(sheet: u32, at: u32, count: u32) -> Op {
        Op::DeleteRows {
            sheet: SheetId(sheet),
            at,
            count,
        }
    }

    /// Print areas are unions, which the formula parser has no expression for,
    /// so structural edits used to leave every one of them stale.
    #[test]
    fn row_insertion_shifts_every_branch_of_a_print_area_union() {
        let mut workbook = wb(&["Data", "Other"]);
        workbook.defined_names = vec![
            defined("_xlnm.Print_Area", "Data!$A$1:$D$20,Data!$F$1:$G$20"),
            defined("Mixed", "Data!$A$1:$D$20,Other!$A$1:$B$2"),
        ];

        let inverse = remap_defined_names(&mut workbook, &insert_rows(0, 0, 2)).unwrap();
        assert!(matches!(inverse, Some(Op::SetDefinedNames { .. })));
        assert_eq!(
            workbook.defined_names[0].formula,
            "Data!$A$3:$D$22,Data!$F$3:$G$22"
        );
        assert_eq!(
            workbook.defined_names[1].formula,
            "Data!$A$3:$D$22,Other!$A$1:$B$2"
        );
    }

    #[test]
    fn whole_axis_print_titles_follow_only_their_own_axis() {
        let mut workbook = wb(&["Data"]);
        workbook.defined_names = vec![
            defined("_xlnm.Print_Titles", "Data!$1:$2,Data!$A:$B"),
            defined("Relative", "Data!1:2"),
        ];

        remap_defined_names(&mut workbook, &insert_rows(0, 0, 3)).unwrap();
        assert_eq!(workbook.defined_names[0].formula, "Data!$4:$5,Data!$A:$B");
        assert_eq!(workbook.defined_names[1].formula, "Data!4:5");

        remap_defined_names(
            &mut workbook,
            &Op::InsertCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(workbook.defined_names[0].formula, "Data!$4:$5,Data!$B:$C");
    }

    #[test]
    fn deleting_a_whole_axis_reference_collapses_it_to_a_ref_error() {
        let mut workbook = wb(&["Data"]);
        workbook.defined_names = vec![
            defined("Titles", "Data!$1:$2,Data!$A:$A"),
            defined("Clipped", "Data!$1:$9"),
        ];

        remap_defined_names(
            &mut workbook,
            &Op::DeleteRows {
                sheet: SheetId(0),
                at: 0,
                count: 2,
            },
        )
        .unwrap();
        assert_eq!(workbook.defined_names[0].formula, "#REF!,Data!$A:$A");
        assert_eq!(workbook.defined_names[1].formula, "Data!$1:$7");
    }

    #[test]
    fn scoped_whole_axis_titles_resolve_against_their_own_sheet() {
        let mut workbook = wb(&["Data", "Other"]);
        workbook.defined_names = vec![
            DefinedName {
                local_sheet: Some(SheetId(0)),
                ..defined("_xlnm.Print_Titles", "$1:$1")
            },
            DefinedName {
                local_sheet: Some(SheetId(1)),
                ..defined("_xlnm.Print_Titles", "$1:$1")
            },
        ];

        remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap();
        assert_eq!(workbook.defined_names[0].formula, "$2:$2");
        assert_eq!(workbook.defined_names[1].formula, "$1:$1");
    }

    #[test]
    fn quoted_and_three_dimensional_union_branches_keep_their_qualifiers() {
        let mut workbook = wb(&["My Data", "Other"]);
        workbook.defined_names = vec![defined("Area", "'My Data'!$A$1:$B$2,'My Data'!$1:$1")];

        remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap();
        assert_eq!(
            workbook.defined_names[0].formula,
            "'My Data'!$A$2:$B$3,'My Data'!$2:$2"
        );
    }

    /// The ops a name is tested against: one far below any data, then a shift
    /// and a collapse on each axis.
    fn grid_edits() -> [Op; 5] {
        [
            Op::InsertRows {
                sheet: SheetId(0),
                at: 9998,
                count: 1,
            },
            insert_rows(0, 0, 1),
            Op::DeleteRows {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
            Op::InsertCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
            Op::DeleteCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        ]
    }

    /// Whole-axis references nested in a call, the range operator applied to a
    /// call, and the `Sheet!#REF!` Excel leaves behind after a deletion all
    /// defeat the formula parser, so every edit on a sheet an ordinary name
    /// mentioned used to be refused — even one far below the last row.
    #[test]
    fn names_the_formula_parser_cannot_read_move_with_the_grid() {
        let cases: &[(&str, [&str; 5])] = &[
            (
                "SUM(Data!$A:$A)",
                [
                    "SUM(Data!$A:$A)",
                    "SUM(Data!$A:$A)",
                    "SUM(Data!$A:$A)",
                    "SUM(Data!$B:$B)",
                    "SUM(#REF!)",
                ],
            ),
            (
                "COUNTA(Data!$A:$A)",
                [
                    "COUNTA(Data!$A:$A)",
                    "COUNTA(Data!$A:$A)",
                    "COUNTA(Data!$A:$A)",
                    "COUNTA(Data!$B:$B)",
                    "COUNTA(#REF!)",
                ],
            ),
            (
                "SUM(Data!$1:$1)",
                [
                    "SUM(Data!$1:$1)",
                    "SUM(Data!$2:$2)",
                    "SUM(#REF!)",
                    "SUM(Data!$1:$1)",
                    "SUM(Data!$1:$1)",
                ],
            ),
            (
                "Data!$A$1:INDEX(Data!$A:$A,COUNTA(Data!$A:$A))",
                [
                    "Data!$A$1:INDEX(Data!$A:$A,COUNTA(Data!$A:$A))",
                    "Data!$A$2:INDEX(Data!$A:$A,COUNTA(Data!$A:$A))",
                    "#REF!:INDEX(Data!$A:$A,COUNTA(Data!$A:$A))",
                    "Data!$B$1:INDEX(Data!$B:$B,COUNTA(Data!$B:$B))",
                    "#REF!:INDEX(#REF!,COUNTA(#REF!))",
                ],
            ),
            (
                "Data!$A$1:INDEX(Data!$A$1:$A$100,COUNTA(Data!$A$1:$A$100))",
                [
                    "Data!$A$1:INDEX(Data!$A$1:$A$100,COUNTA(Data!$A$1:$A$100))",
                    "Data!$A$2:INDEX(Data!$A$2:$A$101,COUNTA(Data!$A$2:$A$101))",
                    "#REF!:INDEX(Data!$A$1:$A$99,COUNTA(Data!$A$1:$A$99))",
                    "Data!$B$1:INDEX(Data!$B$1:$B$100,COUNTA(Data!$B$1:$B$100))",
                    "#REF!:INDEX(#REF!,COUNTA(#REF!))",
                ],
            ),
            (
                "OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),1)",
                [
                    "OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),1)",
                    "OFFSET(Data!$A$2,0,0,COUNTA(Data!$A:$A),1)",
                    "OFFSET(#REF!,0,0,COUNTA(Data!$A:$A),1)",
                    "OFFSET(Data!$B$1,0,0,COUNTA(Data!$B:$B),1)",
                    "OFFSET(#REF!,0,0,COUNTA(#REF!),1)",
                ],
            ),
            (
                "Data!#REF!",
                [
                    "Data!#REF!",
                    "Data!#REF!",
                    "Data!#REF!",
                    "Data!#REF!",
                    "Data!#REF!",
                ],
            ),
        ];

        for (formula, expected) in cases {
            for (op, expected) in grid_edits().iter().zip(expected) {
                let mut workbook = wb(&["Data", "Other"]);
                workbook.defined_names = vec![defined("Region", formula)];
                remap_defined_names(&mut workbook, op)
                    .unwrap_or_else(|error| panic!("{formula} under {op:?}: {error:?}"));
                assert_eq!(
                    workbook.defined_names[0].formula, *expected,
                    "{formula} under {op:?}"
                );
            }
        }
    }

    /// A name scoped to the edited sheet resolves its unqualified references
    /// against that sheet, so the same forms move without a qualifier.
    #[test]
    fn scoped_names_the_formula_parser_cannot_read_move_with_the_grid() {
        let cases: &[(&str, [&str; 5])] = &[
            (
                "SUM($A:$A)",
                [
                    "SUM($A:$A)",
                    "SUM($A:$A)",
                    "SUM($A:$A)",
                    "SUM($B:$B)",
                    "SUM(#REF!)",
                ],
            ),
            (
                "OFFSET($A$1,0,0,COUNTA($A:$A),1)",
                [
                    "OFFSET($A$1,0,0,COUNTA($A:$A),1)",
                    "OFFSET($A$2,0,0,COUNTA($A:$A),1)",
                    "OFFSET(#REF!,0,0,COUNTA($A:$A),1)",
                    "OFFSET($B$1,0,0,COUNTA($B:$B),1)",
                    "OFFSET(#REF!,0,0,COUNTA(#REF!),1)",
                ],
            ),
        ];

        for (formula, expected) in cases {
            for (op, expected) in grid_edits().iter().zip(expected) {
                let mut workbook = wb(&["Data", "Other"]);
                workbook.defined_names = vec![DefinedName {
                    local_sheet: Some(SheetId(0)),
                    ..defined("Region", formula)
                }];
                remap_defined_names(&mut workbook, op)
                    .unwrap_or_else(|error| panic!("{formula} under {op:?}: {error:?}"));
                assert_eq!(
                    workbook.defined_names[0].formula, *expected,
                    "{formula} under {op:?}"
                );
            }
        }
    }

    /// A workbook name's unqualified reference binds to whichever sheet is
    /// active, so it is still refused — unless the edit provably cannot reach
    /// it, as a row edit cannot reach a whole-column reference.
    #[test]
    fn unqualified_references_in_workbook_names_stay_ambiguous() {
        let mut workbook = wb(&["Data", "Other"]);
        workbook.defined_names = vec![defined("Region", "OFFSET($A$1,0,0,COUNTA($A:$A),1)")];
        let error = remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap_err();
        assert!(matches!(error, OpError::DefinedNameNotRewritable { .. }));

        let mut workbook = wb(&["Data", "Other"]);
        workbook.defined_names = vec![defined("Region", "SUM($A:$A)")];
        remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap();
        assert_eq!(workbook.defined_names[0].formula, "SUM($A:$A)");
    }

    /// Three-dimensional qualifiers and a leading `=` survive the token
    /// rewriter, and a reference aimed off the edited sheet stays put.
    #[test]
    fn the_token_rewriter_keeps_qualifiers_it_did_not_move() {
        let mut workbook = wb(&["Data", "Other"]);
        workbook.defined_names = vec![
            defined("Span", "SUM(Data:Other!$A$1,'My Book'!$A:$A)"),
            defined("Prefixed", "=SUM(Data!$1:$1)"),
            defined("Elsewhere", "SUM(Other!$1:$1)"),
        ];

        remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap();
        assert_eq!(
            workbook.defined_names[0].formula,
            "SUM(Data:Other!$A$2,'My Book'!$A:$A)"
        );
        assert_eq!(workbook.defined_names[1].formula, "=SUM(Data!$2:$2)");
        assert_eq!(workbook.defined_names[2].formula, "SUM(Other!$1:$1)");
    }

    /// Sheet names are not ascii, and neither are the qualifiers built from
    /// them; the scanner reads whole characters rather than bytes.
    #[test]
    fn the_token_rewriter_reads_non_ascii_sheet_names() {
        let mut workbook = wb(&["Ümsätze", "Other"]);
        workbook.defined_names = vec![
            defined(
                "Rows",
                "Ümsätze!$A$1:INDEX(Ümsätze!$A:$A,COUNTA(Ümsätze!$A:$A))",
            ),
            defined("Odd", "SUM(Ümsätze!$1:$1,\"Ümsätze!$A$1\")"),
        ];

        remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap();
        assert_eq!(
            workbook.defined_names[0].formula,
            "Ümsätze!$A$2:INDEX(Ümsätze!$A:$A,COUNTA(Ümsätze!$A:$A))"
        );
        assert_eq!(
            workbook.defined_names[1].formula,
            "SUM(Ümsätze!$2:$2,\"Ümsätze!$A$1\")"
        );
    }

    /// `A1#` names whatever the formula at `A1` spilled and `@` picks one cell
    /// out of the reference behind it. Both bind to the address they carry, so
    /// both move with it and both fall with it.
    #[test]
    fn an_array_constant_names_no_cell_and_passes_through() {
        let cases: &[(&str, &str)] = &[
            ("SUM(Data!$1:$1,{1,2,3})", "SUM(Data!$2:$2,{1,2,3})"),
            (
                "SUM(Data!$1:$1,{\"a\";\"b\"})",
                "SUM(Data!$2:$2,{\"a\";\"b\"})",
            ),
            (
                "SUM(Data!$1:$1,{TRUE,FALSE})",
                "SUM(Data!$2:$2,{TRUE,FALSE})",
            ),
        ];
        for (formula, expected) in cases {
            let mut workbook = wb(&["Data", "Other"]);
            workbook.defined_names = vec![defined("Arr", formula)];
            remap_defined_names(&mut workbook, &insert_rows(0, 0, 1))
                .unwrap_or_else(|error| panic!("{formula}: {error:?}"));
            assert_eq!(workbook.defined_names[0].formula, *expected, "{formula}");
        }

        let mut workbook = wb(&["Data"]);
        workbook.defined_names = vec![defined("Open", "SUM(Data!$1:$1,{1,2")];
        assert!(remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).is_err());

        // at the top of a name, where the union splitter sees the commas first
        for (formula, expected) in [
            ("Data!$1:$1+{1,2,3}", "Data!$2:$2+{1,2,3}"),
            ("{1,2,3}", "{1,2,3}"),
        ] {
            let mut workbook = wb(&["Data", "Other"]);
            workbook.defined_names = vec![defined("Arr", formula)];
            remap_defined_names(&mut workbook, &insert_rows(0, 0, 1))
                .unwrap_or_else(|error| panic!("{formula}: {error:?}"));
            assert_eq!(workbook.defined_names[0].formula, expected, "{formula}");
        }
        let mut workbook = wb(&["Data"]);
        workbook.defined_names = vec![DefinedName {
            local_sheet: Some(SheetId(0)),
            ..defined("Consts", "{1,2,3}")
        }];
        remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap();
        assert_eq!(workbook.defined_names[0].formula, "{1,2,3}");
    }

    #[test]
    fn spill_and_implicit_intersection_references_move_with_their_address() {
        let cases: &[(&str, &str, &str)] = &[
            ("SUM(Data!A1#)", "SUM(Data!A2#)", "SUM(#REF!)"),
            ("SUM(@Data!A1)", "SUM(@Data!A2)", "SUM(#REF!)"),
            ("SUM(@Data!A1#)", "SUM(@Data!A2#)", "SUM(#REF!)"),
            (
                "SUM(Data!A1#,Data!$A:$A)",
                "SUM(Data!A2#,Data!$A:$A)",
                "SUM(#REF!,Data!$A:$A)",
            ),
        ];

        for (formula, inserted, deleted) in cases {
            for (op, expected) in [
                (insert_rows(0, 0, 1), inserted),
                (delete_rows(0, 0, 1), deleted),
            ] {
                let mut workbook = wb(&["Data", "Other"]);
                workbook.defined_names = vec![defined("Spilled", formula)];
                remap_defined_names(&mut workbook, &op)
                    .unwrap_or_else(|error| panic!("{formula} under {op:?}: {error:?}"));
                assert_eq!(
                    workbook.defined_names[0].formula, *expected,
                    "{formula} under {op:?}"
                );
            }
        }
    }

    /// Excel allows whitespace around the range operator and a qualifier on
    /// either half of a range. The endpoints are one span however it is
    /// written: clipped apart, a deletion that takes the near one strands it
    /// on `#REF!` while the far one still names a live cell.
    #[test]
    fn a_range_written_around_whitespace_clips_as_one_span() {
        let cases: &[(&str, &str, &str)] = &[
            ("Data!A1:B2", "Data!A2:B3", "Data!A1:B1"),
            ("Data!A1: Data!B2", "Data!A2: Data!B3", "Data!A1: Data!B1"),
            ("Data!A1 : B2", "Data!A2 : B3", "Data!A1 : B1"),
            (
                "SUM(Data!$A:$A,Data!A1: Data!B2)",
                "SUM(Data!$A:$A,Data!A2: Data!B3)",
                "SUM(Data!$A:$A,Data!A1: Data!B1)",
            ),
        ];

        for (formula, inserted, deleted) in cases {
            for (op, expected) in [
                (insert_rows(0, 0, 1), inserted),
                (delete_rows(0, 0, 1), deleted),
            ] {
                let mut workbook = wb(&["Data", "Other"]);
                workbook.defined_names = vec![defined("Region", formula)];
                remap_defined_names(&mut workbook, &op)
                    .unwrap_or_else(|error| panic!("{formula} under {op:?}: {error:?}"));
                assert_eq!(
                    workbook.defined_names[0].formula, *expected,
                    "{formula} under {op:?}"
                );
            }
        }
    }

    /// `Data:Other!` covers the sheets between its endpoints, but a cell
    /// address in front of a qualifier is the range operator instead, and the
    /// unqualified half of a workbook name binds to whichever sheet is active.
    /// Reading the one as the other moved neither.
    #[test]
    fn a_range_endpoint_in_front_of_a_qualifier_is_not_a_sheet_span() {
        for scope in [None, Some(SheetId(0))] {
            let mut workbook = wb(&["Data", "Sheet2"]);
            workbook.defined_names = vec![DefinedName {
                local_sheet: scope,
                ..defined("Mixed", "A1:Sheet2!B2")
            }];
            let original = workbook.clone();

            let error = remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap_err();
            assert!(
                matches!(error, OpError::DefinedNameNotRewritable { .. }),
                "{scope:?}"
            );
            assert_eq!(workbook.defined_names, original.defined_names);
        }

        let mut workbook = wb(&["Data", "Other"]);
        workbook.defined_names = vec![defined("Span", "SUM(Data:Other!$A$1)")];
        remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap();
        assert_eq!(workbook.defined_names[0].formula, "SUM(Data:Other!$A$2)");
    }

    /// Excel refuses to define a name shaped like a cell address, but it reads
    /// workbooks that hold one, and a reference to such a name cannot be told
    /// apart from the cell it is named after. Moving one as the other renamed
    /// the reference and broke it, so the edit refuses.
    #[test]
    fn a_reference_to_a_cell_shaped_defined_name_refuses_the_edit() {
        for name in ["TAX2024", "A1", "ABC1"] {
            let mut workbook = wb(&["Data"]);
            workbook.defined_names = vec![
                defined(name, "Data!$Z$9"),
                defined("UseTax", &format!("SUM({name},Data!$A:$A)")),
            ];
            let original = workbook.clone();

            let error = remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap_err();
            assert!(
                matches!(error, OpError::DefinedNameNotRewritable { .. }),
                "{name}"
            );
            assert_eq!(workbook.defined_names, original.defined_names);
        }

        let mut workbook = wb(&["Data"]);
        workbook.defined_names = vec![
            defined("TaxRate", "Data!$Z$9"),
            defined("UseTax", "SUM(TaxRate,Data!$A:$A)"),
        ];
        remap_defined_names(
            &mut workbook,
            &Op::InsertCols {
                sheet: SheetId(0),
                at: 0,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(workbook.defined_names[0].formula, "Data!$AA$9");
        assert_eq!(workbook.defined_names[1].formula, "SUM(TaxRate,Data!$B:$B)");
    }

    /// A rewrite the scanner guessed at is saved into the workbook, so a token
    /// it cannot place refuses the edit rather than being waved through.
    /// Numbers, booleans and strings name nothing on the grid and still pass.
    #[test]
    fn a_token_the_scanner_cannot_place_refuses_the_edit() {
        let mut workbook = wb(&["Data", "Other"]);
        workbook.defined_names = vec![defined("Mystery", "SUM(Data!$A:$A,Whatever)")];
        let original = workbook.clone();

        let error = remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap_err();
        assert!(matches!(error, OpError::DefinedNameNotRewritable { .. }));
        assert_eq!(workbook.defined_names, original.defined_names);

        let mut workbook = wb(&["Data", "Other"]);
        workbook.defined_names = vec![defined(
            "Literals",
            "SUM(Data!$1:$1,0.5,TRUE,\"Data!$A$1\",Data!#REF!)",
        )];
        remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap();
        assert_eq!(
            workbook.defined_names[0].formula,
            "SUM(Data!$2:$2,0.5,TRUE,\"Data!$A$1\",Data!#REF!)"
        );
    }

    /// A chart aimed at a dynamic range holds the same idiom a defined name
    /// does, and every edit on the sheet it charts used to be refused for it.
    #[test]
    fn chart_references_the_formula_parser_cannot_read_move_with_the_grid() {
        let mut w = wb(&["Data", "Other"]);
        charted(
            &mut w,
            SheetId(1),
            &[
                "OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),1)",
                "(Data!$A$1:INDEX(Data!$A:$A,COUNTA(Data!$A:$A)))",
                "SUM(Other!$1:$1)",
            ],
        );

        remap_charts(&mut w, &insert_rows(0, 0, 1)).unwrap();
        assert_eq!(
            chart_formulas(&w, SheetId(1)),
            [
                "OFFSET(Data!$A$2,0,0,COUNTA(Data!$A:$A),1)",
                "(Data!$A$2:INDEX(Data!$A:$A,COUNTA(Data!$A:$A)))",
                "SUM(Other!$1:$1)",
            ]
        );
    }

    /// A structured table reference is beyond the rewriter, and leaving it
    /// stale is exactly what this refuses to do.
    #[test]
    fn unrewritable_name_aimed_at_the_edited_sheet_refuses_the_edit() {
        let mut workbook = wb(&["Data", "Other"]);
        workbook.defined_names = vec![defined("Dynamic", "SUM(Data!Table1[Amount])")];
        let original = workbook.clone();

        let error = remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap_err();
        assert_eq!(
            error,
            OpError::DefinedNameNotRewritable {
                name: "Dynamic".to_owned()
            }
        );
        assert_eq!(workbook, original);
    }

    #[test]
    fn unrewritable_name_aimed_elsewhere_survives_the_edit_untouched() {
        let mut workbook = wb(&["Data", "Other"]);
        workbook.defined_names = vec![
            defined("Dynamic", "OFFSET(Other!$A$1,0,0,COUNTA(Other!$A:$A),1)"),
            defined("External", "[Book.xlsx]Data!$A:$A"),
            defined("Structured", "Table1[#All]"),
            defined("Area", "Data!$A$1"),
        ];

        remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap();
        assert_eq!(
            workbook.defined_names[0].formula,
            "OFFSET(Other!$A$1,0,0,COUNTA(Other!$A:$A),1)"
        );
        assert_eq!(workbook.defined_names[1].formula, "[Book.xlsx]Data!$A:$A");
        assert_eq!(workbook.defined_names[2].formula, "Table1[#All]");
        assert_eq!(workbook.defined_names[3].formula, "Data!$A$2");
    }

    #[test]
    fn a_scoped_name_the_rewriter_cannot_read_refuses_the_edit() {
        let mut workbook = wb(&["Data"]);
        workbook.defined_names = vec![DefinedName {
            local_sheet: Some(SheetId(0)),
            ..defined("Dynamic", "SUM(Table1[Amount])")
        }];

        let error = remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap_err();
        assert!(matches!(error, OpError::DefinedNameNotRewritable { .. }));
    }

    #[test]
    fn structural_edits_refuse_defined_names_that_would_outgrow_the_length_cap() {
        let mut workbook = wb(&["Data"]);
        let padding = "+0".repeat((MAX_FORMULA_BYTES - 10) / 2);
        workbook.defined_names = vec![defined("Long", &format!("Data!$A$99{padding}"))];
        let original = workbook.clone();

        let error = remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap_err();
        assert!(matches!(error, OpError::DefinedNameNotRewritable { .. }));
        assert_eq!(workbook, original);
    }

    #[test]
    fn rename_keeps_defined_names_whose_formulas_only_call_same_named_functions() {
        for function in ["SUM", "MAX", "IF", "DATE", "TEXT", "PI"] {
            let formula = format!("{function}(A1:A10)");
            let mut workbook = wb(&[function, "Other"]);
            workbook.defined_names = vec![defined("Total", &formula)];
            let renamed = rename_defined_names(&workbook, function, "Renamed").unwrap();
            assert!(renamed.is_none(), "{formula} must not change");
        }
    }

    #[test]
    fn rename_rewrites_qualified_defined_names_and_drops_bare_mentions() {
        let mut workbook = wb(&["Data", "Other"]);
        workbook.defined_names = vec![
            defined("Qualified", "Data!$A$1"),
            defined("ThreeD", "Data:Other!$A$1"),
            defined("Bare", "Data"),
            defined("Literal", "42"),
            defined("Quoted", "\"Data\""),
            defined("External", "[Book.xlsx]Data!$A$1"),
        ];
        let renamed = rename_defined_names(&workbook, "Data", "New Data").unwrap();
        assert_eq!(
            renamed.unwrap(),
            vec![
                defined("Qualified", "'New Data'!$A$1"),
                defined("ThreeD", "'New Data:Other'!$A$1"),
                defined("Literal", "42"),
                defined("Quoted", "\"Data\""),
                defined("External", "[Book.xlsx]Data!$A$1"),
            ]
        );
    }

    #[test]
    fn rename_drops_defined_names_that_mix_a_bare_mention_with_a_reference() {
        let mut workbook = wb(&["SUM", "Other"]);
        workbook.defined_names = vec![
            defined("Callable", "SUM(SUM!A1:A10)"),
            defined("Shadowed", "SUM(SUM,1)"),
            defined("Spaced", "SUM (A1:A10)"),
            defined("Quoted", "'SUM'"),
        ];
        let renamed = rename_defined_names(&workbook, "SUM", "Renamed").unwrap();
        assert_eq!(
            renamed.unwrap(),
            vec![
                defined("Callable", "SUM(Renamed!A1:A10)"),
                defined("Spaced", "SUM (A1:A10)"),
            ]
        );
    }

    #[test]
    fn structural_remap_preserves_a_trailing_location_fragment() {
        let mut wb = Workbook::default();
        wb.sheets.push(Sheet::new("Data"));
        wb.sheets.push(Sheet::new("Target"));
        let link = |location: &str| Hyperlink {
            range: CellRange::parse_a1("A1").unwrap(),
            external_target: None,
            location: Some(location.into()),
            tooltip: None,
            display: None,
        };
        wb.sheets[0]
            .hyperlinks
            .extend([link("#Target!A3#2"), link("Target!A3!B4")]);

        remap_hyperlink_locations(
            &mut wb,
            &Op::InsertRows {
                sheet: SheetId(1),
                at: 1,
                count: 2,
            },
        );

        assert_eq!(
            wb.sheets[0].hyperlinks[0].location.as_deref(),
            Some("#Target!A5#2")
        );
        assert_eq!(
            wb.sheets[0].hyperlinks[1].location.as_deref(),
            Some("Target!A3!B4")
        );
    }

    #[test]
    fn rename_hyperlink_location_preserves_the_leading_hash() {
        assert_eq!(
            rename_hyperlink_location("#Target!A1", "Target", "Renamed").unwrap(),
            "#Renamed!A1"
        );
        assert_eq!(
            rename_hyperlink_location("Target!A1", "Target", "Renamed").unwrap(),
            "Renamed!A1"
        );
        assert_eq!(
            rename_hyperlink_location("#Target!A1", "Target", "New Target").unwrap(),
            "#'New Target'!A1"
        );
    }

    #[test]
    fn rename_hyperlink_location_handles_quoted_and_bare_targets() {
        assert_eq!(
            rename_hyperlink_location("#'My Sheet'!A1", "My Sheet", "Renamed").unwrap(),
            "#Renamed!A1"
        );
        assert_eq!(
            rename_hyperlink_location("'My Sheet'!A1:B2", "My Sheet", "Renamed").unwrap(),
            "Renamed!A1:B2"
        );
        assert_eq!(
            rename_hyperlink_location("#'It''s Data'!A1", "It's Data", "Renamed").unwrap(),
            "#Renamed!A1"
        );
        for location in ["#MyRange", "MyRange", "#Target", "#A1"] {
            assert_eq!(
                rename_hyperlink_location(location, "Target", "Renamed").unwrap(),
                location
            );
        }
        assert_eq!(
            rename_hyperlink_location("#Target!A1#2", "Target", "Renamed").unwrap(),
            "#Renamed!A1#2"
        );
    }

    /// A rename refuses to write a sheet name xml 1.0 cannot carry into text
    /// that references the sheet; text that does not reference it is left
    /// alone and never refused.
    #[test]
    fn a_rename_refuses_to_write_a_name_xml_cannot_carry_into_a_reference() {
        for new_name in ["Data\u{1}", "Data\u{0}", "Data\u{fffe}"] {
            assert_eq!(
                rename_hyperlink_location("#Target!A1", "Target", new_name),
                Err(OpError::SheetNameNotWritable)
            );
            assert_eq!(
                rename_hyperlink_location("#Other!A1", "Target", new_name),
                Ok("#Other!A1".to_owned())
            );
        }
        assert!(rename_hyperlink_location("#Target!A1", "Target", "Tab\tName").is_ok());
    }

    #[test]
    fn interval_clip_edges() {
        assert_eq!(clip_interval(0, 9, 2, 1), Some((0, 8)));
        assert_eq!(clip_interval(4, 6, 4, 3), None);
        assert_eq!(clip_interval(2, 9, 2, 4), Some((2, 5)));
        assert_eq!(clip_interval(0, 1, 5, 2), Some((0, 1)));
    }

    /// a filled-down column parses one text many times; every copy must come
    /// out rewritten alike.
    #[test]
    fn identical_formulas_in_many_cells_all_rewrite_alike() {
        let mut w = wb(&["Data", "Other"]);
        for row in 1..=40 {
            set_formula(&mut w, SheetId(0), &format!("C{row}"), "SUM($A$1:$A$10)+B5");
            set_formula(
                &mut w,
                SheetId(1),
                &format!("C{row}"),
                "Data!$A$1:$A$10+Data!B5",
            );
        }
        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 3,
        };
        let inv = remap_formulas(&mut w, &op).unwrap();
        for row in 1..=40 {
            assert_eq!(
                formula(&w, SheetId(0), &format!("C{row}")).as_deref(),
                Some("SUM($A$4:$A$13)+B8")
            );
            assert_eq!(
                formula(&w, SheetId(1), &format!("C{row}")).as_deref(),
                Some("Data!$A$4:$A$13+Data!B8")
            );
        }
        assert_eq!(inv.len(), 80);
        apply_inverse_removes(&mut w, &inv);
        for row in 1..=40 {
            assert_eq!(
                formula(&w, SheetId(0), &format!("C{row}")).as_deref(),
                Some("SUM($A$1:$A$10)+B5")
            );
            assert_eq!(
                formula(&w, SheetId(1), &format!("C{row}")).as_deref(),
                Some("Data!$A$1:$A$10+Data!B5")
            );
        }
    }

    fn apply_inverse_removes(w: &mut Workbook, inv: &[Op]) {
        let restores = inv
            .iter()
            .filter_map(|op| match op {
                Op::SetCell { sheet, at, cell } => Some((*sheet, *at, cell.clone())),
                _ => None,
            })
            .collect::<Vec<_>>();
        for (sheet, at, cell) in restores {
            w.sheet_mut(sheet)
                .expect("sheet exists during restore")
                .set_cell(at, cell.into());
        }
    }

    /// the parse cache is keyed on text alone; the rewrite must still resolve
    /// against each owning sheet, so identical text can diverge per sheet.
    #[test]
    fn duplicated_formula_text_is_rewritten_per_owning_sheet() {
        let mut w = wb(&["Data", "Other"]);
        set_formula(&mut w, SheetId(0), "B1", "A5*2");
        set_formula(&mut w, SheetId(0), "B9", "A5*2");
        set_formula(&mut w, SheetId(1), "B1", "A5*2");
        set_formula(&mut w, SheetId(1), "B9", "Data!A5*2");
        remap_formulas(
            &mut w,
            &Op::DeleteRows {
                sheet: SheetId(0),
                at: 4,
                count: 1,
            },
        )
        .unwrap();
        assert_eq!(formula(&w, SheetId(0), "B1").as_deref(), Some("#REF!*2"));
        assert_eq!(formula(&w, SheetId(0), "B9").as_deref(), Some("#REF!*2"));
        assert_eq!(formula(&w, SheetId(1), "B1").as_deref(), Some("A5*2"));
        assert_eq!(
            formula(&w, SheetId(1), "B9").as_deref(),
            Some("#REF!*2"),
            "a deleted qualified ref collapses like an unqualified one"
        );
    }

    /// more distinct formulas than the memo can hold: entries evicted before
    /// their cell is reached must still rewrite exactly as if parsed fresh.
    #[test]
    fn distinct_formulas_beyond_memo_cap_all_rewrite_correctly() {
        let count = PARSE_MEMO_CAP * 2 + 7;
        let mut w = wb(&["Sheet1"]);
        for row in 1..=count {
            set_formula(
                &mut w,
                SheetId(0),
                &format!("A{row}"),
                &format!("B{row}+{row}"),
            );
        }
        let op = Op::InsertRows {
            sheet: SheetId(0),
            at: 0,
            count: 3,
        };
        let inv = remap_formulas(&mut w, &op).unwrap();
        assert_eq!(inv.len(), count);
        for row in 1..=count {
            assert_eq!(
                formula(&w, SheetId(0), &format!("A{row}")).as_deref(),
                Some(format!("B{}+{row}", row + 3)).as_deref()
            );
        }
    }

    /// A chart reference over a 3-D span keeps its text when a sheet inside
    /// the span is deleted, quoting included, and narrows only when an
    /// endpoint is the deleted sheet.
    #[test]
    fn a_chart_span_keeps_its_text_when_a_sheet_inside_it_is_deleted() {
        let order = ["Jan", "Feb", "Mar"].map(str::to_owned).to_vec();
        assert_eq!(
            super::drop_removed_sheet("'Jan:Mar'!$A$1", "Feb", &order),
            None
        );
        assert_eq!(
            super::drop_removed_sheet("'Jan:Mar'!$A$1", "Jan", &order).as_deref(),
            Some("Feb:Mar!$A$1")
        );
    }
}
