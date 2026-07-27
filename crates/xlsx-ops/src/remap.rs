//! rewriting stored formulas on row/column insert/delete: refs shift, ranges
//! clip, wholly deleted refs collapse to `#REF!`. runs before cells shift.

use std::collections::HashMap;

use xlsx_calc::lexer::MAX_FORMULA_BYTES;
use xlsx_calc::parse_formula;
use xlsx_calc::parser::Expr;
use xlsx_model::addr::{MAX_COLS, MAX_ROWS, col_to_letters};
use xlsx_model::{
    AnchorCell, AnchorEditAs, CellRange, CellRef, ChartAnchor, DefinedName, ErrorValue, SheetId,
    Workbook,
};

use crate::apply::OpError;
use crate::apply::remap_ref;
use crate::op::{CellState, Op};

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
        | Op::DeleteCols { sheet, .. } => Some(sheet),
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
    for (i, sheet) in wb.sheets.iter().enumerate() {
        let owner = SheetId(i as u32);
        let matches = |ref_sheet: &Option<String>| resolves_to_target(ref_sheet, owner, &index);
        for (cell, c) in sheet.iter_cells() {
            let Some(src) = &c.formula else {
                continue;
            };
            if has_unresolvable_binding(src, &index) {
                return Err(OpError::FormulaNotRewritable { sheet: owner, cell });
            }
            let expr = parse_formula(src)
                .map_err(|_| OpError::FormulaNotRewritable { sheet: owner, cell })?;
            let mut changed = false;
            let new_expr = transform(&expr, op, &matches, &mut changed);
            if changed {
                let formula = new_expr.to_formula();
                parse_formula(&formula)
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
        if has_unresolvable_binding(&defined.formula, &sheets) {
            return Err(OpError::DefinedNameNotRewritable {
                name: defined.name.clone(),
            });
        }
        let rewrite = rewrite_defined_name(
            &defined.formula,
            op,
            &matches,
            defined.local_sheet.is_none() && !global_target,
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
                if has_unresolvable_binding(&reference.formula, &sheets) {
                    return Err(OpError::ChartRefNotRewritable {
                        part: chart.part.clone(),
                    });
                }
                match rewrite_chart_ref(&reference.formula, op, &matches) {
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
) -> DefinedNameRewrite {
    let trimmed = source.trim();
    let Some(inner) = paren_group(trimmed) else {
        return rewrite_defined_name(trimmed, op, matches_target, false);
    };
    match rewrite_defined_name(inner, op, matches_target, false) {
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
    let (axis, at, count, inserting) = match *op {
        Op::InsertRows { at, count, .. } => (Axis::Row, at, count, true),
        Op::DeleteRows { at, count, .. } => (Axis::Row, at, count, false),
        Op::InsertCols { at, count, .. } => (Axis::Col, at, count, true),
        Op::DeleteCols { at, count, .. } => (Axis::Col, at, count, false),
        _ => return Ok(None),
    };
    let limit = axis.limit();
    let shift = |index: u32| shift_anchor_index(index, at, count, inserting, limit);
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
            let moved_from = shift_anchor_cell(from, axis, &shift).ok_or(())?;
            let moved_to = match edit_as {
                AnchorEditAs::TwoCell => shift_anchor_cell(to, axis, &shift).ok_or(())?,
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
                reference.formula = rename_formula_sheet(&reference.formula, old_name, new_name);
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
    for qualifier in sheet_qualifiers(component) {
        if !qualifier.bound || qualifier.external {
            continue;
        }
        if !qualifier.is_span() {
            if sheet_names_equal(&qualifier.first, removed_name) {
                return ComponentRemoval::Collapsed;
            }
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

/// `Some(Some(text))` narrows the span onto the sheets that remain,
/// `Some(None)` means nothing it covered survives, `None` means the removal
/// does not touch it.
fn narrowed_span(
    qualifier: &Qualifier,
    removed_name: &str,
    order_before: &[String],
) -> Option<Option<String>> {
    let position = |name: &str| {
        order_before
            .iter()
            .position(|sheet| sheet_names_equal(sheet, name))
    };
    let (Some(start), Some(end)) = (position(&qualifier.first), position(&qualifier.last)) else {
        return (sheet_names_equal(&qualifier.first, removed_name)
            || sheet_names_equal(&qualifier.last, removed_name))
        .then_some(None);
    };
    let covered = &order_before[start.min(end)..=start.max(end)];
    if !covered
        .iter()
        .any(|sheet| sheet_names_equal(sheet, removed_name))
    {
        return None;
    }
    let remaining = covered
        .iter()
        .filter(|sheet| !sheet_names_equal(sheet, removed_name))
        .collect::<Vec<_>>();
    let (Some(first), Some(last)) = (remaining.first(), remaining.last()) else {
        return Some(None);
    };
    Some(Some(qualifier_token(first, last)))
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
) -> DefinedNameRewrite {
    let (prefix, source) = match source.strip_prefix('=') {
        Some(source) => ("=", source),
        None => ("", source),
    };
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
        let Ok(expr) = parse_formula(component) else {
            if global && mentions_unqualified_reference(component) {
                return DefinedNameRewrite::Ambiguous;
            }
            return DefinedNameRewrite::Unsupported;
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
        if parse_formula(&formula).is_err() {
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

fn contains_unqualified_reference(expr: &Expr) -> bool {
    match expr {
        Expr::Ref { sheet: None, .. } | Expr::Range { sheet: None, .. } => true,
        Expr::Unary { expr, .. } | Expr::Percent(expr) => contains_unqualified_reference(expr),
        Expr::Binary { lhs, rhs, .. } => {
            contains_unqualified_reference(lhs) || contains_unqualified_reference(rhs)
        }
        Expr::FuncCall { args, .. } => args.iter().any(contains_unqualified_reference),
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

/// A whole-row (`Sheet!$1:$5`) or whole-column (`Sheet!$A:$C`) reference. Print
/// titles are written this way and the formula lexer has no token for it.
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
        if !matches_target(&self.sheet) || !self.axis.is_edited_by(op) {
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

/// One sheet qualifier found in a formula. A 3-D qualifier names two
/// endpoints and covers every sheet between them in workbook order, whether it
/// is written `Jan:Mar!` or, as Excel writes it, `'Jan:Mar'!`.
struct Qualifier {
    span: core::ops::Range<usize>,
    first: String,
    last: String,
    /// preceded by `]`, so it names a sheet in another workbook.
    external: bool,
    /// followed by `!`, so it really qualifies a reference.
    bound: bool,
    /// followed by `(`, so it is a function call rather than a sheet name.
    call: bool,
}

impl Qualifier {
    fn is_span(&self) -> bool {
        self.first != self.last
    }
}

/// Every sheet qualifier in `source`, in order, skipping strings and the
/// `[book]` prefixes of external references.
fn sheet_qualifiers(source: &str) -> Vec<Qualifier> {
    let bytes = source.as_bytes();
    let mut out = Vec::new();
    let mut index = 0;
    let mut bracket_depth = 0_u32;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            index = skip_string(source, index);
            continue;
        }
        if bytes[index] == b'[' {
            bracket_depth = bracket_depth.saturating_add(1);
            index += 1;
            continue;
        }
        if bytes[index] == b']' {
            bracket_depth = bracket_depth.saturating_sub(1);
            index += 1;
            continue;
        }
        if bracket_depth != 0 {
            index += next_char_len(source, index);
            continue;
        }
        let Some(first) = parse_sheet_token(source, index) else {
            index += next_char_len(source, index);
            continue;
        };
        let external = first.start > 0 && bytes[first.start - 1] == b']';
        if bytes.get(first.end) == Some(&b'!') {
            let (start, end) = quoted_endpoints(&first.name);
            out.push(Qualifier {
                span: first.start..first.end,
                first: start,
                last: end,
                external,
                bound: true,
                call: false,
            });
            index = first.end + 1;
            continue;
        }
        if bytes.get(first.end) == Some(&b':')
            && let Some(second) = parse_sheet_token(source, first.end + 1)
            && bytes.get(second.end) == Some(&b'!')
        {
            out.push(Qualifier {
                span: first.start..second.end,
                first: first.name,
                last: second.name,
                external,
                bound: true,
                call: false,
            });
            index = second.end + 1;
            continue;
        }
        let call = is_function_call(source, &first);
        index = first.end;
        out.push(Qualifier {
            span: first.start..first.end,
            first: first.name.clone(),
            last: first.name,
            external,
            bound: false,
            call,
        });
    }
    out
}

/// `'Jan:Mar'` names the span from `Jan` to `Mar`; a sheet name cannot itself
/// contain a colon, so the split is unambiguous.
fn quoted_endpoints(name: &str) -> (String, String) {
    match name.split_once(':') {
        Some((first, last)) => (first.to_owned(), last.to_owned()),
        None => (name.to_owned(), name.to_owned()),
    }
}

/// How a qualifier resolves against the workbook the edit targets.
struct SheetIndex<'a> {
    names: &'a HashMap<String, SheetId>,
    target: SheetId,
    target_name: &'a str,
}

impl SheetIndex<'_> {
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

/// Whether a qualifier that names the edited sheet cannot be resolved through
/// workbook order, so nothing can be said about what it covers and leaving it
/// alone would silently strand it.
fn has_unresolvable_binding(source: &str, index: &SheetIndex<'_>) -> bool {
    sheet_qualifiers(source).iter().any(|qualifier| {
        qualifier.bound
            && !qualifier.external
            && !(index.names.contains_key(&qualifier.first.to_lowercase())
                && index.names.contains_key(&qualifier.last.to_lowercase()))
            && index.covers(&qualifier.first, &qualifier.last)
    })
}

/// Whether an unrewritable formula names the edited sheet as a reference
/// qualifier, so leaving it untouched would strand it on the pre-edit
/// addresses.
fn mentions_sheet(source: &str, index: &SheetIndex<'_>) -> bool {
    sheet_qualifiers(source).iter().any(|qualifier| {
        qualifier.bound && !qualifier.external && index.covers(&qualifier.first, &qualifier.last)
    })
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
            let rewritten = rename_formula_sheet(source, old_name, new_name);
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

/// Splits a hyperlink location into its reference and any trailing `#` fragment.
fn split_location_reference(source: &str) -> Option<(Expr, &str)> {
    if let Ok(expr) = parse_formula(source) {
        return Some((expr, ""));
    }
    let (reference, _) = source.rsplit_once('#')?;
    let expr = parse_formula(reference).ok()?;
    Some((expr, &source[reference.len()..]))
}

/// rename the sheet inside an internal hyperlink location. locations are
/// commonly written `#Sheet!A1`; the leading `#` is not part of the reference
/// and must not reach the formula rewriter, but its presence is preserved.
pub(crate) fn rename_hyperlink_location(location: &str, old_name: &str, new_name: &str) -> String {
    match location.strip_prefix('#') {
        Some(reference) => format!("#{}", rename_formula_sheet(reference, old_name, new_name)),
        None => rename_formula_sheet(location, old_name, new_name),
    }
}

pub(crate) fn rename_formula_sheet(source: &str, old_name: &str, new_name: &str) -> String {
    scan_sheet_rename(source, old_name, new_name).formula
}

/// A formula rewritten for a sheet rename, plus whether it still names the old
/// sheet in a position the rewriter could not resolve.
struct SheetRename {
    formula: String,
    ambiguous: bool,
}

/// Rewrites qualified sheet references and detects ambiguous bare names.
fn scan_sheet_rename(source: &str, old_name: &str, new_name: &str) -> SheetRename {
    if old_name == new_name {
        return SheetRename {
            formula: source.to_string(),
            ambiguous: false,
        };
    }
    let mut replacements = Vec::new();
    let mut ambiguous = false;
    for qualifier in sheet_qualifiers(source) {
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

/// An unquoted token followed by `(` is a function call, never a sheet name.
fn is_function_call(source: &str, token: &ParsedSheetToken) -> bool {
    !source[token.start..].starts_with('\'') && source[token.end..].trim_start().starts_with('(')
}

/// Rewrites defined names through a sheet rename, dropping any whose formula
/// names the old sheet as a bare token, where the intent cannot be resolved.
/// Returns the inverse op when the list changed.
pub(crate) fn rename_defined_names(
    wb: &mut Workbook,
    old_name: &str,
    new_name: &str,
) -> Option<Op> {
    if old_name == new_name || wb.defined_names.is_empty() {
        return None;
    }
    let previous = wb.defined_names.clone();
    let mut rewritten: Vec<DefinedName> = Vec::with_capacity(previous.len());
    for defined in &previous {
        let renamed = scan_sheet_rename(&defined.formula, old_name, new_name);
        if renamed.ambiguous {
            continue;
        }
        let mut defined = defined.clone();
        defined.formula = renamed.formula;
        rewritten.push(defined);
    }
    if rewritten == previous {
        return None;
    }
    wb.defined_names = rewritten;
    Some(Op::SetDefinedNames {
        defined_names: previous,
    })
}

fn next_char_len(source: &str, index: usize) -> usize {
    source[index..]
        .chars()
        .next()
        .map(char::len_utf8)
        .unwrap_or(1)
}

struct ParsedSheetToken {
    start: usize,
    end: usize,
    name: String,
}

fn parse_sheet_token(source: &str, start: usize) -> Option<ParsedSheetToken> {
    let bytes = source.as_bytes();
    if bytes.get(start) == Some(&b'\'') {
        let mut index = start + 1;
        let mut name = String::new();
        while index < bytes.len() {
            if bytes[index] == b'\'' {
                if bytes.get(index + 1) == Some(&b'\'') {
                    name.push('\'');
                    index += 2;
                } else {
                    return Some(ParsedSheetToken {
                        start,
                        end: index + 1,
                        name,
                    });
                }
            } else {
                let character = source[index..].chars().next()?;
                name.push(character);
                index += character.len_utf8();
            }
        }
        return None;
    }
    let first = source[start..].chars().next()?;
    if !is_unquoted_sheet_char(first) {
        return None;
    }
    let mut end = start;
    while end < bytes.len() {
        let character = source[end..].chars().next()?;
        if !is_unquoted_sheet_char(character) {
            break;
        }
        end += character.len_utf8();
    }
    Some(ParsedSheetToken {
        start,
        end,
        name: source[start..end].to_string(),
    })
}

fn skip_string(source: &str, start: usize) -> usize {
    let bytes = source.as_bytes();
    let mut index = start + 1;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            index += 1;
            if bytes.get(index) == Some(&b'"') {
                index += 1;
            } else {
                break;
            }
        } else {
            index += source[index..]
                .chars()
                .next()
                .map(char::len_utf8)
                .unwrap_or(1);
        }
    }
    index
}

fn sheet_names_equal(left: &str, right: &str) -> bool {
    left.to_lowercase() == right.to_lowercase()
}

fn is_unquoted_sheet_char(character: char) -> bool {
    !character.is_whitespace()
        && !matches!(
            character,
            '"' | '\''
                | '!'
                | ':'
                | '+'
                | '-'
                | '*'
                | '/'
                | '^'
                | '&'
                | '='
                | '<'
                | '>'
                | '('
                | ')'
                | ','
                | '%'
                | '['
                | ']'
        )
}

fn sheet_token(name: &str) -> String {
    if is_simple_sheet_name(name) {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

/// A qualifier naming one sheet, or the span between two. Excel writes a span
/// unquoted when both endpoints are simple and `'first:last'` otherwise.
fn qualifier_token(first: &str, last: &str) -> String {
    if sheet_names_equal(first, last) {
        return sheet_token(first);
    }
    if is_simple_sheet_name(first) && is_simple_sheet_name(last) {
        return format!("{first}:{last}");
    }
    format!(
        "'{}:{}'",
        first.replace('\'', "''"),
        last.replace('\'', "''")
    )
}

fn is_simple_sheet_name(name: &str) -> bool {
    !name.is_empty()
        && name
            .chars()
            .all(|character| character.is_ascii_alphanumeric() || character == '_')
        && !name
            .chars()
            .next()
            .is_some_and(|character| character.is_ascii_digit())
        && CellRef::parse_a1(name).is_err()
        && !is_r1c1_reference(name)
}

fn is_r1c1_reference(name: &str) -> bool {
    let upper = name.to_ascii_uppercase();
    if matches!(upper.as_str(), "R" | "C") {
        return true;
    }
    let Some(rest) = upper.strip_prefix('R') else {
        return false;
    };
    let digits = rest.bytes().take_while(u8::is_ascii_digit).count();
    rest.as_bytes().get(digits) == Some(&b'C')
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
        Expr::FuncCall { name, args } => Expr::FuncCall {
            name: name.clone(),
            args: args
                .iter()
                .map(|a| transform(a, op, matches_target, changed))
                .collect(),
        },
        _ => expr.clone(),
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
fn remap_span(range: CellRange, op: &Op) -> Remapped<CellRange> {
    match *op {
        Op::DeleteRows { at, count, .. } => clip_span(range, Axis::Row, at, count),
        Op::DeleteCols { at, count, .. } => clip_span(range, Axis::Col, at, count),
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

#[derive(Clone, Copy, PartialEq, Eq)]
enum Axis {
    Row,
    Col,
}

impl Axis {
    fn is_edited_by(self, op: &Op) -> bool {
        matches!(
            (self, op),
            (Axis::Row, Op::InsertRows { .. } | Op::DeleteRows { .. })
                | (Axis::Col, Op::InsertCols { .. } | Op::DeleteCols { .. })
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

    fn charted(wb: &mut Workbook, sheet: SheetId, formulas: &[&str]) {
        wb.sheet_mut(sheet).unwrap().charts = vec![xlsx_model::SheetChart {
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
        charted(
            &mut w,
            SheetId(0),
            &["OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),1)"],
        );
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

    /// Excel's 3-D reference covers every sheet between its endpoints in
    /// workbook order, so an edit on an interior sheet moves it just as one on
    /// an endpoint does.
    #[test]
    fn a_three_dimensional_reference_follows_an_edit_on_any_sheet_it_covers() {
        for target in 0..3 {
            let mut workbook = wb(&["Jan", "Feb", "Mar", "Report"]);
            set_formula(&mut workbook, SheetId(3), "A1", "SUM('Jan:Mar'!$A$1:$A$5)");
            charted(&mut workbook, SheetId(3), &["'Jan:Mar'!$A$1:$A$5"]);
            workbook.defined_names = vec![defined("Span", "'Jan:Mar'!$A$1:$A$5")];
            let op = Op::InsertRows {
                sheet: SheetId(target),
                at: 0,
                count: 2,
            };
            remap_formulas(&mut workbook, &op).unwrap();
            remap_defined_names(&mut workbook, &op).unwrap();
            remap_charts(&mut workbook, &op).unwrap();
            assert_eq!(
                formula(&workbook, SheetId(3), "A1").as_deref(),
                Some("SUM('Jan:Mar'!$A$3:$A$7)"),
                "sheet {target}"
            );
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
    fn rename_rejects_formula_growth_past_the_length_cap() {
        let mut workbook = wb(&["S", "Formula"]);
        set_formula(&mut workbook, SheetId(1), "A1", "S!A1");
        let original = workbook.clone();
        let error = rename_sheet_references(&mut workbook, "S", &"x".repeat(MAX_FORMULA_BYTES))
            .unwrap_err();
        assert!(matches!(error, OpError::FormulaNotRewritable { .. }));
        assert_eq!(workbook, original);
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

    /// A whole-axis reference nested in a call is beyond the rewriter, and
    /// leaving it stale is exactly what this refuses to do.
    #[test]
    fn unrewritable_name_aimed_at_the_edited_sheet_refuses_the_edit() {
        let mut workbook = wb(&["Data", "Other"]);
        workbook.defined_names = vec![defined(
            "Dynamic",
            "OFFSET(Data!$A$1,0,0,COUNTA(Data!$A:$A),1)",
        )];
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
            ..defined("Dynamic", "OFFSET($A$1,0,0,COUNTA($A:$A),1)")
        }];

        let error = remap_defined_names(&mut workbook, &insert_rows(0, 0, 1)).unwrap_err();
        assert!(matches!(error, OpError::DefinedNameNotRewritable { .. }));
    }

    #[test]
    fn structural_edits_refuse_defined_names_that_would_outgrow_the_length_cap() {
        let mut workbook = wb(&["Data"]);
        let padding = "+0".repeat((MAX_FORMULA_BYTES - 16) / 2);
        workbook.defined_names = vec![defined("Long", &format!("Data!$A$9{padding}"))];
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
            let inverse = rename_defined_names(&mut workbook, function, "Renamed");
            assert!(inverse.is_none(), "{formula} must not change");
            assert_eq!(workbook.defined_names, vec![defined("Total", &formula)]);
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
        let inverse = rename_defined_names(&mut workbook, "Data", "New Data");
        assert!(inverse.is_some());
        assert_eq!(
            workbook.defined_names,
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
        rename_defined_names(&mut workbook, "SUM", "Renamed");
        assert_eq!(
            workbook.defined_names,
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
            rename_hyperlink_location("#Target!A1", "Target", "Renamed"),
            "#Renamed!A1"
        );
        assert_eq!(
            rename_hyperlink_location("Target!A1", "Target", "Renamed"),
            "Renamed!A1"
        );
        assert_eq!(
            rename_hyperlink_location("#Target!A1", "Target", "New Target"),
            "#'New Target'!A1"
        );
    }

    #[test]
    fn rename_hyperlink_location_handles_quoted_and_bare_targets() {
        assert_eq!(
            rename_hyperlink_location("#'My Sheet'!A1", "My Sheet", "Renamed"),
            "#Renamed!A1"
        );
        assert_eq!(
            rename_hyperlink_location("'My Sheet'!A1:B2", "My Sheet", "Renamed"),
            "Renamed!A1:B2"
        );
        assert_eq!(
            rename_hyperlink_location("#'It''s Data'!A1", "It's Data", "Renamed"),
            "#Renamed!A1"
        );
        for location in ["#MyRange", "MyRange", "#Target", "#A1"] {
            assert_eq!(
                rename_hyperlink_location(location, "Target", "Renamed"),
                location
            );
        }
        assert_eq!(
            rename_hyperlink_location("#Target!A1#2", "Target", "Renamed"),
            "#Renamed!A1#2"
        );
    }

    #[test]
    fn interval_clip_edges() {
        assert_eq!(clip_interval(0, 9, 2, 1), Some((0, 8)));
        assert_eq!(clip_interval(4, 6, 4, 3), None);
        assert_eq!(clip_interval(2, 9, 2, 4), Some((2, 5)));
        assert_eq!(clip_interval(0, 1, 5, 2), Some((0, 1)));
    }
}
