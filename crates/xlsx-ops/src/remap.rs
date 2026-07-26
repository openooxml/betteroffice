//! rewriting stored formulas on row/column insert/delete: refs shift, ranges
//! clip, wholly deleted refs collapse to `#REF!`. runs before cells shift.

use std::collections::HashMap;

use xlsx_calc::lexer::MAX_FORMULA_BYTES;
use xlsx_calc::parse_formula;
use xlsx_calc::parser::Expr;
use xlsx_model::addr::{MAX_COLS, MAX_ROWS, col_to_letters};
use xlsx_model::{CellRange, CellRef, DefinedName, ErrorValue, SheetId, Workbook};

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

    let mut restores: Vec<Op> = Vec::new();
    let mut edits: Vec<(SheetId, CellRef, String)> = Vec::new();
    for (i, sheet) in wb.sheets.iter().enumerate() {
        let owner = SheetId(i as u32);
        let matches =
            |ref_sheet: &Option<String>| resolves_to_target(ref_sheet, owner, target, &names);
        for (cell, c) in sheet.iter_cells() {
            let Some(src) = &c.formula else {
                continue;
            };
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
    let mut restores = Vec::new();
    let mut edits = Vec::new();
    for (index, sheet) in wb.sheets.iter().enumerate() {
        let owner = SheetId(index as u32);
        let matches =
            |ref_sheet: &Option<String>| resolves_to_target(ref_sheet, owner, target, &names);
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

/// Rewrites defined-name formulas through a row or column op, the same way
/// cell formulas are. A scoped name resolves unqualified references against
/// its own sheet; a global name only follows qualified ones. A name aimed at
/// the edited sheet that the rewriter cannot express is refused rather than
/// left pointing at the pre-edit addresses.
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

    let previous = wb.defined_names.clone();
    let mut rewritten = Vec::with_capacity(previous.len());
    for defined in &previous {
        let mut updated = defined.clone();
        let scoped = defined.local_sheet == Some(target);
        let global_target =
            defined.local_sheet.is_none() && wb.sheets.len() == 1 && target == SheetId(0);
        let matches = |ref_sheet: &Option<String>| match ref_sheet {
            Some(name) => names.get(&name.to_lowercase()).copied() == Some(target),
            None => scoped || global_target,
        };
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
            _ if scoped || mentions_sheet(&defined.formula, &target_name) => {
                return Err(OpError::DefinedNameNotRewritable {
                    name: defined.name.clone(),
                });
            }
            _ => {}
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

/// Whether an unrewritable formula names `sheet` as a reference qualifier, so
/// leaving it untouched would strand it on the pre-edit addresses.
fn mentions_sheet(source: &str, sheet: &str) -> bool {
    let bytes = source.as_bytes();
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
        let qualified_by =
            |token: &ParsedSheetToken| !external && sheet_names_equal(&token.name, sheet);
        if bytes.get(first.end) == Some(&b'!') {
            if qualified_by(&first) {
                return true;
            }
            index = first.end + 1;
            continue;
        }
        if bytes.get(first.end) == Some(&b':')
            && let Some(second) = parse_sheet_token(source, first.end + 1)
            && bytes.get(second.end) == Some(&b'!')
        {
            if qualified_by(&first) || qualified_by(&second) {
                return true;
            }
            index = second.end + 1;
            continue;
        }
        index = first.end;
    }
    false
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

/// Single pass over `source`: rewrites every qualified reference to `old_name`
/// and flags bare tokens that match it but carry no `!` to bind them.
fn scan_sheet_rename(source: &str, old_name: &str, new_name: &str) -> SheetRename {
    if old_name == new_name {
        return SheetRename {
            formula: source.to_string(),
            ambiguous: false,
        };
    }
    let bytes = source.as_bytes();
    let mut replacements = Vec::new();
    let mut ambiguous = false;
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
            if !external && sheet_names_equal(&first.name, old_name) {
                replacements.push((first.start, first.end));
            }
            index = first.end + 1;
            continue;
        }
        if bytes.get(first.end) == Some(&b':')
            && let Some(second) = parse_sheet_token(source, first.end + 1)
            && bytes.get(second.end) == Some(&b'!')
        {
            if !external && sheet_names_equal(&first.name, old_name) {
                replacements.push((first.start, first.end));
            }
            if !external && sheet_names_equal(&second.name, old_name) {
                replacements.push((second.start, second.end));
            }
            index = second.end + 1;
            continue;
        }
        if !external
            && sheet_names_equal(&first.name, old_name)
            && !is_function_call(source, &first)
        {
            ambiguous = true;
        }
        index = first.end;
    }
    let formula = if replacements.is_empty() {
        source.to_string()
    } else {
        let replacement = sheet_token(new_name);
        let mut output = String::with_capacity(source.len());
        let mut copied_until = 0;
        for (start, end) in replacements {
            output.push_str(&source[copied_until..start]);
            output.push_str(&replacement);
            copied_until = end;
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
    let simple = !name.is_empty()
        && name
            .chars()
            .all(|character| character.is_ascii_alphanumeric() || character == '_')
        && !name
            .chars()
            .next()
            .is_some_and(|character| character.is_ascii_digit())
        && CellRef::parse_a1(name).is_err()
        && !is_r1c1_reference(name);
    if simple {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
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

/// whether a reference read from `owner` points at the edited sheet `target`.
/// unqualified refs bind to `owner`; unknown sheet names never match.
fn resolves_to_target(
    ref_sheet: &Option<String>,
    owner: SheetId,
    target: SheetId,
    names: &HashMap<String, SheetId>,
) -> bool {
    let resolved = match ref_sheet {
        None => Some(owner),
        Some(name) => names.get(&name.to_lowercase()).copied(),
    };
    resolved == Some(target)
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
                defined("ThreeD", "'New Data':Other!$A$1"),
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
