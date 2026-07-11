//! lookup and reference functions: VLOOKUP/HLOOKUP/MATCH exact and approximate
//! modes, INDEX area form, XLOOKUP exact-match subset.

use std::cmp::Ordering;

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{Area, EvalContext, as_area, cmp_values, err, evaluate, num};
use crate::parser::Expr;

use super::{nth_int, nth_number};

/// VLOOKUP(value, table, col_index, [range_lookup]). range_lookup defaults to
/// TRUE (approximate match on a first column sorted ascending).
pub(crate) fn vlookup(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    table_lookup(args, ctx, true)
}

/// HLOOKUP(value, table, row_index, [range_lookup]).
pub(crate) fn hlookup(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    table_lookup(args, ctx, false)
}

fn table_lookup(args: &[Expr], ctx: &EvalContext<'_>, vertical: bool) -> CellValue {
    if args.len() < 3 || args.len() > 4 {
        return err(ErrorValue::Value);
    }
    let target = evaluate(&args[0], ctx);
    if let CellValue::Error { value } = target {
        return err(value);
    }
    let area = match as_area(&args[1], ctx) {
        Some(a) => a,
        None => return err(ErrorValue::Value),
    };
    let index = match nth_int(args, ctx, 2) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    if index < 1 {
        return err(ErrorValue::Value);
    }
    let approximate = if args.len() == 4 {
        match nth_number(args, ctx, 3) {
            Ok(v) => v != 0.0,
            Err(e) => return err(e),
        }
    } else {
        true
    };
    let (lines, depth) = if vertical {
        (area.rows, area.cols)
    } else {
        (area.cols, area.rows)
    };
    if index as usize > depth {
        return err(ErrorValue::Ref);
    }
    let key = |i: usize| {
        if vertical {
            area.get(ctx, i, 0)
        } else {
            area.get(ctx, 0, i)
        }
    };
    let found = if approximate {
        approximate_row(lines, &target, key)
    } else {
        (0..lines).find(|&i| cmp_values(&key(i), &target) == Ordering::Equal)
    };
    match found {
        Some(i) => {
            let off = index as usize - 1;
            if vertical {
                area.get(ctx, i, off)
            } else {
                area.get(ctx, off, i)
            }
        }
        None => err(ErrorValue::NA),
    }
}

/// MATCH(value, lookup_area, [match_type]). match_type 1 (default) finds the
/// largest value <= target in an ascending list; 0 is exact; -1 finds the
/// smallest value >= target in a descending list.
pub(crate) fn match_(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 2 || args.len() > 3 {
        return err(ErrorValue::Value);
    }
    let target = evaluate(&args[0], ctx);
    if let CellValue::Error { value } = target {
        return err(value);
    }
    let area = match as_area(&args[1], ctx) {
        Some(a) => a,
        None => return err(ErrorValue::Value),
    };
    let match_type = if args.len() == 3 {
        match nth_int(args, ctx, 2) {
            Ok(n) => n,
            Err(e) => return err(e),
        }
    } else {
        1
    };
    let values = area.values(ctx);
    let pos = match match_type {
        0 => values
            .iter()
            .position(|v| cmp_values(v, &target) == Ordering::Equal),
        1 => approximate_row(values.len(), &target, |i| values[i].clone()),
        _ => {
            let mut found = None;
            for (i, v) in values.iter().enumerate() {
                if cmp_values(v, &target) != Ordering::Less {
                    found = Some(i);
                } else {
                    break;
                }
            }
            found
        }
    };
    match pos {
        Some(i) => num(i as f64 + 1.0),
        None => err(ErrorValue::NA),
    }
}

/// INDEX(area, row_num, [col_num]). for a single-row or single-column area the
/// lone index selects along that axis. 1-based; out of range -> #REF!.
pub(crate) fn index(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 2 || args.len() > 3 {
        return err(ErrorValue::Value);
    }
    let area = match as_area(&args[0], ctx) {
        Some(a) => a,
        None => return err(ErrorValue::Value),
    };
    let first = match nth_int(args, ctx, 1) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let second = if args.len() == 3 {
        match nth_int(args, ctx, 2) {
            Ok(n) => Some(n),
            Err(e) => return err(e),
        }
    } else {
        None
    };
    let (row, col) = match second {
        Some(c) => (first, c),
        None if area.rows == 1 => (1, first),
        None if area.cols == 1 => (first, 1),
        None => return err(ErrorValue::Ref),
    };
    if row < 1 || col < 1 || row as usize > area.rows || col as usize > area.cols {
        return err(ErrorValue::Ref);
    }
    area.get(ctx, row as usize - 1, col as usize - 1)
}

/// XLOOKUP(value, lookup_array, return_array, [if_not_found], ...): exact-match
/// subset; a missing value returns `if_not_found` or #N/A.
pub(crate) fn xlookup(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 3 || args.len() > 6 {
        return err(ErrorValue::Value);
    }
    let target = evaluate(&args[0], ctx);
    if let CellValue::Error { value } = target {
        return err(value);
    }
    let lookup = match as_area(&args[1], ctx) {
        Some(a) => a,
        None => return err(ErrorValue::Value),
    };
    let result = match as_area(&args[2], ctx) {
        Some(a) => a,
        None => return err(ErrorValue::Value),
    };
    let keys = lookup.values(ctx);
    let results = result.values(ctx);
    if keys.len() != results.len() {
        return err(ErrorValue::Value);
    }
    match keys
        .iter()
        .position(|v| cmp_values(v, &target) == Ordering::Equal)
    {
        Some(i) => results[i].clone(),
        None => {
            if args.len() >= 4 {
                evaluate(&args[3], ctx)
            } else {
                err(ErrorValue::NA)
            }
        }
    }
}

/// CHOOSE(index, value1, value2, ...): the index-th value (1-based); only the
/// chosen argument is evaluated.
pub(crate) fn choose(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 2 {
        return err(ErrorValue::Value);
    }
    let idx = match nth_int(args, ctx, 0) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let choices = &args[1..];
    if idx < 1 || idx as usize > choices.len() {
        return err(ErrorValue::Value);
    }
    evaluate(&choices[idx as usize - 1], ctx)
}

/// ROW([reference]): the 1-based row of the reference's top-left cell;
/// referenceless form is #VALUE! (calling cell unknown).
pub(crate) fn row(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    reference_scalar(args, ctx, |area| area.start.row as f64 + 1.0)
}

/// COLUMN([reference]): the 1-based column of the reference's top-left cell.
pub(crate) fn column(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    reference_scalar(args, ctx, |area| area.start.col as f64 + 1.0)
}

/// ROWS(area): the number of rows in a reference.
pub(crate) fn rows(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    reference_dim(args, ctx, |area| area.rows)
}

/// COLUMNS(area): the number of columns in a reference.
pub(crate) fn columns(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    reference_dim(args, ctx, |area| area.cols)
}

/// last index whose value is <= target scanning in order (excel's ascending
/// approximate match); target below the first value -> None.
fn approximate_row(
    len: usize,
    target: &CellValue,
    key: impl Fn(usize) -> CellValue,
) -> Option<usize> {
    let mut found = None;
    for i in 0..len {
        if cmp_values(&key(i), target) != Ordering::Greater {
            found = Some(i);
        } else {
            break;
        }
    }
    found
}

fn reference_scalar(args: &[Expr], ctx: &EvalContext<'_>, pick: fn(&Area) -> f64) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    match as_area(&args[0], ctx) {
        Some(area) => num(pick(&area)),
        None => err(ErrorValue::Value),
    }
}

fn reference_dim(args: &[Expr], ctx: &EvalContext<'_>, pick: fn(&Area) -> usize) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    match as_area(&args[0], ctx) {
        Some(area) => num(pick(&area) as f64),
        None => err(ErrorValue::Value),
    }
}
