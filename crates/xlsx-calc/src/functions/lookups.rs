//! lookup and reference functions: VLOOKUP/HLOOKUP/MATCH exact and approximate
//! modes, INDEX area form, XLOOKUP exact-match subset, OFFSET.

use std::cmp::Ordering;

use xlsx_model::{CellRange, CellRef, CellValue, ErrorValue};

use crate::eval::{Area, EvalContext, as_area, cmp_values, err, evaluate, num};
use crate::parser::Expr;
use crate::reference::offset_rect;

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
    let mut found = None;
    for i in 0..lines {
        let key = if vertical {
            area.get(ctx, i, 0)
        } else {
            area.get(ctx, 0, i)
        };
        let key = match key {
            Ok(key) => key,
            Err(error) => return err(error),
        };
        let ordering = cmp_values(&key, &target);
        if approximate {
            if ordering != Ordering::Greater {
                found = Some(i);
            } else {
                break;
            }
        } else if ordering == Ordering::Equal {
            found = Some(i);
            break;
        }
    }
    match found {
        Some(i) => {
            let off = index as usize - 1;
            let value = if vertical {
                area.get(ctx, i, off)
            } else {
                area.get(ctx, off, i)
            };
            match value {
                Ok(value) => value,
                Err(error) => err(error),
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
    let values = match area.values(ctx) {
        Ok(values) => values,
        Err(error) => return err(error),
    };
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
    match area.get(ctx, row as usize - 1, col as usize - 1) {
        Ok(value) => value,
        Err(error) => err(error),
    }
}

/// OFFSET(reference, rows, cols, [height], [width]): a negative size extends
/// back from the shifted corner; a zero size or a rectangle off the sheet is
/// #REF!, and a multi-cell result in scalar context is #VALUE!.
pub(crate) fn offset(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match offset_area(args, ctx) {
        Ok(area) if area.rows == 1 && area.cols == 1 => match area.get(ctx, 0, 0) {
            Ok(value) => value,
            Err(error) => err(error),
        },
        Ok(_) => err(ErrorValue::Value),
        Err(error) => err(error),
    }
}

/// OFFSET's reference result, for callers that take an area rather than a
/// value; `as_area` routes nested OFFSET calls back through here.
/// INDIRECT(text, [a1]): the reference `text` spells. only a reference-shaped
/// expression is accepted, so the text cannot smuggle in another call.
pub(crate) fn indirect_area(args: &[Expr], ctx: &EvalContext<'_>) -> Result<Area, ErrorValue> {
    if args.is_empty() || args.len() > 2 {
        return Err(ErrorValue::Value);
    }
    if let Some(style) = args.get(1)
        && !crate::eval::to_bool(&evaluate(style, ctx)).unwrap_or(true)
    {
        return Err(ErrorValue::Ref);
    }
    let text = match evaluate(&args[0], ctx) {
        CellValue::Text { value } => value,
        CellValue::Error { value } => return Err(value),
        _ => return Err(ErrorValue::Ref),
    };
    let expr = crate::parse_formula(&text).map_err(|_| ErrorValue::Ref)?;
    if !matches!(
        expr,
        Expr::Ref { .. } | Expr::Range { .. } | Expr::ColumnRange { .. }
    ) {
        return Err(ErrorValue::Ref);
    }
    as_area(&expr, ctx).ok_or(ErrorValue::Ref)
}

/// INDIRECT used as a value: the top-left cell of the reference it spells.
pub(crate) fn indirect(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match indirect_area(args, ctx) {
        Ok(area) => match area.get(ctx, 0, 0) {
            Ok(value) => value,
            Err(e) => err(e),
        },
        Err(e) => err(e),
    }
}

pub(crate) fn offset_area(args: &[Expr], ctx: &EvalContext<'_>) -> Result<Area, ErrorValue> {
    if args.len() < 3 || args.len() > 5 {
        return Err(ErrorValue::Value);
    }
    let anchor = as_area(&args[0], ctx).ok_or(ErrorValue::Value)?;
    let rows = nth_int(args, ctx, 1)?;
    let cols = nth_int(args, ctx, 2)?;
    let height = match args.get(3) {
        Some(_) => Some(nth_int(args, ctx, 3)?),
        None => None,
    };
    let width = match args.get(4) {
        Some(_) => Some(nth_int(args, ctx, 4)?),
        None => None,
    };
    let bounds = CellRange::new(
        anchor.start,
        CellRef::new(
            anchor.start.row + anchor.rows as u32 - 1,
            anchor.start.col + anchor.cols as u32 - 1,
        ),
    );
    let rect = offset_rect(bounds, rows, cols, height, width).ok_or(ErrorValue::Ref)?;
    Ok(Area {
        sheet: anchor.sheet,
        start: rect.start,
        rows: (rect.end.row - rect.start.row + 1) as usize,
        cols: (rect.end.col - rect.start.col + 1) as usize,
    })
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
    if lookup.cell_count() != result.cell_count() {
        return err(ErrorValue::Value);
    }
    let count = lookup.cell_count().unwrap_or(0);
    let cols = u64::try_from(lookup.cols).unwrap_or(0);
    for index in 0..count {
        let row = match usize::try_from(index / cols) {
            Ok(row) => row,
            Err(_) => return err(ErrorValue::Num),
        };
        let col = match usize::try_from(index % cols) {
            Ok(col) => col,
            Err(_) => return err(ErrorValue::Num),
        };
        let key = match lookup.get(ctx, row, col) {
            Ok(key) => key,
            Err(error) => return err(error),
        };
        if cmp_values(&key, &target) == Ordering::Equal {
            return match result.get(ctx, row, col) {
                Ok(value) => value,
                Err(error) => err(error),
            };
        }
    }
    if args.len() >= 4 {
        evaluate(&args[3], ctx)
    } else {
        err(ErrorValue::NA)
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

/// ROW([reference]): the 1-based row of the reference's top-left cell, or of
/// the calling cell when no reference is given.
pub(crate) fn row(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    reference_scalar(
        args,
        ctx,
        |area| area.start.row as f64 + 1.0,
        |cell| cell.row as f64 + 1.0,
    )
}

/// COLUMN([reference]): the 1-based column of the reference's top-left cell, or
/// of the calling cell when no reference is given.
pub(crate) fn column(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    reference_scalar(
        args,
        ctx,
        |area| area.start.col as f64 + 1.0,
        |cell| cell.col as f64 + 1.0,
    )
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

fn reference_scalar(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    pick: fn(&Area) -> f64,
    here: fn(CellRef) -> f64,
) -> CellValue {
    match args {
        [] => match ctx.cell {
            Some(cell) => num(here(cell)),
            None => err(ErrorValue::Value),
        },
        [arg] => match as_area(arg, ctx) {
            Some(area) => num(pick(&area)),
            None => err(ErrorValue::Value),
        },
        _ => err(ErrorValue::Value),
    }
}

fn reference_dim(args: &[Expr], ctx: &EvalContext<'_>, pick: fn(&Area) -> usize) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    match as_area(&args[0], ctx) {
        Some(area) => num(pick(&area) as f64),
        None => match evaluate(&args[0], ctx) {
            value @ CellValue::Error { .. } => value,
            _ => err(ErrorValue::Value),
        },
    }
}

/// TRANSPOSE(array): a 1x1 input transposes to itself and a blank to 0; a
/// wider one is the array form, which the engine does not implement.
pub(crate) fn transpose(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    let value = match as_area(&args[0], ctx) {
        Some(area) if area.rows == 1 && area.cols == 1 => match area.get(ctx, 0, 0) {
            Ok(value) => value,
            Err(error) => return err(error),
        },
        Some(_) => {
            ctx.record_unsupported_function();
            return err(ErrorValue::Value);
        }
        None => evaluate(&args[0], ctx),
    };
    match value {
        CellValue::Empty => num(0.0),
        value => value,
    }
}
