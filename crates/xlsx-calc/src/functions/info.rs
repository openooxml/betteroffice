//! information / type-test functions. the IS* predicates inspect their
//! argument's type and never propagate an error.

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{EvalContext, boolean, err, evaluate, num};
use crate::parser::Expr;

/// ISEVEN(number): truncates toward zero, then tests parity.
pub(crate) fn iseven(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    parity(args, ctx, 0)
}

/// ISODD(number): truncates toward zero, then tests parity.
pub(crate) fn isodd(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    parity(args, ctx, 1)
}

fn parity(args: &[Expr], ctx: &EvalContext<'_>, want: i64) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    match crate::eval::to_number(&evaluate(&args[0], ctx)) {
        Ok(n) if !n.is_finite() => err(ErrorValue::Num),
        Ok(n) => boolean((n.trunc() as i64).rem_euclid(2) == want),
        Err(e) => err(e),
    }
}

pub(crate) fn isblank(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    test(args, ctx, |v| matches!(v, CellValue::Empty))
}

pub(crate) fn isnumber(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    test(args, ctx, |v| matches!(v, CellValue::Number { .. }))
}

pub(crate) fn istext(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    test(args, ctx, |v| matches!(v, CellValue::Text { .. }))
}

pub(crate) fn islogical(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    test(args, ctx, |v| matches!(v, CellValue::Bool { .. }))
}

pub(crate) fn iserror(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    test(args, ctx, |v| matches!(v, CellValue::Error { .. }))
}

/// ISERR: any error except #N/A.
pub(crate) fn iserr(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    test(
        args,
        ctx,
        |v| matches!(v, CellValue::Error { value } if *value != ErrorValue::NA),
    )
}

pub(crate) fn isna(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    test(args, ctx, |v| {
        matches!(
            v,
            CellValue::Error {
                value: ErrorValue::NA
            }
        )
    })
}

/// NA(): the #N/A error literal.
pub(crate) fn na(args: &[Expr], _ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() {
        err(ErrorValue::NA)
    } else {
        err(ErrorValue::Value)
    }
}

/// N(value): numbers and bools become numbers, text becomes 0, errors pass
/// through, everything else is 0.
pub(crate) fn n(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    match evaluate(&args[0], ctx) {
        CellValue::Number { value } => num(value),
        CellValue::Bool { value } => num(if value { 1.0 } else { 0.0 }),
        CellValue::Error { value } => err(value),
        _ => num(0.0),
    }
}

fn test(args: &[Expr], ctx: &EvalContext<'_>, pred: fn(&CellValue) -> bool) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    let checkpoint = ctx.budget_error_checkpoint();
    let value = evaluate(&args[0], ctx);
    ctx.handle_budget_errors_since(checkpoint);
    boolean(pred(&value))
}
