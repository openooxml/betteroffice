//! logical functions. all take lazy arguments so only the taken branch is
//! evaluated: IF/IFS/SWITCH/IFERROR/IFNA never touch the paths they skip.

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{EvalContext, boolean, cmp_values, err, evaluate, range_values, to_bool};
use crate::parser::Expr;

/// IF(condition, then, [else]); omitted else yields FALSE.
pub(crate) fn if_(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 2 || args.len() > 3 {
        return err(ErrorValue::Value);
    }
    match to_bool(&evaluate(&args[0], ctx)) {
        Ok(true) => evaluate(&args[1], ctx),
        Ok(false) => {
            if args.len() == 3 {
                evaluate(&args[2], ctx)
            } else {
                boolean(false)
            }
        }
        Err(e) => err(e),
    }
}

/// IFERROR(value, value_if_error): the fallback replaces any error.
pub(crate) fn iferror(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let checkpoint = ctx.budget_error_checkpoint();
    match evaluate(&args[0], ctx) {
        CellValue::Error { .. } => {
            ctx.handle_budget_errors_since(checkpoint);
            evaluate(&args[1], ctx)
        }
        v => v,
    }
}

/// IFNA(value, value_if_na): the fallback replaces only #N/A.
pub(crate) fn ifna(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    match evaluate(&args[0], ctx) {
        CellValue::Error {
            value: ErrorValue::NA,
        } => evaluate(&args[1], ctx),
        v => v,
    }
}

/// IFS(cond1, val1, cond2, val2, ...): the value for the first true condition;
/// no true condition -> #N/A.
pub(crate) fn ifs(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() || !args.len().is_multiple_of(2) {
        return err(ErrorValue::Value);
    }
    for pair in args.chunks(2) {
        match to_bool(&evaluate(&pair[0], ctx)) {
            Ok(true) => return evaluate(&pair[1], ctx),
            Ok(false) => {}
            Err(e) => return err(e),
        }
    }
    err(ErrorValue::NA)
}

/// SWITCH(expr, match1, result1, ..., [default]): the result whose match
/// equals `expr`; a trailing odd argument is the default.
pub(crate) fn switch(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 3 {
        return err(ErrorValue::Value);
    }
    let subject = evaluate(&args[0], ctx);
    if let CellValue::Error { value } = subject {
        return err(value);
    }
    let rest = &args[1..];
    let mut i = 0;
    while i + 1 < rest.len() {
        let candidate = evaluate(&rest[i], ctx);
        if let CellValue::Error { value } = candidate {
            return err(value);
        }
        if cmp_values(&subject, &candidate) == std::cmp::Ordering::Equal {
            return evaluate(&rest[i + 1], ctx);
        }
        i += 2;
    }
    if rest.len() % 2 == 1 {
        evaluate(&rest[rest.len() - 1], ctx)
    } else {
        err(ErrorValue::NA)
    }
}

pub(crate) fn and(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    fold_bools(args, ctx, true, |a, b| a && b)
}

pub(crate) fn or(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    fold_bools(args, ctx, false, |a, b| a || b)
}

/// XOR: true when an odd number of logical inputs are true.
pub(crate) fn xor(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    fold_bools(args, ctx, false, |a, b| a ^ b)
}

pub(crate) fn not(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    match to_bool(&evaluate(&args[0], ctx)) {
        Ok(b) => boolean(!b),
        Err(e) => err(e),
    }
}

/// fold booleans across arguments and ranges; blanks/text are ignored, errors
/// propagate, and at least one logical value is required.
fn fold_bools(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    init: bool,
    combine: fn(bool, bool) -> bool,
) -> CellValue {
    let mut acc = init;
    let mut seen = false;
    for arg in args {
        let values = match arg {
            Expr::Range { sheet, range } => match range_values(sheet, range, ctx) {
                Ok(v) => v,
                Err(e) => return err(e),
            },
            _ => vec![evaluate(arg, ctx)],
        };
        for v in values {
            match v {
                CellValue::Bool { value } => {
                    acc = combine(acc, value);
                    seen = true;
                }
                CellValue::Number { value } => {
                    acc = combine(acc, value != 0.0);
                    seen = true;
                }
                CellValue::Empty | CellValue::Text { .. } => {}
                CellValue::Error { value } => return err(value),
            }
        }
    }
    if seen {
        boolean(acc)
    } else {
        err(ErrorValue::Value)
    }
}
