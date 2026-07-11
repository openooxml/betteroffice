//! math and trig functions. rounding follows excel: ROUND is half-away-from-zero,
//! ROUNDUP/ROUNDDOWN are directional, non-finite results are `#NUM!`.

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{Area, EvalContext, as_area, err, evaluate, num, to_number};
use crate::parser::Expr;

use super::criteria::{self, Criterion};
use super::{collect_numbers, finite, nth_int, nth_number};

pub(crate) fn sum(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match collect_numbers(args, ctx) {
        Ok(nums) => num(nums.iter().sum()),
        Err(e) => err(e),
    }
}

/// PRODUCT of every numeric argument; no numbers -> 0 (excel).
pub(crate) fn product(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match collect_numbers(args, ctx) {
        Ok(nums) if nums.is_empty() => num(0.0),
        Ok(nums) => num(nums.iter().product()),
        Err(e) => err(e),
    }
}

/// SUMIF(range, criteria, [sum_range]). sum_range is anchored at its top-left
/// with the criteria range's shape, so a mismatched size still aligns.
pub(crate) fn sumif(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 && args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let crit_area = match as_area(&args[0], ctx) {
        Some(a) => a,
        None => return err(ErrorValue::Value),
    };
    let criterion = criteria::criterion_from_arg(&args[1], ctx);
    let sum_area = if args.len() == 3 {
        match as_area(&args[2], ctx) {
            Some(a) => a,
            None => return err(ErrorValue::Value),
        }
    } else {
        match as_area(&args[0], ctx) {
            Some(a) => a,
            None => return err(ErrorValue::Value),
        }
    };
    let pairs = [(crit_area, criterion)];
    sum_matching(&pairs, &sum_area, ctx)
}

/// SUMIFS(sum_range, crit_range1, crit1, ...). all ranges share dimensions.
pub(crate) fn sumifs(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 3 {
        return err(ErrorValue::Value);
    }
    let sum_area = match as_area(&args[0], ctx) {
        Some(a) => a,
        None => return err(ErrorValue::Value),
    };
    match criteria::collect_pairs(&args[1..], ctx) {
        Some(pairs) if pairs[0].0.rows == sum_area.rows && pairs[0].0.cols == sum_area.cols => {
            sum_matching(&pairs, &sum_area, ctx)
        }
        _ => err(ErrorValue::Value),
    }
}

fn sum_matching(
    pairs: &[(Area, Criterion)],
    value_area: &Area,
    ctx: &EvalContext<'_>,
) -> CellValue {
    let cols = pairs[0].0.cols;
    let mut total = 0.0;
    for i in criteria::matching_indices(pairs, ctx) {
        let (r, c) = (i / cols, i % cols);
        if let CellValue::Number { value } = value_area.get(ctx, r, c) {
            total += value;
        }
    }
    num(total)
}

/// SUMPRODUCT(array1, [array2], ...). element-wise product summed; non-numeric
/// cells count as 0. all arrays must share the same length.
pub(crate) fn sumproduct(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() {
        return err(ErrorValue::Value);
    }
    let mut arrays: Vec<Vec<f64>> = Vec::with_capacity(args.len());
    for arg in args {
        match as_area(arg, ctx) {
            Some(area) => {
                let mut col = Vec::with_capacity(area.len());
                for v in area.values(ctx) {
                    match v {
                        CellValue::Number { value } => col.push(value),
                        CellValue::Error { value } => return err(value),
                        _ => col.push(0.0),
                    }
                }
                arrays.push(col);
            }
            None => match to_number(&evaluate(arg, ctx)) {
                Ok(n) => arrays.push(vec![n]),
                Err(e) => return err(e),
            },
        }
    }
    let len = arrays[0].len();
    if arrays.iter().any(|a| a.len() != len) {
        return err(ErrorValue::Value);
    }
    let mut total = 0.0;
    for i in 0..len {
        let mut prod = 1.0;
        for a in &arrays {
            prod *= a[i];
        }
        total += prod;
    }
    num(total)
}

pub(crate) fn abs(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::abs)
}

pub(crate) fn sign(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, |x| {
        if x > 0.0 {
            1.0
        } else if x < 0.0 {
            -1.0
        } else {
            0.0
        }
    })
}

pub(crate) fn sqrt(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_number(args, ctx) {
        Ok(x) if x < 0.0 => err(ErrorValue::Num),
        Ok(x) => num(x.sqrt()),
        Err(e) => err(e),
    }
}

pub(crate) fn exp(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_number(args, ctx) {
        Ok(x) => finite(x.exp()),
        Err(e) => err(e),
    }
}

pub(crate) fn ln(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    positive_log(args, ctx, f64::ln)
}

pub(crate) fn log10(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    positive_log(args, ctx, f64::log10)
}

/// LOG(number, [base]); base defaults to 10.
pub(crate) fn log(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() || args.len() > 2 {
        return err(ErrorValue::Value);
    }
    let base = if args.len() == 2 {
        match nth_number(args, ctx, 1) {
            Ok(b) => b,
            Err(e) => return err(e),
        }
    } else {
        10.0
    };
    match nth_number(args, ctx, 0) {
        Ok(x) if x <= 0.0 || base <= 0.0 || base == 1.0 => err(ErrorValue::Num),
        Ok(x) => finite(x.log(base)),
        Err(e) => err(e),
    }
}

/// LN / LOG10: positive inputs only. uses the dedicated libm routine (not
/// `x.log(base)`) so exact powers round-trip precisely.
fn positive_log(args: &[Expr], ctx: &EvalContext<'_>, f: fn(f64) -> f64) -> CellValue {
    match one_number(args, ctx) {
        Ok(x) if x <= 0.0 => err(ErrorValue::Num),
        Ok(x) => finite(f(x)),
        Err(e) => err(e),
    }
}

pub(crate) fn pi(args: &[Expr], _ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() {
        num(std::f64::consts::PI)
    } else {
        err(ErrorValue::Value)
    }
}

pub(crate) fn power(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(x), Ok(y)) => finite(x.powf(y)),
        (Err(e), _) | (_, Err(e)) => err(e),
    }
}

/// MOD(n, d) = n - d*INT(n/d); sign follows the divisor. d = 0 -> #DIV/0!.
pub(crate) fn mod_(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(_), Ok(0.0)) => err(ErrorValue::Div0),
        (Ok(n), Ok(d)) => num(n - d * (n / d).floor()),
        (Err(e), _) | (_, Err(e)) => err(e),
    }
}

pub(crate) fn int(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::floor)
}

/// TRUNC(number, [digits]); digits default 0. truncates toward zero.
pub(crate) fn trunc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    directional(args, ctx, |scaled| scaled.trunc())
}

/// ROUNDUP(number, digits): away from zero.
pub(crate) fn roundup(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    directional(args, ctx, |scaled| {
        if scaled < 0.0 {
            scaled.floor()
        } else {
            scaled.ceil()
        }
    })
}

/// ROUNDDOWN(number, digits): toward zero (same as TRUNC with digits).
pub(crate) fn rounddown(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    directional(args, ctx, |scaled| scaled.trunc())
}

/// ROUND(number, digits): half away from zero.
pub(crate) fn round(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let x = match nth_number(args, ctx, 0) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let digits = match nth_int(args, ctx, 1) {
        Ok(d) => d as i32,
        Err(e) => return err(e),
    };
    let factor = 10f64.powi(digits);
    finite((x * factor).round() / factor)
}

/// MROUND(number, multiple): nearest multiple, half away from zero. opposite
/// signs -> #NUM!; multiple 0 -> 0.
pub(crate) fn mround(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(_), Ok(0.0)) => num(0.0),
        (Ok(x), Ok(m)) if (x < 0.0) != (m < 0.0) && x != 0.0 => err(ErrorValue::Num),
        (Ok(x), Ok(m)) => num((x / m).round() * m),
        (Err(e), _) | (_, Err(e)) => err(e),
    }
}

/// CEILING(number, significance): away from zero to a multiple of significance.
/// positive number with negative significance is #NUM!; significance 0 -> 0.
pub(crate) fn ceiling(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    round_to_multiple(args, ctx, f64::ceil)
}

/// FLOOR(number, significance): toward zero to a multiple of significance.
pub(crate) fn floor(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    round_to_multiple(args, ctx, f64::floor)
}

fn round_to_multiple(args: &[Expr], ctx: &EvalContext<'_>, f: fn(f64) -> f64) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(_), Ok(0.0)) => num(0.0),
        (Ok(x), Ok(s)) if x > 0.0 && s < 0.0 => err(ErrorValue::Num),
        (Ok(x), Ok(s)) => num(f(x / s) * s),
        (Err(e), _) | (_, Err(e)) => err(e),
    }
}

fn one_number(args: &[Expr], ctx: &EvalContext<'_>) -> Result<f64, ErrorValue> {
    if args.len() != 1 {
        return Err(ErrorValue::Value);
    }
    nth_number(args, ctx, 0)
}

fn unary(args: &[Expr], ctx: &EvalContext<'_>, f: fn(f64) -> f64) -> CellValue {
    match one_number(args, ctx) {
        Ok(x) => num(f(x)),
        Err(e) => err(e),
    }
}

/// scale by 10^digits, apply a rounding rule, unscale.
fn directional(args: &[Expr], ctx: &EvalContext<'_>, rule: fn(f64) -> f64) -> CellValue {
    if args.is_empty() || args.len() > 2 {
        return err(ErrorValue::Value);
    }
    let x = match nth_number(args, ctx, 0) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let digits = if args.len() == 2 {
        match nth_int(args, ctx, 1) {
            Ok(d) => d as i32,
            Err(e) => return err(e),
        }
    } else {
        0
    };
    let factor = 10f64.powi(digits);
    finite(rule(x * factor) / factor)
}
