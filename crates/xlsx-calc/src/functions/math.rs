//! math and trig functions. rounding follows excel: ROUND is half-away-from-zero,
//! ROUNDUP/ROUNDDOWN are directional, non-finite results are `#NUM!`.

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{Area, EvalContext, as_area, bound_area, err, evaluate, num, to_number, to_text};
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
        Some(a) => bound_area(a, ctx),
        None => return err(ErrorValue::Value),
    };
    let criterion = criteria::criterion_from_arg(&args[1], ctx);
    let values = if args.len() == 3 { &args[2] } else { &args[0] };
    let sum_area = match criteria::aligned_area(values, ctx, crit_area.rows, crit_area.cols) {
        Some(a) => a,
        None => match as_area(values, ctx) {
            Some(a) => a,
            None => return err(ErrorValue::Value),
        },
    };
    let pairs = [(crit_area, criterion)];
    sum_matching(&pairs, &sum_area, ctx)
}

/// SUMIFS(sum_range, crit_range1, crit1, ...). all ranges share dimensions.
pub(crate) fn sumifs(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 3 {
        return err(ErrorValue::Value);
    }
    match criteria::collect_pairs(&args[1..], ctx) {
        Some(pairs) => {
            let (rows, cols) = (pairs[0].0.rows, pairs[0].0.cols);
            match criteria::aligned_area(&args[0], ctx, rows, cols) {
                Some(sum_area) => sum_matching(&pairs, &sum_area, ctx),
                None => err(ErrorValue::Value),
            }
        }
        None => err(ErrorValue::Value),
    }
}

fn sum_matching(
    pairs: &[(Area, Criterion)],
    value_area: &Area,
    ctx: &EvalContext<'_>,
) -> CellValue {
    let cols = pairs[0].0.cols;
    let mut total = 0.0;
    let indices = match criteria::matching_indices(pairs, ctx) {
        Ok(indices) => indices,
        Err(error) => return err(error),
    };
    for i in indices {
        let (r, c) = (i / cols, i % cols);
        match value_area.get_ref(ctx, r, c) {
            Ok(v) => {
                if let CellValue::Number { value } = *v {
                    total += value;
                }
            }
            Err(error) => return err(error),
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
    let mut arrays: Vec<((usize, usize), Vec<f64>)> = Vec::with_capacity(args.len());
    for arg in args {
        match as_area(arg, ctx) {
            Some(area) => {
                let values = match area.values_ref(ctx) {
                    Ok(values) => values,
                    Err(error) => return err(error),
                };
                let mut col = Vec::with_capacity(values.len());
                for v in values {
                    match v.as_ref() {
                        CellValue::Number { value } => col.push(*value),
                        CellValue::Error { value } => return err(*value),
                        _ => col.push(0.0),
                    }
                }
                arrays.push(((area.rows, area.cols), col));
            }
            // `--(range=key)` and the other computed operands are arrays, not
            // references, so they are read as blocks rather than coerced
            None => match crate::array::evaluate_array(arg, ctx) {
                crate::array::Value::Array(array) => {
                    let mut col = Vec::with_capacity(array.rows() * array.cols());
                    for row in 0..array.rows() {
                        for c in 0..array.cols() {
                            match array.at(row, c) {
                                CellValue::Number { value } => col.push(value),
                                CellValue::Error { value } => return err(value),
                                _ => col.push(0.0),
                            }
                        }
                    }
                    arrays.push(((array.rows(), array.cols()), col));
                }
                value => match to_number(&value.into_scalar()) {
                    Ok(n) => arrays.push(((1, 1), vec![n])),
                    Err(e) => return err(e),
                },
            },
        }
    }
    // excel pairs the operands by position, so they must share a shape: a
    // column and a row of the same length are not the same operand
    let shape = arrays[0].0;
    if arrays.iter().any(|(dims, _)| *dims != shape) {
        return err(ErrorValue::Value);
    }
    let mut total = 0.0;
    for i in 0..arrays[0].1.len() {
        let mut prod = 1.0;
        for (_, a) in &arrays {
            prod *= a[i];
        }
        total += prod;
    }
    num(total)
}

/// MMULT(array1, array2): the matrix product; `cols(array1)` must equal
/// `rows(array2)` and any non-numeric operand cell is `#VALUE!`. a cell holds
/// one value, so this is the top-left element excel caches in the anchor.
pub(crate) fn mmult(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let left = match matrix(&args[0], ctx) {
        Ok(m) => m,
        Err(e) => return err(e),
    };
    let right = match matrix(&args[1], ctx) {
        Ok(m) => m,
        Err(e) => return err(e),
    };
    if left.cols != right.rows {
        return err(ErrorValue::Value);
    }
    let mut total = 0.0;
    for k in 0..left.cols {
        total += left.values[k] * right.values[k * right.cols];
    }
    finite(total)
}

/// a rectangle of numbers in row-major order.
struct Matrix {
    rows: usize,
    cols: usize,
    values: Vec<f64>,
}

/// read an argument as a numeric rectangle; a non-reference is 1x1 and any
/// non-numeric cell is `#VALUE!`.
fn matrix(arg: &Expr, ctx: &EvalContext<'_>) -> Result<Matrix, ErrorValue> {
    let Some(area) = as_area(arg, ctx) else {
        return match evaluate(arg, ctx) {
            CellValue::Number { value } => Ok(Matrix {
                rows: 1,
                cols: 1,
                values: vec![value],
            }),
            CellValue::Error { value } => Err(value),
            _ => Err(ErrorValue::Value),
        };
    };
    let cells = area.values_ref(ctx)?;
    let mut values = Vec::with_capacity(cells.len());
    for cell in cells {
        match *cell {
            CellValue::Number { value } => values.push(value),
            CellValue::Error { value } => return Err(value),
            _ => return Err(ErrorValue::Value),
        }
    }
    Ok(Matrix {
        rows: area.rows,
        cols: area.cols,
        values,
    })
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

/// every cell of the arguments as a number, refusing anything that is not one.
fn strict_numbers(args: &[Expr], ctx: &EvalContext<'_>) -> Result<Vec<f64>, ErrorValue> {
    let mut nums = Vec::new();
    for arg in args {
        match as_area(arg, ctx) {
            Some(area) => {
                for value in area.values_ref(ctx)? {
                    match value.as_ref() {
                        CellValue::Number { value } => nums.push(*value),
                        CellValue::Error { value } => return Err(*value),
                        CellValue::Empty => {}
                        _ => return Err(ErrorValue::Value),
                    }
                }
            }
            None => match evaluate(arg, ctx) {
                CellValue::Number { value } => nums.push(value),
                CellValue::Error { value } => return Err(value),
                _ => return Err(ErrorValue::Value),
            },
        }
    }
    Ok(nums)
}

/// QUOTIENT(numerator, denominator): the integer part of the division,
/// truncated toward zero.
pub(crate) fn quotient(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(_), Ok(0.0)) => err(ErrorValue::Div0),
        (Ok(n), Ok(d)) => finite((n / d).trunc()),
        (Err(e), _) | (_, Err(e)) => err(e),
    }
}

/// SERIESSUM(x, n, m, coefficients): sum of `a_i * x^(n + (i-1) * m)`.
pub(crate) fn seriessum(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 4 {
        return err(ErrorValue::Value);
    }
    let (x, n, m) = match (
        nth_number(args, ctx, 0),
        nth_number(args, ctx, 1),
        nth_number(args, ctx, 2),
    ) {
        (Ok(x), Ok(n), Ok(m)) => (x, n, m),
        (Err(e), _, _) | (_, Err(e), _) | (_, _, Err(e)) => return err(e),
    };
    // a non-numeric coefficient is #VALUE!, not a term to skip: dropping one
    // would shift every later power
    let coefficients = match strict_numbers(&args[3..], ctx) {
        Ok(c) => c,
        Err(e) => return err(e),
    };
    let mut total = 0.0;
    for (i, a) in coefficients.iter().enumerate() {
        total += a * x.powf(n + (i as f64) * m);
    }
    finite(total)
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

pub(crate) fn tanh(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::tanh)
}

pub(crate) fn sin(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::sin)
}

pub(crate) fn cos(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::cos)
}

pub(crate) fn tan(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::tan)
}

pub(crate) fn sinh(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::sinh)
}

pub(crate) fn cosh(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::cosh)
}

/// ASIN/ACOS are `#NUM!` outside [-1, 1]; `finite` turns the NaN into it.
pub(crate) fn asin(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_number(args, ctx) {
        Ok(x) => finite(x.asin()),
        Err(e) => err(e),
    }
}

pub(crate) fn acos(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match one_number(args, ctx) {
        Ok(x) => finite(x.acos()),
        Err(e) => err(e),
    }
}

pub(crate) fn atan(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::atan)
}

/// ATAN2(x, y) — excel takes the x coordinate first, the reverse of `f64::atan2`.
/// both zero is `#DIV/0!`.
pub(crate) fn atan2(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(0.0), Ok(0.0)) => err(ErrorValue::Div0),
        (Ok(x), Ok(y)) => finite(y.atan2(x)),
        (Err(e), _) | (_, Err(e)) => err(e),
    }
}

pub(crate) fn degrees(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::to_degrees)
}

pub(crate) fn radians(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, f64::to_radians)
}

/// RANDBETWEEN(bottom, top): a volatile integer draw. measured against excel:
/// a raw `bottom > top` is `#NUM!`, the draw spans `ceil(bottom)..=floor(top)`,
/// and a span that rounds away to nothing collapses to `ceil(bottom)`.
pub(crate) fn randbetween(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let (bottom, top) = match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(bottom), Ok(top)) => (bottom, top),
        (Err(e), _) | (_, Err(e)) => return err(e),
    };
    if bottom > top {
        return err(ErrorValue::Num);
    }
    let (lo, hi) = (bottom.ceil(), top.floor());
    if hi < lo {
        return num(lo);
    }
    let span = hi - lo + 1.0;
    finite(lo + (ctx.next_random_unit() * span).floor().min(hi - lo))
}

/// SUMSQ(number, ...): the sum of the squares of every numeric argument.
pub(crate) fn sumsq(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match super::collect_numbers_deep(args, ctx) {
        Ok(nums) => finite(nums.iter().map(|x| x * x).sum()),
        Err(e) => err(e),
    }
}

/// CONVERT(number, from, to): excel's unit conversion, over the unit families
/// it defines. A metric unit may carry an SI prefix; the two units must share
/// a family.
pub(crate) fn convert(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let value = match nth_number(args, ctx, 0) {
        Ok(value) => value,
        Err(e) => return err(e),
    };
    let (from, to) = match (
        to_text(&evaluate(&args[1], ctx)),
        to_text(&evaluate(&args[2], ctx)),
    ) {
        (Ok(a), Ok(b)) => (a, b),
        (Err(e), _) | (_, Err(e)) => return err(e),
    };
    let (Some((from_family, from_scale)), Some((to_family, to_scale))) = (unit(&from), unit(&to))
    else {
        return err(ErrorValue::NA);
    };
    if from_family != to_family {
        return err(ErrorValue::NA);
    }
    if from_family == TEMPERATURE {
        return match to_celsius(&from, value).and_then(|c| from_celsius(&to, c)) {
            Some(out) => finite(out),
            None => err(ErrorValue::NA),
        };
    }
    finite(value * from_scale / to_scale)
}

const TEMPERATURE: u8 = 4;

/// a unit's family and its size in the family's base unit, after any SI
/// prefix. The base units are the metre, the gram, the second and the litre.
fn unit(name: &str) -> Option<(u8, f64)> {
    if let Some(found) = base_unit(name) {
        return Some(found);
    }
    let mut chars = name.chars();
    let head = chars.next()?;
    let rest = chars.as_str();
    let (family, scale) = base_unit(rest)?;
    if !matches!(family, 0..=3) {
        return None;
    }
    Some((family, scale * si_prefix(head)?))
}

fn base_unit(name: &str) -> Option<(u8, f64)> {
    Some(match name {
        // 0: length, in metres
        "m" => (0, 1.0),
        "mi" => (0, 1609.344),
        "Nmi" => (0, 1852.0),
        "in" => (0, 0.0254),
        "ft" => (0, 0.3048),
        "yd" => (0, 0.9144),
        "ang" => (0, 1e-10),
        "Pica" | "pica" => (0, 0.0254 / 72.0),
        // 1: mass, in grams
        "g" => (1, 1.0),
        "sg" => (1, 14593.903),
        "lbm" => (1, 453.59237),
        "u" => (1, 1.660_538_782e-24),
        "ozm" => (1, 28.349523125),
        // 2: time, in seconds
        "sec" | "s" => (2, 1.0),
        "mn" | "min" => (2, 60.0),
        "hr" => (2, 3600.0),
        "day" | "d" => (2, 86400.0),
        "yr" => (2, 31_557_600.0),
        // 3: liquid measure, in litres
        "l" | "L" | "lt" => (3, 1.0),
        "tsp" => (3, 0.004_928_921_593_75),
        "tbs" => (3, 0.014_786_764_781_25),
        "oz" => (3, 0.0295735295625),
        "cup" => (3, 0.2365882365),
        "pt" | "us_pt" => (3, 0.473176473),
        "qt" => (3, 0.946352946),
        "gal" => (3, 3.785411784),
        // 4: temperature, converted rather than scaled
        "C" | "cel" => (TEMPERATURE, 1.0),
        "F" | "fah" => (TEMPERATURE, 1.0),
        "K" | "kel" => (TEMPERATURE, 1.0),
        _ => return None,
    })
}

fn si_prefix(symbol: char) -> Option<f64> {
    Some(match symbol {
        'Y' => 1e24,
        'Z' => 1e21,
        'E' => 1e18,
        'P' => 1e15,
        'T' => 1e12,
        'G' => 1e9,
        'M' => 1e6,
        'k' => 1e3,
        'h' => 1e2,
        'e' => 1e1,
        'd' => 1e-1,
        'c' => 1e-2,
        'm' => 1e-3,
        'u' => 1e-6,
        'n' => 1e-9,
        'p' => 1e-12,
        'f' => 1e-15,
        'a' => 1e-18,
        'z' => 1e-21,
        'y' => 1e-24,
        _ => return None,
    })
}

fn to_celsius(name: &str, value: f64) -> Option<f64> {
    Some(match name {
        "C" | "cel" => value,
        "F" | "fah" => (value - 32.0) * 5.0 / 9.0,
        "K" | "kel" => value - 273.15,
        _ => return None,
    })
}

fn from_celsius(name: &str, value: f64) -> Option<f64> {
    Some(match name {
        "C" | "cel" => value,
        "F" | "fah" => value * 9.0 / 5.0 + 32.0,
        "K" | "kel" => value + 273.15,
        _ => return None,
    })
}
