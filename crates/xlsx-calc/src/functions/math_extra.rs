//! math and trig functions past the core set: inverse hyperbolic and
//! reciprocal trig, roman and radix numerals, the significance roundings,
//! combinatorics, paired-array sums, matrices and the 30/360 day count.

use xlsx_model::{CellValue, ErrorValue};

use crate::array::{self, Value};
use crate::eval::{EvalContext, err, evaluate, text, to_text};
use crate::fmath;
use crate::parser::Expr;

use super::{
    argument_cells, coerce_number, collect_numbers, datetime, finite, nth_number, omitted,
};

/// trig arguments at or past 2^27 in magnitude are `#NUM!`.
const TRIG_LIMIT: f64 = 134_217_728.0;
/// 2^53: BASE refuses a number at or past it.
const WHOLE_LIMIT: f64 = 9_007_199_254_740_992.0;
/// the longest text ARABIC and DECIMAL read, and the widest BASE pads to.
const MAX_DIGITS: usize = 255;
/// sqrt(pi), correctly rounded.
const SQRT_PI: f64 = 1.772_453_850_905_516;
/// past this magnitude every double is whole, so a quotient there needs no
/// rounding.
const ALL_WHOLE: f64 = 4_503_599_627_370_496.0;
/// the largest whole number FACTDOUBLE reaches before a double overflows.
const LAST_DOUBLE_FACTORIAL: f64 = 300.0;
/// the largest side a generated block may have.
const MAX_SIDE: f64 = array::MAX_ARRAY_CELLS as f64;

/// ACOSH(number): `#NUM!` below 1.
pub(crate) fn acosh(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, |x| {
        if x < 1.0 {
            return Err(ErrorValue::Num);
        }
        Ok(fmath::acosh(x))
    })
}

pub(crate) fn asinh(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, |x| Ok(fmath::asinh(x)))
}

/// ATANH(number): `#NUM!` unless strictly between -1 and 1.
pub(crate) fn atanh(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, |x| {
        if x.abs() >= 1.0 {
            return Err(ErrorValue::Num);
        }
        Ok(fmath::atanh(x))
    })
}

/// ACOT(number): the arccotangent, in (0, pi).
pub(crate) fn acot(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, |x| Ok(fmath::atan2(1.0, x)))
}

/// ACOTH(number): `#NUM!` unless the magnitude is past 1. read as
/// `ln(1 + 2 / (|x| - 1)) / 2`, so a number near 1 keeps its digits.
pub(crate) fn acoth(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, |x| {
        if x.abs() <= 1.0 {
            return Err(ErrorValue::Num);
        }
        Ok(x.signum() * 0.5 * fmath::ln_1p(2.0 / (x.abs() - 1.0)))
    })
}

/// COT(number): `#DIV/0!` at zero.
pub(crate) fn cot(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    reciprocal(args, ctx, |x| 1.0 / fmath::tan(x), true)
}

/// COTH(number): `#DIV/0!` at zero.
pub(crate) fn coth(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    reciprocal(args, ctx, |x| 1.0 / fmath::tanh(x), true)
}

/// CSC(number): `#DIV/0!` at zero.
pub(crate) fn csc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    reciprocal(args, ctx, |x| 1.0 / fmath::sin(x), true)
}

/// CSCH(number): `#DIV/0!` at zero.
pub(crate) fn csch(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    reciprocal(args, ctx, hyperbolic_cosecant, true)
}

pub(crate) fn sec(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    reciprocal(args, ctx, |x| 1.0 / fmath::cos(x), false)
}

pub(crate) fn sech(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    reciprocal(args, ctx, hyperbolic_secant, false)
}

/// where `cosh` and `sinh` overflow; their reciprocals there are `2 e^-|x|`
/// to the last bit, a subnormal that dividing by infinity would flush to 0.
const HYPERBOLIC_TAIL: f64 = 709.0;

fn hyperbolic_secant(x: f64) -> f64 {
    if x.abs() > HYPERBOLIC_TAIL {
        return 2.0 * fmath::exp(-x.abs());
    }
    1.0 / fmath::cosh(x)
}

fn hyperbolic_cosecant(x: f64) -> f64 {
    if x.abs() > HYPERBOLIC_TAIL {
        return x.signum() * 2.0 * fmath::exp(-x.abs());
    }
    1.0 / fmath::sinh(x)
}

/// a reciprocal trig function: excel refuses a magnitude at or past 2^27 with
/// `#NUM!`, and a function with its pole at zero answers `#DIV/0!` there.
fn reciprocal(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    f: fn(f64) -> f64,
    pole_at_zero: bool,
) -> CellValue {
    unary(args, ctx, |x| {
        if x.abs() >= TRIG_LIMIT {
            return Err(ErrorValue::Num);
        }
        if pole_at_zero && x == 0.0 {
            return Err(ErrorValue::Div0);
        }
        Ok(f(x))
    })
}

/// ARABIC(text): the value of a roman numeral, read case-insensitively with
/// an optional leading minus; surrounding spaces are ignored and empty text
/// is 0. a numeral before a larger one subtracts, so non-canonical forms such
/// as `IIII` still read. a number, a logical, an unknown letter or text past
/// 255 characters is `#VALUE!`.
pub(crate) fn arabic(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 1, 1)?;
        let raw = match evaluate(&args[0], ctx) {
            CellValue::Text { value } => value,
            CellValue::Empty => String::new(),
            CellValue::Error { value } => return Err(value),
            _ => return Err(ErrorValue::Value),
        };
        if raw.chars().count() > MAX_DIGITS {
            return Err(ErrorValue::Value);
        }
        let trimmed = raw.trim_matches(' ');
        let (negative, numeral) = match trimmed.strip_prefix('-') {
            Some(rest) => (true, rest),
            None => (false, trimmed),
        };
        let values = numeral
            .chars()
            .map(numeral_value)
            .collect::<Option<Vec<f64>>>()
            .ok_or(ErrorValue::Value)?;
        let mut total = 0.0;
        for (i, value) in values.iter().enumerate() {
            if values.get(i + 1).is_some_and(|next| next > value) {
                total -= value;
            } else {
                total += value;
            }
        }
        Ok(if negative && total != 0.0 {
            -total
        } else {
            total
        })
    })
}

fn numeral_value(c: char) -> Option<f64> {
    Some(match c.to_ascii_uppercase() {
        'I' => 1.0,
        'V' => 5.0,
        'X' => 10.0,
        'L' => 50.0,
        'C' => 100.0,
        'D' => 500.0,
        'M' => 1000.0,
        _ => return None,
    })
}

/// the numerals ROMAN writes, largest first, each with the least form that
/// allows it. form 0 is the classic notation and every higher form lets a
/// subtrahend sit one numeral further down, so 499 is `CDXCIX` in form 0 and
/// `ID` in form 4.
const ROMAN_NUMERALS: [(u32, &str, u8); 25] = [
    (1000, "M", 0),
    (999, "IM", 4),
    (995, "VM", 3),
    (990, "XM", 2),
    (950, "LM", 1),
    (900, "CM", 0),
    (500, "D", 0),
    (499, "ID", 4),
    (495, "VD", 3),
    (490, "XD", 2),
    (450, "LD", 1),
    (400, "CD", 0),
    (100, "C", 0),
    (99, "IC", 2),
    (95, "VC", 1),
    (90, "XC", 0),
    (50, "L", 0),
    (49, "IL", 2),
    (45, "VL", 1),
    (40, "XL", 0),
    (10, "X", 0),
    (9, "IX", 0),
    (5, "V", 0),
    (4, "IV", 0),
    (1, "I", 0),
];

/// ROMAN(number, [form]): the truncated number as a roman numeral, 0 as
/// empty text; a number outside 0..=3999 is `#VALUE!`. form 0..=4 picks how
/// concise the numeral is, TRUE meaning 0 and FALSE 4; any other form is
/// `#VALUE!`.
pub(crate) fn roman(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let numeral = (|| {
        arity(args, 1, 2)?;
        let number = nth_number(args, ctx, 0)?;
        let form = match args.get(1) {
            Some(arg) if !omitted(arg) => match evaluate(arg, ctx) {
                CellValue::Bool { value } => {
                    if value {
                        0.0
                    } else {
                        4.0
                    }
                }
                value => match coerce_number(&value, ctx)? {
                    form if form < 0.0 => return Err(ErrorValue::Value),
                    form => form.trunc(),
                },
            },
            _ => 0.0,
        };
        if !(0.0..4000.0).contains(&number) || !(0.0..=4.0).contains(&form) {
            return Err(ErrorValue::Value);
        }
        let mut rest = number as u32;
        let mut out = String::new();
        for (value, letters, least_form) in ROMAN_NUMERALS {
            if f64::from(least_form) > form {
                continue;
            }
            while rest >= value {
                out.push_str(letters);
                rest -= value;
            }
        }
        Ok(out)
    })();
    match numeral {
        Ok(numeral) => text(ctx, numeral),
        Err(error) => err(error),
    }
}

/// BASE(number, radix, [min_length]): the truncated number written in radix
/// 2..=36 with uppercase digits, zero-padded to min_length. a negative number
/// or one at or past 2^53, a radix outside 2..=36 and a min_length outside
/// 0..=255 are `#NUM!`.
pub(crate) fn base(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let digits = (|| {
        arity(args, 2, 3)?;
        let number = nth_number(args, ctx, 0)?;
        let radix = nth_number(args, ctx, 1)?.trunc();
        let width = optional(args, ctx, 2, 0.0)?;
        if number < 0.0 || width < 0.0 {
            return Err(ErrorValue::Num);
        }
        let (number, width) = (number.trunc(), width.trunc());
        if !(0.0..WHOLE_LIMIT).contains(&number)
            || !(2.0..=36.0).contains(&radix)
            || !(0.0..=MAX_DIGITS as f64).contains(&width)
        {
            return Err(ErrorValue::Num);
        }
        let radix = radix as u32;
        let mut rest = number as u64;
        let mut digits = Vec::new();
        loop {
            let digit = (rest % u64::from(radix)) as u32;
            digits.push(char::from_digit(digit, radix).map_or('0', |c| c.to_ascii_uppercase()));
            rest /= u64::from(radix);
            if rest == 0 {
                break;
            }
        }
        digits.resize(digits.len().max(width as usize), '0');
        Ok(digits.into_iter().rev().collect::<String>())
    })();
    match digits {
        Ok(digits) => text(ctx, digits),
        Err(error) => err(error),
    }
}

/// DECIMAL(text, radix): the value of text written in radix 2..=36, read
/// case-insensitively; empty text is 0. a radix outside 2..=36 or a character
/// that is not one of its digits is `#NUM!`; a logical or text past 255
/// characters is `#VALUE!`.
pub(crate) fn decimal(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 2, 2)?;
        let digits = match evaluate(&args[0], ctx) {
            CellValue::Bool { .. } => return Err(ErrorValue::Value),
            value => to_text(&value)?,
        };
        let radix = nth_number(args, ctx, 1)?.trunc();
        if digits.chars().count() > MAX_DIGITS {
            return Err(ErrorValue::Value);
        }
        if !(2.0..=36.0).contains(&radix) {
            return Err(ErrorValue::Num);
        }
        digits.chars().try_fold(0.0, |total, c| {
            c.to_digit(radix as u32)
                .map(|digit| total * radix + f64::from(digit))
                .ok_or(ErrorValue::Num)
        })
    })
}

/// CEILING.MATH(number, [significance], [mode]): up to a multiple of
/// significance, 1 by default and its sign ignored. a negative number rounds
/// toward zero unless mode is nonzero, which rounds it away from zero.
pub(crate) fn ceiling_math(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    signed_rounding(args, ctx, f64::ceil, f64::floor)
}

/// FLOOR.MATH(number, [significance], [mode]): down to a multiple of
/// significance, 1 by default and its sign ignored. a negative number rounds
/// away from zero unless mode is nonzero, which rounds it toward zero.
pub(crate) fn floor_math(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    signed_rounding(args, ctx, f64::floor, f64::ceil)
}

fn signed_rounding(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    rule: fn(f64) -> f64,
    negative_mode_rule: fn(f64) -> f64,
) -> CellValue {
    compute(|| {
        arity(args, 1, 3)?;
        let number = nth_number(args, ctx, 0)?;
        let significance = optional(args, ctx, 1, 1.0)?;
        let mode = optional(args, ctx, 2, 0.0)?;
        let rule = if number < 0.0 && mode != 0.0 {
            negative_mode_rule
        } else {
            rule
        };
        Ok(to_multiple(number, significance, rule))
    })
}

/// CEILING.PRECISE(number, [significance]), and ISO.CEILING: toward positive
/// infinity to a multiple of significance, 1 by default and its sign ignored.
pub(crate) fn ceiling_precise(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    precise_rounding(args, ctx, f64::ceil)
}

/// FLOOR.PRECISE(number, [significance]): toward negative infinity to a
/// multiple of significance, 1 by default and its sign ignored.
pub(crate) fn floor_precise(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    precise_rounding(args, ctx, f64::floor)
}

fn precise_rounding(args: &[Expr], ctx: &EvalContext<'_>, rule: fn(f64) -> f64) -> CellValue {
    compute(|| {
        arity(args, 1, 2)?;
        let number = nth_number(args, ctx, 0)?;
        let significance = optional(args, ctx, 1, 1.0)?;
        Ok(to_multiple(number, significance, rule))
    })
}

/// the multiple of `|significance|` that `rule` picks for `number`; zero
/// significance gives 0, and a negative zero comes back as zero.
///
/// a whole significance needs no help: a decimal that is a multiple of it is
/// a whole number, which a double holds exactly below 2^53. a fractional one
/// such as 0.1 is not exact in binary, so 0.3 / 0.1 = 2.9999999999999996;
/// there the quotient is snapped to a whole number within one ulp first, as
/// excel does not lose that multiple either, and only below 2^51 where an ulp
/// is still smaller than any real half. a significance whose reciprocal is
/// whole divides by that reciprocal rather than multiplying, so
/// FLOOR.MATH(0.3, 0.1) is 0.3 and not 0.30000000000000004. a quotient past
/// 2^52 is already whole, so the number is its own multiple, and one that
/// underflows keeps its sign.
fn to_multiple(number: f64, significance: f64, rule: fn(f64) -> f64) -> f64 {
    let significance = significance.abs();
    if significance == 0.0 || number == 0.0 {
        return 0.0;
    }
    let fractional = significance.fract() != 0.0;
    let scale = Some(1.0 / significance)
        .filter(|scale| fractional && *scale > 1.0 && scale.is_finite())
        .and_then(near_whole);
    let quotient = match scale {
        Some(scale) => number * scale,
        None => number / significance,
    };
    if !quotient.is_finite() || quotient.abs() >= ALL_WHOLE {
        return number;
    }
    let quotient = if quotient == 0.0 {
        number.signum() * f64::MIN_POSITIVE
    } else if fractional {
        near_whole(quotient).unwrap_or(quotient)
    } else {
        quotient
    };
    let multiple = rule(quotient);
    let rounded = match scale {
        Some(scale) => multiple / scale,
        None => multiple * significance,
    };
    if rounded == 0.0 { 0.0 } else { rounded }
}

/// the whole number `x` lies within one ulp of, if any, below 2^51.
fn near_whole(x: f64) -> Option<f64> {
    let whole = x.round();
    (whole.abs() < ALL_WHOLE / 2.0 && (x - whole).abs() <= f64::EPSILON * whole.abs())
        .then_some(whole)
}

/// COMBINA(number, chosen): combinations with repetition, that is
/// C(number + chosen - 1, chosen), both truncated. a negative argument,
/// number < chosen (the constraint microsoft documents) and a result past a
/// double are `#NUM!`.
pub(crate) fn combina(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 2, 2)?;
        let number = nth_number(args, ctx, 0)?;
        let chosen = nth_number(args, ctx, 1)?;
        if number < 0.0 || chosen < 0.0 {
            return Err(ErrorValue::Num);
        }
        let (number, chosen) = (number.trunc(), chosen.trunc());
        if number < chosen {
            return Err(ErrorValue::Num);
        }
        binomial_from(number, chosen).ok_or(ErrorValue::Num)
    })
}

/// `first (first + 1) ... (first + k - 1) / k!`, that is
/// C(first + k - 1, k), for whole `first >= k >= 0`. both sides are passed as
/// written, so no count is recovered by subtracting rounded totals; the
/// partial products are whole numbers, exact below 2^53. every factor past
/// the first is at least 1.5, so the loop overflows to `None` within a couple
/// of thousand steps however large `k` is.
fn binomial_from(first: f64, k: f64) -> Option<f64> {
    let mut result = 1.0;
    let steps = k as u64;
    for step in 0..steps {
        let factor = first + step as f64;
        let divisor = (step + 1) as f64;
        let product = result * factor;
        result = if product.is_finite() {
            product / divisor
        } else {
            result / divisor * factor
        };
        if !result.is_finite() {
            return None;
        }
    }
    Some(result.round())
}

/// FACTDOUBLE(number): the double factorial of the truncated number. a
/// negative number is `#NUM!`, checked before truncating as FACT does, and so
/// is a result past a double.
pub(crate) fn factdouble(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, |x| {
        if x < 0.0 || x.trunc() > LAST_DOUBLE_FACTORIAL {
            return Err(ErrorValue::Num);
        }
        let n = x.trunc() as u64;
        Ok((1..=n)
            .rev()
            .step_by(2)
            .map(|factor| factor as f64)
            .product())
    })
}

/// MULTINOMIAL(number1, ...): (n1 + n2 + ...)! / (n1! n2! ...) over the
/// truncated values, built as a product of binomials so only a result past a
/// double overflows. a negative value is `#NUM!`; text or a logical in a
/// range is `#VALUE!`, and typed arguments coerce as numbers do; blank cells
/// are skipped.
pub(crate) fn multinomial(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 1, 255)?;
        let mut total = 0.0;
        let mut result = 1.0;
        for arg in args {
            let (cells, from_sheet) = argument_cells(arg, ctx)?;
            for cell in cells {
                let value = match cell {
                    CellValue::Number { value } => value,
                    CellValue::Error { value } => return Err(value),
                    CellValue::Empty if from_sheet => continue,
                    other if !from_sheet => coerce_number(&other, ctx)?,
                    _ => return Err(ErrorValue::Value),
                };
                if value < 0.0 {
                    return Err(ErrorValue::Num);
                }
                let n = value.trunc();
                let (small, large) = if n < total { (n, total) } else { (total, n) };
                result *= binomial_from(large + 1.0, small).ok_or(ErrorValue::Num)?;
                total += n;
            }
        }
        Ok(result)
    })
}

/// SQRTPI(number): `sqrt(number * pi)`, taken as `sqrt(number) sqrt(pi)` so
/// neither a huge nor a subnormal number loses the product; a negative number
/// is `#NUM!`.
pub(crate) fn sqrtpi(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    unary(args, ctx, |x| {
        if x < 0.0 {
            return Err(ErrorValue::Num);
        }
        Ok(x.sqrt() * SQRT_PI)
    })
}

/// SUMX2MY2(array_x, array_y): the sum of `x^2 - y^2` over paired values,
/// each taken as `(x - y)(x + y)` so equal magnitudes cancel to 0 even where
/// the squares overflow.
pub(crate) fn sumx2my2(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    paired_sum(args, ctx, |x, y| {
        if x.abs() == y.abs() {
            0.0
        } else {
            (x - y) * (x + y)
        }
    })
}

/// SUMX2PY2(array_x, array_y): the sum of `x^2 + y^2` over paired values.
pub(crate) fn sumx2py2(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    paired_sum(args, ctx, |x, y| x * x + y * y)
}

/// SUMXMY2(array_x, array_y): the sum of `(x - y)^2` over paired values.
pub(crate) fn sumxmy2(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    paired_sum(args, ctx, |x, y| (x - y) * (x - y))
}

/// a sum over the positions of two blocks: blocks of different sizes are
/// `#N/A`, an error in either propagates, and a pair where either side is not
/// a number is skipped.
fn paired_sum(args: &[Expr], ctx: &EvalContext<'_>, term: fn(f64, f64) -> f64) -> CellValue {
    compute(|| {
        arity(args, 2, 2)?;
        let xs = block_cells(&args[0], ctx)?;
        let ys = block_cells(&args[1], ctx)?;
        if xs.len() != ys.len() {
            return Err(ErrorValue::NA);
        }
        let mut total = 0.0;
        for pair in xs.iter().zip(&ys) {
            match pair {
                (CellValue::Error { value }, _) | (_, CellValue::Error { value }) => {
                    return Err(*value);
                }
                (CellValue::Number { value: x }, CellValue::Number { value: y }) => {
                    total += term(*x, *y);
                }
                _ => {}
            }
        }
        Ok(total)
    })
}

/// every cell an argument evaluates to, row-major.
fn block_cells(arg: &Expr, ctx: &EvalContext<'_>) -> Result<Vec<CellValue>, ErrorValue> {
    match array::evaluate_array(arg, ctx) {
        Value::Scalar(CellValue::Error { value }) => Err(value),
        Value::Scalar(value) => Ok(vec![value]),
        Value::Array(block) => Ok(block.into_values()),
        Value::Lambda(_) => Err(ErrorValue::Value),
    }
}

/// PERCENTOF(data_subset, data_all): SUM(data_subset) / SUM(data_all), each
/// read the way SUM reads it; a zero total is `#DIV/0!`.
pub(crate) fn percentof(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 2, 2)?;
        let part: f64 = collect_numbers(&args[..1], ctx)?.iter().sum();
        let whole: f64 = collect_numbers(&args[1..], ctx)?.iter().sum();
        if whole == 0.0 {
            return Err(ErrorValue::Div0);
        }
        Ok(part / whole)
    })
}

/// MDETERM(array): the determinant of a square numeric block, by gaussian
/// elimination with partial pivoting as excel computes it, so a singular
/// block can come back as rounding residue: the 1..9 block is 6.66e-16.
pub(crate) fn mdeterm(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 1, 1)?;
        let (n, mut a) = square_matrix(&args[0], ctx)?;
        charge_cube(ctx, n)?;
        let mut det = Scaled::ONE;
        for col in 0..n {
            let pivot = pivot_row(&a, n, col);
            if a[pivot * n + col] == 0.0 {
                return Ok(0.0);
            }
            if pivot != col {
                swap_rows(&mut a, n, pivot, col);
                det.negate();
            }
            let p = a[col * n + col];
            det.multiply(p);
            for row in col + 1..n {
                let factor = a[row * n + col] / p;
                if factor != 0.0 {
                    for k in col + 1..n {
                        a[row * n + k] -= factor * a[col * n + k];
                    }
                }
            }
        }
        Ok(det.value())
    })
}

/// a product kept as `mantissa * 2^exponent`, so a determinant whose running
/// product leaves the double range on the way still comes back when the
/// final value fits.
struct Scaled {
    mantissa: f64,
    exponent: i32,
}

/// how far the mantissa may drift before it is renormalised: 2^512.
const SCALE_STEP: i32 = 512;

impl Scaled {
    const ONE: Self = Scaled {
        mantissa: 1.0,
        exponent: 0,
    };

    fn negate(&mut self) {
        self.mantissa = -self.mantissa;
    }

    fn multiply(&mut self, factor: f64) {
        let step = 2f64.powi(SCALE_STEP);
        self.mantissa *= factor;
        if !self.mantissa.is_finite() || self.mantissa == 0.0 {
            return;
        }
        while self.mantissa.abs() > step {
            self.mantissa /= step;
            self.exponent += SCALE_STEP;
        }
        while self.mantissa.abs() < 1.0 / step {
            self.mantissa *= step;
            self.exponent -= SCALE_STEP;
        }
    }

    fn value(&self) -> f64 {
        let step = 2f64.powi(SCALE_STEP);
        let mut value = self.mantissa;
        let mut exponent = self.exponent;
        while exponent >= SCALE_STEP && value.is_finite() {
            value *= step;
            exponent -= SCALE_STEP;
        }
        while exponent <= -SCALE_STEP && value != 0.0 {
            value /= step;
            exponent += SCALE_STEP;
        }
        value * 2f64.powi(exponent)
    }
}

/// MINVERSE(array): the inverse of a square numeric block, by gauss-jordan
/// elimination with partial pivoting; a singular block is `#NUM!`.
pub(crate) fn minverse(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    let matrix = arity(args, 1, 1).and_then(|()| square_matrix(&args[0], ctx));
    let (n, a) = match matrix {
        Ok(matrix) => matrix,
        Err(error) => return Value::Scalar(err(error)),
    };
    array::number_block(ctx, n, n, cube(n), || inverse(n, a))
}

fn inverse(n: usize, mut a: Vec<f64>) -> Result<Vec<f64>, ErrorValue> {
    let mut inv = identity(n);
    for col in 0..n {
        let pivot = pivot_row(&a, n, col);
        if a[pivot * n + col] == 0.0 {
            return Err(ErrorValue::Num);
        }
        if pivot != col {
            swap_rows(&mut a, n, pivot, col);
            swap_rows(&mut inv, n, pivot, col);
        }
        let p = a[col * n + col];
        for k in 0..n {
            a[col * n + k] /= p;
            inv[col * n + k] /= p;
        }
        for row in 0..n {
            let factor = a[row * n + col];
            if row == col || factor == 0.0 {
                continue;
            }
            for k in 0..n {
                a[row * n + k] -= factor * a[col * n + k];
                inv[row * n + k] -= factor * inv[col * n + k];
            }
        }
    }
    Ok(inv)
}

/// MUNIT(dimension): the identity block of the truncated dimension; below 1
/// is `#VALUE!`.
pub(crate) fn munit(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    let dimension = arity(args, 1, 1).and_then(|()| {
        let n = nth_number(args, ctx, 0)?.trunc();
        if !(1.0..=MAX_SIDE).contains(&n) {
            return Err(ErrorValue::Value);
        }
        Ok(n as usize)
    });
    match dimension {
        Ok(n) => array::number_block(ctx, n, n, 0, || Ok(identity(n))),
        Err(error) => Value::Scalar(err(error)),
    }
}

/// RANDARRAY([rows], [columns], [min], [max], [whole_number]): a block of
/// draws from the seeded stream RAND reads, uniform over [min, max) or, for
/// whole numbers, over ceil(min)..=floor(max). sizes default to 1 and the
/// bounds to 0 and 1. a size that truncates to 0 is `#CALC!`; a negative
/// size, min > max, or a whole-number range holding no whole number is
/// `#VALUE!`.
pub(crate) fn randarray(args: &[Expr], ctx: &EvalContext<'_>) -> Value {
    let spec = (|| {
        arity(args, 0, 5)?;
        let rows = block_side(args, ctx, 0)?;
        let cols = block_side(args, ctx, 1)?;
        let min = optional(args, ctx, 2, 0.0)?;
        let max = optional(args, ctx, 3, 1.0)?;
        let whole = optional(args, ctx, 4, 0.0)? != 0.0;
        let (low, high) = if whole {
            (min.ceil(), max.floor())
        } else {
            (min, max)
        };
        if min > max || low > high {
            return Err(ErrorValue::Value);
        }
        Ok((rows, cols, low, high, whole))
    })();
    let (rows, cols, low, high, whole) = match spec {
        Ok(spec) => spec,
        Err(error) => return Value::Scalar(err(error)),
    };
    array::number_block(ctx, rows, cols, 0, || {
        Ok((0..rows * cols)
            .map(|_| {
                let unit = ctx.next_random_unit();
                if whole {
                    let span = high - low + 1.0;
                    if span.is_finite() {
                        low + (unit * span).floor().min(high - low)
                    } else {
                        between(low, high, unit).floor().clamp(low, high)
                    }
                } else {
                    between(low, high, unit)
                }
            })
            .collect())
    })
}

/// the point `unit` of the way from `low` to `high`, weighted so a span past
/// the largest double does not overflow.
fn between(low: f64, high: f64, unit: f64) -> f64 {
    ((1.0 - unit) * low + unit * high).clamp(low, high)
}

/// a RANDARRAY size: 1 when omitted, truncated, `#CALC!` at 0 and `#VALUE!`
/// below it or past the largest block.
fn block_side(args: &[Expr], ctx: &EvalContext<'_>, index: usize) -> Result<usize, ErrorValue> {
    let side = optional(args, ctx, index, 1.0)?;
    if side < 0.0 {
        return Err(ErrorValue::Value);
    }
    let side = side.trunc();
    if side == 0.0 {
        return Err(ErrorValue::Calc);
    }
    if !(1.0..=MAX_SIDE).contains(&side) {
        return Err(ErrorValue::Value);
    }
    Ok(side as usize)
}

/// a square block of numbers, row-major. any cell that is not a number is
/// `#VALUE!` (an error cell propagates), and so is a block that is not square.
fn square_matrix(arg: &Expr, ctx: &EvalContext<'_>) -> Result<(usize, Vec<f64>), ErrorValue> {
    let (rows, cols, cells) = match array::evaluate_array(arg, ctx) {
        Value::Scalar(value) => (1, 1, vec![value]),
        Value::Array(block) => (block.rows(), block.cols(), block.into_values()),
        Value::Lambda(_) => return Err(ErrorValue::Value),
    };
    let values = cells
        .into_iter()
        .map(|cell| match cell {
            CellValue::Number { value } => Ok(value),
            CellValue::Error { value } => Err(value),
            _ => Err(ErrorValue::Value),
        })
        .collect::<Result<Vec<f64>, ErrorValue>>()?;
    if rows != cols {
        return Err(ErrorValue::Value);
    }
    Ok((rows, values))
}

/// the row at or below `col` holding the largest magnitude in that column,
/// the first such on a tie.
fn pivot_row(a: &[f64], n: usize, col: usize) -> usize {
    let mut best = col;
    for row in col + 1..n {
        if a[row * n + col].abs() > a[best * n + col].abs() {
            best = row;
        }
    }
    best
}

fn swap_rows(a: &mut [f64], n: usize, i: usize, j: usize) {
    for k in 0..n {
        a.swap(i * n + k, j * n + k);
    }
}

fn identity(n: usize) -> Vec<f64> {
    (0..n * n)
        .map(|i| if i % (n + 1) == 0 { 1.0 } else { 0.0 })
        .collect()
}

/// the inner-loop steps an elimination over an `n` by `n` block takes.
fn cube(n: usize) -> u64 {
    (n as u64).saturating_pow(3)
}

fn charge_cube(ctx: &EvalContext<'_>, n: usize) -> Result<(), ErrorValue> {
    if ctx.consume_cells(array::step_units(cube(n))) {
        Ok(())
    } else {
        Err(ErrorValue::Num)
    }
}

/// DAYS360(start_date, end_date, [method]): the days between two dates on a
/// 360-day year, negative when the end comes first. the US (NASD) method, the
/// default, moves a start on the last day of its month to the 30th, february
/// included, and an end on the 31st to the 30th only when the start is then
/// the 30th; otherwise an end on the 31st counts as the 1st of the next
/// month. the European method moves either date from the 31st to the 30th.
pub(crate) fn days360(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 2, 3)?;
        let (y1, m1, mut d1) = date_argument(args, ctx, 0)?;
        let (y2, m2, mut d2) = date_argument(args, ctx, 1)?;
        let european = optional(args, ctx, 2, 0.0)? != 0.0;
        if european {
            d1 = d1.min(30);
            d2 = d2.min(30);
        } else {
            if d1 == 31 || (m1 == 2 && d1 == february_days(y1)) {
                d1 = 30;
            }
            if d2 == 31 && d1 == 30 {
                d2 = 30;
            }
        }
        Ok(((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) as f64)
    })
}

/// the days excel gives february, counting its phantom 1900-02-29.
fn february_days(year: i64) -> i64 {
    let leap = (year % 4 == 0 && year % 100 != 0) || year % 400 == 0;
    if leap || year == 1900 { 29 } else { 28 }
}

/// a date argument as (year, month, day) on excel's 1900 calendar: a serial in
/// the workbook's date system, truncated, or text read as DATEVALUE reads it.
/// a negative serial or one past 9999-12-31 is `#NUM!`.
fn date_argument(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
) -> Result<(i64, i64, i64), ErrorValue> {
    let value = evaluate(&args[index], ctx);
    let parsed = match &value {
        CellValue::Text { value } => datetime::parse_date_text(value, ctx),
        _ => None,
    };
    let serial = match parsed {
        Some(serial) => serial,
        None => {
            let serial = ctx.number(&value)?;
            if serial < 0.0 {
                return Err(ErrorValue::Num);
            }
            let serial = serial.trunc();
            serial as i64
        }
    };
    datetime::serial_to_ymd(datetime::to_1900(serial, ctx.date_system)).ok_or(ErrorValue::Num)
}

fn compute(body: impl FnOnce() -> Result<f64, ErrorValue>) -> CellValue {
    match body() {
        Ok(value) => finite(value),
        Err(error) => err(error),
    }
}

fn unary(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    f: impl FnOnce(f64) -> Result<f64, ErrorValue>,
) -> CellValue {
    compute(|| {
        arity(args, 1, 1)?;
        f(nth_number(args, ctx, 0)?)
    })
}

fn arity(args: &[Expr], min: usize, max: usize) -> Result<(), ErrorValue> {
    if (min..=max).contains(&args.len()) {
        Ok(())
    } else {
        Err(ErrorValue::Value)
    }
}

/// an optional numeric argument; absent or written as a gap, it is `default`.
fn optional(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
    default: f64,
) -> Result<f64, ErrorValue> {
    match args.get(index) {
        Some(arg) if !omitted(arg) => nth_number(args, ctx, index),
        _ => Ok(default),
    }
}
