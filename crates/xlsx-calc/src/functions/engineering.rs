//! engineering functions: base conversions, bitwise arithmetic, complex
//! numbers written as text, and bessel functions.

use std::f64::consts::{EULER_GAMMA, FRAC_2_PI, LN_2, LN_10, PI};

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{EvalContext, err, evaluate, format_number, text};
use crate::fmath;
use crate::parser::Expr;

use super::{argument_cells, coerce_number, finite, omitted};

/// the most digits a converted number may have, in any base.
const MAX_DIGITS: usize = 10;
/// the bitwise functions take integers below 2^48.
const BIT_LIMIT: f64 = 281_474_976_710_656.0;
/// the largest shift BITLSHIFT and BITRSHIFT accept either way.
const MAX_SHIFT: f64 = 53.0;
/// the most arguments IMSUM and IMPRODUCT take.
const MAX_ARGS: usize = 255;
/// below this bessel J and Y come from their power series.
const SERIES_BELOW: f64 = 2.0;
/// from here on bessel J and Y come from hankel's asymptotic expansion.
const HANKEL_FROM: f64 = 25.0;
/// the highest bessel order, the largest 32-bit integer.
const MAX_ORDER: f64 = 2_147_483_647.0;
/// series past this many orders take their leading factor from stirling's
/// series instead of a product.
const PRODUCT_ORDERS: usize = 100_000;
/// the power of two series rescale by to stay inside the double range.
const RESCALE_BITS: i32 = 500;
/// below this natural logarithm a result is lost to underflow.
const UNDERFLOW: f64 = -750.0;
/// loops longer than this spend the recalculation budget.
const LONG_LOOP: usize = 4096;

/// a base excel converts to and from: ten digits holding a two's-complement
/// number of `bits` bits.
#[derive(Clone, Copy)]
struct Radix {
    base: u32,
    bits: u32,
}

const BINARY: Radix = Radix { base: 2, bits: 10 };
const OCTAL: Radix = Radix { base: 8, bits: 30 };
const HEXADECIMAL: Radix = Radix { base: 16, bits: 40 };

impl Radix {
    /// values lie in `-limit..limit`.
    fn limit(self) -> i64 {
        1 << (self.bits - 1)
    }

    /// up to ten digits, the top bit of a full width number meaning negative.
    /// letters may be either case; anything else is `#NUM!`.
    fn read(self, digits: &str) -> Result<i64, ErrorValue> {
        if digits.len() > MAX_DIGITS {
            return Err(ErrorValue::Num);
        }
        let mut value: i64 = 0;
        for ch in digits.chars() {
            let digit = ch.to_digit(self.base).ok_or(ErrorValue::Num)?;
            value = value * i64::from(self.base) + i64::from(digit);
        }
        if value >= self.limit() {
            value -= 1 << self.bits;
        }
        Ok(value)
    }

    /// `value` in this base: a negative one as its ten-digit two's complement,
    /// a positive one padded to `places` when given. `places` is truncated
    /// and must be 1 to 10 either way, though a negative value ignores it;
    /// too many digits for `places` is `#NUM!`.
    fn write(self, value: i64, places: Option<f64>) -> Result<String, ErrorValue> {
        if value < -self.limit() || value >= self.limit() {
            return Err(ErrorValue::Num);
        }
        let places = places.map(f64::trunc);
        let valid = |places: f64| (1.0..=MAX_DIGITS as f64).contains(&places);
        if !places.is_none_or(valid) {
            return Err(ErrorValue::Num);
        }
        if value < 0 {
            return Ok(digits_in(value + (1 << self.bits), self.base));
        }
        let digits = digits_in(value, self.base);
        let Some(places) = places else {
            return Ok(digits);
        };
        if digits.len() > places as usize {
            return Err(ErrorValue::Num);
        }
        Ok(format!("{digits:0>width$}", width = places as usize))
    }
}

/// a non-negative integer's digits, uppercase.
fn digits_in(mut value: i64, base: u32) -> String {
    let mut out = Vec::new();
    loop {
        let digit = (value % i64::from(base)) as u32;
        out.push(
            char::from_digit(digit, base)
                .unwrap_or('0')
                .to_ascii_uppercase(),
        );
        value /= i64::from(base);
        if value == 0 {
            break;
        }
    }
    out.iter().rev().collect()
}

/// BIN2DEC(number): a binary number of up to ten digits as decimal.
pub(crate) fn bin2dec(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    to_decimal(args, ctx, BINARY)
}

/// BIN2HEX(number, [places]).
pub(crate) fn bin2hex(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    convert(args, ctx, BINARY, HEXADECIMAL)
}

/// BIN2OCT(number, [places]).
pub(crate) fn bin2oct(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    convert(args, ctx, BINARY, OCTAL)
}

/// DEC2BIN(number, [places]): -512 to 511.
pub(crate) fn dec2bin(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    from_decimal(args, ctx, BINARY)
}

/// DEC2HEX(number, [places]): -549755813888 to 549755813887.
pub(crate) fn dec2hex(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    from_decimal(args, ctx, HEXADECIMAL)
}

/// DEC2OCT(number, [places]): -536870912 to 536870911.
pub(crate) fn dec2oct(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    from_decimal(args, ctx, OCTAL)
}

/// HEX2BIN(number, [places]).
pub(crate) fn hex2bin(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    convert(args, ctx, HEXADECIMAL, BINARY)
}

/// HEX2DEC(number).
pub(crate) fn hex2dec(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    to_decimal(args, ctx, HEXADECIMAL)
}

/// HEX2OCT(number, [places]).
pub(crate) fn hex2oct(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    convert(args, ctx, HEXADECIMAL, OCTAL)
}

/// OCT2BIN(number, [places]).
pub(crate) fn oct2bin(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    convert(args, ctx, OCTAL, BINARY)
}

/// OCT2DEC(number).
pub(crate) fn oct2dec(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    to_decimal(args, ctx, OCTAL)
}

/// OCT2HEX(number, [places]).
pub(crate) fn oct2hex(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    convert(args, ctx, OCTAL, HEXADECIMAL)
}

fn to_decimal(args: &[Expr], ctx: &EvalContext<'_>, from: Radix) -> CellValue {
    compute(|| {
        arity(args, 1, 1)?;
        Ok(from.read(&digits_arg(args, ctx, 0)?)? as f64)
    })
}

fn convert(args: &[Expr], ctx: &EvalContext<'_>, from: Radix, to: Radix) -> CellValue {
    written_text(ctx, || {
        arity(args, 1, 2)?;
        let digits = digits_arg(args, ctx, 0)?;
        let places = optional_number(args, ctx, 1)?;
        to.write(from.read(&digits)?, places)
    })
}

/// the number is truncated first.
fn from_decimal(args: &[Expr], ctx: &EvalContext<'_>, to: Radix) -> CellValue {
    written_text(ctx, || {
        arity(args, 1, 2)?;
        let value = number_arg(args, ctx, 0)?.trunc();
        let places = optional_number(args, ctx, 1)?;
        let limit = to.limit() as f64;
        if value < -limit || value >= limit {
            return Err(ErrorValue::Num);
        }
        to.write(value as i64, places)
    })
}

/// the digits a conversion reads: text as written, a number as its decimal
/// digits, a blank as no digits. a logical is `#VALUE!`.
fn digits_arg(args: &[Expr], ctx: &EvalContext<'_>, index: usize) -> Result<String, ErrorValue> {
    match evaluate(&args[index], ctx) {
        CellValue::Text { value } => Ok(value),
        CellValue::Number { value } => Ok(format_number(value)),
        CellValue::Empty => Ok(String::new()),
        CellValue::Bool { .. } => Err(ErrorValue::Value),
        CellValue::Error { value } => Err(value),
    }
}

/// BITAND(number1, number2).
pub(crate) fn bitand(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    bitwise(args, ctx, |a, b| a & b)
}

/// BITOR(number1, number2).
pub(crate) fn bitor(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    bitwise(args, ctx, |a, b| a | b)
}

/// BITXOR(number1, number2).
pub(crate) fn bitxor(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    bitwise(args, ctx, |a, b| a ^ b)
}

/// BITLSHIFT(number, shift_amount): a negative amount shifts right.
pub(crate) fn bitlshift(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    shift(args, ctx, 1.0)
}

/// BITRSHIFT(number, shift_amount): a negative amount shifts left.
pub(crate) fn bitrshift(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    shift(args, ctx, -1.0)
}

fn bitwise(args: &[Expr], ctx: &EvalContext<'_>, op: fn(u64, u64) -> u64) -> CellValue {
    compute(|| {
        arity(args, 2, 2)?;
        let a = bit_operand(args, ctx, 0)?;
        let b = bit_operand(args, ctx, 1)?;
        Ok(op(a, b) as f64)
    })
}

/// the amount is truncated and may be at most 53 either way. a left shift
/// is not held to 48 bits: only its operand is.
fn shift(args: &[Expr], ctx: &EvalContext<'_>, direction: f64) -> CellValue {
    compute(|| {
        arity(args, 2, 2)?;
        let value = bit_operand(args, ctx, 0)?;
        let amount = number_arg(args, ctx, 1)?.trunc() * direction;
        if amount.abs() > MAX_SHIFT {
            return Err(ErrorValue::Num);
        }
        Ok(if amount >= 0.0 {
            value as f64 * 2f64.powi(amount as i32)
        } else {
            (value >> (-amount) as u32) as f64
        })
    })
}

/// an integer from 0 to 2^48-1; anything else that is a number is `#NUM!`.
fn bit_operand(args: &[Expr], ctx: &EvalContext<'_>, index: usize) -> Result<u64, ErrorValue> {
    let value = number_arg(args, ctx, index)?;
    if !(0.0..BIT_LIMIT).contains(&value) || value.fract() != 0.0 {
        return Err(ErrorValue::Num);
    }
    Ok(value as u64)
}

/// DELTA(number1, [number2]): 1 when the numbers are equal, else 0.
pub(crate) fn delta(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 1, 2)?;
        let a = number_arg(args, ctx, 0)?;
        let b = optional_number(args, ctx, 1)?.unwrap_or(0.0);
        Ok(if a == b { 1.0 } else { 0.0 })
    })
}

/// GESTEP(number, [step]): 1 when the number is at least `step`, else 0.
pub(crate) fn gestep(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 1, 2)?;
        let value = number_arg(args, ctx, 0)?;
        let step = optional_number(args, ctx, 1)?.unwrap_or(0.0);
        Ok(if value >= step { 1.0 } else { 0.0 })
    })
}

/// a complex number.
#[derive(Clone, Copy, Debug, PartialEq)]
struct Complex {
    re: f64,
    im: f64,
}

/// the imaginary unit a complex text was written with, `i` or `j`; none for
/// a plain real number.
type Suffix = Option<char>;

impl Complex {
    const ZERO: Complex = Complex { re: 0.0, im: 0.0 };

    fn new(re: f64, im: f64) -> Self {
        Self { re, im }
    }

    fn is_zero(self) -> bool {
        self.re == 0.0 && self.im == 0.0
    }

    fn add(self, other: Self) -> Self {
        Self::new(self.re + other.re, self.im + other.im)
    }

    fn sub(self, other: Self) -> Self {
        Self::new(self.re - other.re, self.im - other.im)
    }

    fn mul(self, other: Self) -> Self {
        Self::new(
            self.re * other.re - self.im * other.im,
            self.re * other.im + self.im * other.re,
        )
    }

    fn div(self, other: Self) -> Option<Self> {
        let scale = other.re * other.re + other.im * other.im;
        if scale == 0.0 {
            return None;
        }
        Some(Self::new(
            (self.re * other.re + self.im * other.im) / scale,
            (self.im * other.re - self.re * other.im) / scale,
        ))
    }

    fn abs(self) -> f64 {
        hypot(self.re, self.im)
    }

    fn arg(self) -> f64 {
        fmath::atan2(self.im, self.re)
    }

    /// z^n through the polar form, as excel takes it: `IMPOWER("2+3i",3)`
    /// is `-46+9.00000000000001i`. zero only takes a positive power.
    fn powf(self, n: f64) -> Option<Self> {
        if self.is_zero() {
            return (n > 0.0).then_some(Self::ZERO);
        }
        let radius = fmath::pow(self.abs(), n);
        let angle = n * self.arg();
        Some(Self::new(
            radius * fmath::cos(angle),
            radius * fmath::sin(angle),
        ))
    }

    /// the principal root through the polar form, so `IMSQRT("-4")` keeps the
    /// cosine's residue: `1.22464679914735E-16+2i`.
    fn sqrt(self) -> Self {
        let radius = self.abs().sqrt();
        let angle = self.arg() / 2.0;
        Self::new(radius * fmath::cos(angle), radius * fmath::sin(angle))
    }

    fn exp(self) -> Self {
        let scale = fmath::exp(self.re);
        Self::new(scale * fmath::cos(self.im), scale * fmath::sin(self.im))
    }

    fn ln(self) -> Self {
        Self::new(fmath::ln(self.abs()), self.arg())
    }

    fn sin(self) -> Self {
        let (a, b) = (self.re, self.im);
        Self::new(
            fmath::sin(a) * fmath::cosh(b),
            fmath::cos(a) * fmath::sinh(b),
        )
    }

    fn cos(self) -> Self {
        let (a, b) = (self.re, self.im);
        Self::new(
            fmath::cos(a) * fmath::cosh(b),
            -fmath::sin(a) * fmath::sinh(b),
        )
    }

    fn sinh(self) -> Self {
        let (a, b) = (self.re, self.im);
        Self::new(
            fmath::sinh(a) * fmath::cos(b),
            fmath::cosh(a) * fmath::sin(b),
        )
    }

    fn cosh(self) -> Self {
        let (a, b) = (self.re, self.im);
        Self::new(
            fmath::cosh(a) * fmath::cos(b),
            fmath::sinh(a) * fmath::sin(b),
        )
    }

    /// the reciprocal trigonometric forms over the double angle, as excel
    /// takes them: `IMSEC("4+3i")` is `-0.0652940278579471-0.0752249603027732i`
    /// where the plain reciprocal of the cosine prints `-0.065294027857947`.
    fn tan(self) -> Option<Self> {
        let (a, b) = (self.re, self.im);
        over(
            fmath::sin(2.0 * a),
            fmath::sinh(2.0 * b),
            fmath::cos(2.0 * a) + fmath::cosh(2.0 * b),
        )
    }

    fn cot(self) -> Option<Self> {
        let (a, b) = (self.re, self.im);
        over(
            fmath::sin(2.0 * a),
            -fmath::sinh(2.0 * b),
            fmath::cosh(2.0 * b) - fmath::cos(2.0 * a),
        )
    }

    fn sec(self) -> Option<Self> {
        let (a, b) = (self.re, self.im);
        over(
            2.0 * fmath::cos(a) * fmath::cosh(b),
            2.0 * fmath::sin(a) * fmath::sinh(b),
            fmath::cos(2.0 * a) + fmath::cosh(2.0 * b),
        )
    }

    fn csc(self) -> Option<Self> {
        let (a, b) = (self.re, self.im);
        over(
            2.0 * fmath::sin(a) * fmath::cosh(b),
            -2.0 * fmath::cos(a) * fmath::sinh(b),
            fmath::cosh(2.0 * b) - fmath::cos(2.0 * a),
        )
    }

    fn sech(self) -> Option<Self> {
        let (a, b) = (self.re, self.im);
        over(
            2.0 * fmath::cosh(a) * fmath::cos(b),
            -2.0 * fmath::sinh(a) * fmath::sin(b),
            fmath::cosh(2.0 * a) + fmath::cos(2.0 * b),
        )
    }

    fn csch(self) -> Option<Self> {
        let (a, b) = (self.re, self.im);
        over(
            2.0 * fmath::sinh(a) * fmath::cos(b),
            -2.0 * fmath::cosh(a) * fmath::sin(b),
            fmath::cosh(2.0 * a) - fmath::cos(2.0 * b),
        )
    }
}

/// `(re + im i) / scale`, none when the scale is zero.
fn over(re: f64, im: f64, scale: f64) -> Option<Complex> {
    (scale != 0.0).then(|| Complex::new(re / scale, im / scale))
}

fn hypot(x: f64, y: f64) -> f64 {
    #[cfg(target_arch = "wasm32")]
    {
        x.hypot(y)
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        libm::hypot(x, y)
    }
}

/// read a complex number the way excel writes one: `3+4i`, `-2.5j`, `i`,
/// `-i`, `1E2-3i`, or a plain real. the unit is a trailing lowercase `i` or
/// `j`; spaces and anything else make it unreadable. empty text is zero.
fn parse_complex(s: &str) -> Option<(Complex, Suffix)> {
    let suffix = match s.as_bytes().last() {
        None => return Some((Complex::ZERO, None)),
        Some(b'i') => 'i',
        Some(b'j') => 'j',
        Some(_) => return real_number(s).map(|re| (Complex::new(re, 0.0), None)),
    };
    let body = &s[..s.len() - 1];
    let bytes = body.as_bytes();
    let split = (1..bytes.len())
        .rev()
        .find(|&k| matches!(bytes[k], b'+' | b'-') && !matches!(bytes[k - 1], b'e' | b'E'));
    let (re, coefficient) = match split {
        Some(k) => (real_number(&body[..k])?, &body[k..]),
        None => (0.0, body),
    };
    let im = match coefficient {
        "" | "+" => 1.0,
        "-" => -1.0,
        digits => real_number(digits)?,
    };
    Some((Complex::new(re, im), Some(suffix)))
}

/// a signed decimal with an optional exponent, nothing else.
fn real_number(s: &str) -> Option<f64> {
    let bytes = s.as_bytes();
    let mut at = usize::from(matches!(bytes.first(), Some(b'+' | b'-')));
    let mut digits = 0;
    let mut point = false;
    while let Some(&b) = bytes.get(at) {
        match b {
            b'0'..=b'9' => digits += 1,
            b'.' if !point => point = true,
            _ => break,
        }
        at += 1;
    }
    if digits == 0 {
        return None;
    }
    if matches!(bytes.get(at), Some(b'e' | b'E')) {
        at += 1;
        if matches!(bytes.get(at), Some(b'+' | b'-')) {
            at += 1;
        }
        let start = at;
        while bytes.get(at).is_some_and(u8::is_ascii_digit) {
            at += 1;
        }
        if at == start {
            return None;
        }
    }
    if at != bytes.len() {
        return None;
    }
    s.parse().ok().filter(|value: &f64| value.is_finite())
}

/// a complex number the way excel writes one: each part to 15 significant
/// digits, a zero part left out, a unit coefficient written as the bare
/// unit, and zero itself as `0`.
fn format_complex(z: Complex, unit: char) -> String {
    let mut out = String::new();
    if z.re != 0.0 {
        out.push_str(&significant(z.re));
    }
    if z.im != 0.0 {
        if z.im < 0.0 {
            out.push('-');
        } else if !out.is_empty() {
            out.push('+');
        }
        let magnitude = significant(z.im.abs());
        if magnitude != "1" {
            out.push_str(&magnitude);
        }
        out.push(unit);
    }
    if out.is_empty() {
        out.push('0');
    }
    out
}

/// `x` rounded to 15 significant digits without trailing zeros, in plain
/// notation when its decimal exponent is from -4 to 14 and as `1.5E-16` or
/// `1E+15` outside that.
fn significant(x: f64) -> String {
    if x == 0.0 {
        return "0".to_string();
    }
    let scientific = format!("{:.14e}", x.abs());
    let (mantissa, exponent) = scientific.split_once('e').unwrap_or((&scientific, "0"));
    let exponent: i32 = exponent.parse().unwrap_or(0);
    let digits: String = mantissa.chars().filter(|c| *c != '.').collect();
    let digits = digits.trim_end_matches('0');
    let sign = if x < 0.0 { "-" } else { "" };
    if !(-4..15).contains(&exponent) {
        let (head, tail) = digits.split_at(1);
        let point = if tail.is_empty() { "" } else { "." };
        let exponent_sign = if exponent < 0 { '-' } else { '+' };
        return format!(
            "{sign}{head}{point}{tail}E{exponent_sign}{:02}",
            exponent.abs()
        );
    }
    if exponent < 0 {
        let zeros = "0".repeat((-exponent - 1) as usize);
        return format!("{sign}0.{zeros}{digits}");
    }
    let whole = exponent as usize + 1;
    if digits.len() <= whole {
        format!("{sign}{digits:0<whole$}")
    } else {
        format!("{sign}{}.{}", &digits[..whole], &digits[whole..])
    }
}

/// a value as a complex number: text is parsed (`#NUM!` when unreadable), a
/// number is real, a blank is zero and a logical is `#VALUE!`.
fn complex_value(value: CellValue) -> Result<(Complex, Suffix), ErrorValue> {
    match value {
        CellValue::Text { value } => parse_complex(&value).ok_or(ErrorValue::Num),
        CellValue::Number { value } => Ok((Complex::new(value, 0.0), None)),
        CellValue::Empty => Ok((Complex::ZERO, None)),
        CellValue::Bool { .. } => Err(ErrorValue::Value),
        CellValue::Error { value } => Err(value),
    }
}

fn complex_arg(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
) -> Result<(Complex, Suffix), ErrorValue> {
    complex_value(evaluate(&args[index], ctx))
}

/// the unit two operands share; `i` against `j` is `#VALUE!`.
fn agree(a: Suffix, b: Suffix) -> Result<Suffix, ErrorValue> {
    match (a, b) {
        (Some(x), Some(y)) if x != y => Err(ErrorValue::Value),
        (Some(x), _) | (None, Some(x)) => Ok(Some(x)),
        (None, None) => Ok(None),
    }
}

/// a complex result written out in its operands' unit, `i` by default; a
/// part that overflowed is `#NUM!`.
fn written(
    ctx: &EvalContext<'_>,
    body: impl FnOnce() -> Result<(Complex, Suffix), ErrorValue>,
) -> CellValue {
    match body() {
        Ok((z, _)) if !z.re.is_finite() || !z.im.is_finite() => err(ErrorValue::Num),
        Ok((z, suffix)) => text(ctx, format_complex(z, suffix.unwrap_or('i'))),
        Err(error) => err(error),
    }
}

/// COMPLEX(real_num, i_num, [suffix]): the suffix is `i` or `j`, lowercase;
/// empty text means `i`.
pub(crate) fn complex(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    written(ctx, || {
        arity(args, 2, 3)?;
        let re = number_arg(args, ctx, 0)?;
        let im = number_arg(args, ctx, 1)?;
        let suffix = match args.get(2) {
            Some(arg) if !omitted(arg) => match evaluate(arg, ctx) {
                CellValue::Text { value } => match value.as_str() {
                    "" | "i" => 'i',
                    "j" => 'j',
                    _ => return Err(ErrorValue::Value),
                },
                CellValue::Empty => 'i',
                CellValue::Error { value } => return Err(value),
                _ => return Err(ErrorValue::Value),
            },
            _ => 'i',
        };
        Ok((Complex::new(re, im), Some(suffix)))
    })
}

/// IMABS(inumber): the modulus.
pub(crate) fn imabs(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_number(args, ctx, |z| Ok(z.abs()))
}

/// IMAGINARY(inumber): the imaginary coefficient.
pub(crate) fn imaginary(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_number(args, ctx, |z| Ok(z.im))
}

/// IMREAL(inumber): the real coefficient.
pub(crate) fn imreal(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_number(args, ctx, |z| Ok(z.re))
}

/// IMARGUMENT(inumber): the angle in radians, -pi to pi; zero has none and is
/// `#DIV/0!`.
pub(crate) fn imargument(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_number(args, ctx, |z| {
        if z.is_zero() {
            Err(ErrorValue::Div0)
        } else {
            Ok(z.arg())
        }
    })
}

/// IMCONJUGATE(inumber).
pub(crate) fn imconjugate(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, |z| Some(Complex::new(z.re, -z.im)))
}

/// IMCOS(inumber).
pub(crate) fn imcos(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, |z| Some(z.cos()))
}

/// IMCOSH(inumber).
pub(crate) fn imcosh(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, |z| Some(z.cosh()))
}

/// IMCOT(inumber): `#NUM!` at the sine's zeros.
pub(crate) fn imcot(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, Complex::cot)
}

/// IMCSC(inumber): `#NUM!` at the sine's zeros.
pub(crate) fn imcsc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, Complex::csc)
}

/// IMCSCH(inumber): `#NUM!` at zero.
pub(crate) fn imcsch(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, Complex::csch)
}

/// IMEXP(inumber).
pub(crate) fn imexp(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, |z| Some(z.exp()))
}

/// IMLN(inumber): the principal logarithm; zero is `#NUM!`.
pub(crate) fn imln(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, |z| Some(z.ln()))
}

/// IMLOG10(inumber).
pub(crate) fn imlog10(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, |z| {
        Some(Complex::new(fmath::log10(z.abs()), z.arg() / LN_10))
    })
}

/// IMLOG2(inumber).
pub(crate) fn imlog2(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, |z| {
        Some(Complex::new(fmath::ln(z.abs()) / LN_2, z.arg() / LN_2))
    })
}

/// IMSEC(inumber).
pub(crate) fn imsec(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, Complex::sec)
}

/// IMSECH(inumber).
pub(crate) fn imsech(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, Complex::sech)
}

/// IMSIN(inumber).
pub(crate) fn imsin(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, |z| Some(z.sin()))
}

/// IMSINH(inumber).
pub(crate) fn imsinh(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, |z| Some(z.sinh()))
}

/// IMSQRT(inumber).
pub(crate) fn imsqrt(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, |z| Some(z.sqrt()))
}

/// IMTAN(inumber).
pub(crate) fn imtan(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_unary(args, ctx, Complex::tan)
}

/// IMDIV(inumber1, inumber2): dividing by zero is `#NUM!`.
pub(crate) fn imdiv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_binary(args, ctx, |a, b| a.div(b))
}

/// IMSUB(inumber1, inumber2).
pub(crate) fn imsub(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_binary(args, ctx, |a, b| Some(a.sub(b)))
}

/// IMPOWER(inumber, number): any real power, through the polar form.
pub(crate) fn impower(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    written(ctx, || {
        arity(args, 2, 2)?;
        let (z, suffix) = complex_arg(args, ctx, 0)?;
        let n = number_arg(args, ctx, 1)?;
        Ok((z.powf(n).ok_or(ErrorValue::Num)?, suffix))
    })
}

/// IMSUM(inumber1, [inumber2], ...): ranges and arrays add every cell; blank
/// cells add nothing.
pub(crate) fn imsum(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_fold(args, ctx, Complex::ZERO, Complex::add)
}

/// IMPRODUCT(inumber1, [inumber2], ...): ranges and arrays multiply every
/// cell; blank cells are skipped.
pub(crate) fn improduct(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    im_fold(args, ctx, Complex::new(1.0, 0.0), Complex::mul)
}

fn im_number(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    f: impl FnOnce(Complex) -> Result<f64, ErrorValue>,
) -> CellValue {
    compute(|| {
        arity(args, 1, 1)?;
        let (z, _) = complex_arg(args, ctx, 0)?;
        f(z).map(|value| value + 0.0)
    })
}

fn im_unary(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    f: impl FnOnce(Complex) -> Option<Complex>,
) -> CellValue {
    written(ctx, || {
        arity(args, 1, 1)?;
        let (z, suffix) = complex_arg(args, ctx, 0)?;
        Ok((f(z).ok_or(ErrorValue::Num)?, suffix))
    })
}

fn im_binary(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    f: impl FnOnce(Complex, Complex) -> Option<Complex>,
) -> CellValue {
    written(ctx, || {
        arity(args, 2, 2)?;
        let (a, first) = complex_arg(args, ctx, 0)?;
        let (b, second) = complex_arg(args, ctx, 1)?;
        let suffix = agree(first, second)?;
        Ok((f(a, b).ok_or(ErrorValue::Num)?, suffix))
    })
}

fn im_fold(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    start: Complex,
    step: fn(Complex, Complex) -> Complex,
) -> CellValue {
    written(ctx, || {
        arity(args, 1, MAX_ARGS)?;
        let mut total = start;
        let mut suffix = None;
        for arg in args {
            let (cells, from_sheet) = argument_cells(arg, ctx)?;
            for cell in cells {
                if from_sheet && matches!(cell, CellValue::Empty) {
                    continue;
                }
                let (z, unit) = complex_value(cell)?;
                suffix = agree(suffix, unit)?;
                total = step(total, z);
            }
        }
        Ok((total, suffix))
    })
}

/// BESSELI(x, n): the modified bessel function of the first kind. `n` is
/// truncated and must not be negative.
pub(crate) fn besseli(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    bessel(args, ctx, false, |x, n| {
        let value = power_series(x.abs(), n, 1.0, ctx)?;
        Ok(if x < 0.0 && n % 2 == 1 { -value } else { value })
    })
}

/// BESSELJ(x, n): the bessel function of the first kind. `n` is truncated
/// and must not be negative.
pub(crate) fn besselj(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    bessel(args, ctx, false, |x, n| {
        let value = first_kind(x.abs(), n, ctx)?;
        Ok(if x < 0.0 && n % 2 == 1 { -value } else { value })
    })
}

/// BESSELK(x, n): the modified bessel function of the second kind, for
/// positive `x`.
pub(crate) fn besselk(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    bessel(args, ctx, true, |x, n| {
        let (k0, k1) = modified_second_low(x);
        recur_up(k0, k1, n, ctx, |k, low, high| low + 2.0 * k / x * high)
    })
}

/// BESSELY(x, n): the bessel function of the second kind, for positive `x`.
pub(crate) fn bessely(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    bessel(args, ctx, true, |x, n| {
        let (y0, y1) = second_kind_low(x);
        recur_up(y0, y1, n, ctx, |k, low, high| 2.0 * k / x * high - low)
    })
}

fn bessel(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    positive: bool,
    f: impl FnOnce(f64, usize) -> Result<f64, ErrorValue>,
) -> CellValue {
    compute(|| {
        arity(args, 2, 2)?;
        let x = number_arg(args, ctx, 0)?;
        let n = number_arg(args, ctx, 1)?.trunc();
        if !(0.0..=MAX_ORDER).contains(&n) || (positive && x <= 0.0) {
            return Err(ErrorValue::Num);
        }
        f(x, n as usize)
    })
}

/// the order `n` member of a sequence that grows by `step(k, f(k-1), f(k))`
/// from `f(0)` and `f(1)`, the stable direction for Y and K.
fn recur_up(
    f0: f64,
    f1: f64,
    n: usize,
    ctx: &EvalContext<'_>,
    step: impl Fn(f64, f64, f64) -> f64,
) -> Result<f64, ErrorValue> {
    if n == 0 {
        return Ok(f0);
    }
    charge(ctx, n)?;
    let (mut low, mut high) = (f0, f1);
    for k in 1..n {
        if !high.is_finite() {
            break;
        }
        (low, high) = (high, step(k as f64, low, high));
    }
    Ok(high)
}

/// J_n(x) for x >= 0: the power series below 2; past hankel's threshold and
/// below the turning point it recurs up from J_0 and J_1; elsewhere it runs
/// miller's backward recurrence.
fn first_kind(x: f64, n: usize, ctx: &EvalContext<'_>) -> Result<f64, ErrorValue> {
    if x < SERIES_BELOW {
        return power_series(x, n, -1.0, ctx);
    }
    if x >= HANKEL_FROM && (n as f64) < x {
        let (j0, _) = hankel(0, x);
        let (j1, _) = hankel(1, x);
        return recur_up(j0, j1, n, ctx, |k, low, high| 2.0 * k / x * high - low);
    }
    if n as f64 > x && n as f64 * fmath::ln(x / 2.0) - ln_factorial(n) < UNDERFLOW {
        return Ok(0.0);
    }
    let top = miller_start(n, x);
    charge(ctx, top)?;
    let mut wanted = 0.0;
    let mut sum = 0.0;
    let mut above = 0.0;
    let mut here = 1.0;
    for k in (0..=top).rev() {
        if k == n {
            wanted = here;
        }
        if k == 0 {
            sum += here;
            break;
        }
        if k % 2 == 0 {
            sum += 2.0 * here;
        }
        (above, here) = (here, 2.0 * k as f64 / x * here - above);
        if here.abs() > 1e250 {
            here *= 1e-250;
            above *= 1e-250;
            sum *= 1e-250;
            wanted *= 1e-250;
        }
    }
    Ok(wanted / sum)
}

/// J_0..J_top at `x` by miller's backward recurrence, normalized by
/// J_0 + 2(J_2 + J_4 + ...) = 1.
fn first_kind_table(x: f64, top: usize) -> Vec<f64> {
    let mut values = vec![0.0; top + 1];
    let mut above = 0.0;
    values[top] = 1.0;
    for k in (1..=top).rev() {
        let below = 2.0 * k as f64 / x * values[k] - above;
        above = values[k];
        values[k - 1] = below;
        if below.abs() > 1e250 {
            above *= 1e-250;
            for value in &mut values[k - 1..] {
                *value *= 1e-250;
            }
        }
    }
    let sum: f64 = values[0] + 2.0 * values.iter().skip(2).step_by(2).sum::<f64>();
    values.iter().map(|value| value / sum).collect()
}

/// an even order high enough above both `n` and `x` that the backward
/// recurrence has settled long before it reaches them.
fn miller_start(n: usize, x: f64) -> usize {
    let top = (n as f64).max(x);
    let start = (top + 20.0 + (60.0 * top).sqrt()) as usize;
    start.saturating_add(start % 2)
}

/// (J_order(x), Y_order(x)) for order 0 or 1 from hankel's asymptotic
/// expansion, summed until its terms stop shrinking; exact to rounding for
/// x >= 25.
fn hankel(order: u32, x: f64) -> (f64, f64) {
    let mu = f64::from(4 * order * order);
    let (mut p, mut q, mut term) = (1.0, 0.0, 1.0_f64);
    for k in 1..200 {
        let odd = f64::from(2 * k - 1);
        let next = term * (mu - odd * odd) / (8.0 * f64::from(k) * x);
        if k > 1 && next.abs() >= term.abs() {
            break;
        }
        term = next;
        let signed = if (k / 2) % 2 == 0 { term } else { -term };
        if k % 2 == 1 {
            q += signed;
        } else {
            p += signed;
        }
        if term.abs() < 1e-17 {
            break;
        }
    }
    let (s, c) = (fmath::sin(x), fmath::cos(x));
    let (cos_chi, sin_chi) = if order == 0 {
        (c + s, s - c)
    } else {
        (s - c, -(s + c))
    };
    let scale = 1.0 / (PI * x).sqrt();
    (
        scale * (p * cos_chi - q * sin_chi),
        scale * (p * sin_chi + q * cos_chi),
    )
}

/// (Y_0(x), Y_1(x)) for x > 0: their power series below 2, hankel's
/// expansion from 25 on, and between them neumann's series over miller's J
/// values.
fn second_kind_low(x: f64) -> (f64, f64) {
    if x < SERIES_BELOW {
        return second_kind_series(x);
    }
    if x >= HANKEL_FROM {
        return (hankel(0, x).1, hankel(1, x).1);
    }
    let j = first_kind_table(x, miller_start(1, x));
    let log = fmath::ln(x / 2.0) + EULER_GAMMA;
    let (mut even, mut odd) = (0.0, 0.0);
    for k in 1..(j.len() - 1) / 2 {
        let sign = if k % 2 == 1 { -1.0 } else { 1.0 };
        even += sign * j[2 * k] / k as f64;
        odd += sign * (j[2 * k - 1] - j[2 * k + 1]) / k as f64;
    }
    let y0 = FRAC_2_PI * (log * j[0] - 2.0 * even);
    let y1 = FRAC_2_PI * (log * j[1] + odd - j[0] / x);
    (y0, y1)
}

/// the sum over k of sign^k (x/2)^(2k+n) / (k! (n+k)!) for x >= 0: I_n(x)
/// with a sign of 1, every term positive, and J_n(x) with -1, used where x is
/// small enough that its alternating terms barely cancel. the leading factor
/// (x/2)^n / n! and the series each carry a power-of-two exponent, so neither
/// under- or overflows on its own: BESSELI(6000,9000) is about 1.93E+25.
fn power_series(x: f64, n: usize, sign: f64, ctx: &EvalContext<'_>) -> Result<f64, ErrorValue> {
    if x == 0.0 {
        return Ok(if n == 0 { 1.0 } else { 0.0 });
    }
    let half = x / 2.0;
    let quarter = half * half;
    let rescale = 2f64.powi(RESCALE_BITS);
    charge(ctx, x as usize)?;
    let (mut sum, mut term, mut sum_exponent) = (1.0_f64, 1.0_f64, 0);
    let mut k = 0.0;
    let order = n as f64;
    loop {
        let ratio = sign * quarter / ((k + 1.0) * (order + k + 1.0));
        term *= ratio;
        sum += term;
        k += 1.0;
        if sum.abs() > rescale {
            sum /= rescale;
            term /= rescale;
            sum_exponent += RESCALE_BITS;
        }
        if (term.abs() < 1e-17 * sum.abs() && ratio.abs() < 1.0) || !sum.is_finite() {
            break;
        }
    }
    if n > PRODUCT_ORDERS {
        let lead = order * fmath::ln(half) - ln_factorial(n);
        let series = fmath::ln(sum.abs()) + f64::from(sum_exponent) * LN_2;
        return Ok(sum.signum() * fmath::exp(lead + series));
    }
    charge(ctx, n)?;
    let (mut lead, mut lead_exponent) = (1.0, 0);
    for k in 1..=n {
        lead *= half / k as f64;
        if lead < 1.0 / rescale {
            lead *= rescale;
            lead_exponent -= RESCALE_BITS;
        } else if lead > rescale {
            lead /= rescale;
            lead_exponent += RESCALE_BITS;
        }
    }
    Ok(times_power_of_two(lead * sum, lead_exponent + sum_exponent))
}

/// `value` times 2^`exponent`, for an exponent that is a multiple of
/// [`RESCALE_BITS`].
fn times_power_of_two(mut value: f64, mut exponent: i32) -> f64 {
    let rescale = 2f64.powi(RESCALE_BITS);
    while exponent > 0 && value.is_finite() {
        value *= rescale;
        exponent -= RESCALE_BITS;
    }
    while exponent < 0 && value != 0.0 {
        value /= rescale;
        exponent += RESCALE_BITS;
    }
    value
}

/// (Y_0(x), Y_1(x)) from the ascending series, with harmonic numbers H_k:
/// Y_0 = 2/pi ((ln(x/2) + g) J_0 - sum H_k a_k) and
/// Y_1 = -2/(pi x) + 2/pi ln(x/2) J_1 - x/(2 pi) sum (2 H_k - 2g + 1/(k+1)) b_k,
/// where a_k = (-x^2/4)^k / k!^2 and b_k = (-x^2/4)^k / (k! (k+1)!).
fn second_kind_series(x: f64) -> (f64, f64) {
    let q = -x * x / 4.0;
    let (mut a, mut b, mut harmonic) = (1.0_f64, 1.0_f64, 0.0);
    let (mut j0, mut s0, mut j1, mut s1) = (0.0, 0.0, 0.0, 0.0);
    for k in 1..100 {
        j0 += a;
        s0 += harmonic * a;
        j1 += b;
        s1 += (2.0 * (harmonic - EULER_GAMMA) + 1.0 / k as f64) * b;
        if a.abs() < 1e-18 && b.abs() < 1e-18 {
            break;
        }
        let k = k as f64;
        harmonic += 1.0 / k;
        a *= q / (k * k);
        b *= q / (k * (k + 1.0));
    }
    let log = fmath::ln(x / 2.0);
    let y0 = FRAC_2_PI * ((log + EULER_GAMMA) * j0 - s0);
    let y1 = -FRAC_2_PI / x + FRAC_2_PI * log * j1 * x / 2.0 - x / (2.0 * PI) * s1;
    (y0, y1)
}

/// ln(n!), by stirling's series above 100.
fn ln_factorial(n: usize) -> f64 {
    if n <= 100 {
        return (2..=n).map(|k| fmath::ln(k as f64)).sum();
    }
    let n = n as f64;
    let inverse = 1.0 / n;
    let square = inverse * inverse;
    n * fmath::ln(n) - n
        + 0.5 * fmath::ln(2.0 * PI * n)
        + inverse * (1.0 / 12.0 - square * (1.0 / 360.0 - square / 1260.0))
}

/// (K_0(x), K_1(x)) for x > 0: their ascending series below 2, and from 2
/// on K_v(x) = integral over t > 0 of exp(-x cosh t) cosh(v t). the
/// trapezoid rule converges geometrically on it; the step keeps the aliasing
/// error under e^-45.
fn modified_second_low(x: f64) -> (f64, f64) {
    if x < SERIES_BELOW {
        return modified_second_series(x);
    }
    let h = 0.2_f64.min(2.0 * PI / (90.0 * x).sqrt());
    let mut k0 = 0.5;
    let mut k1 = 0.5;
    for j in 1..100_000 {
        let t = j as f64 * h;
        let half = fmath::sinh(t / 2.0);
        let decay = fmath::exp(-2.0 * x * half * half);
        let f1 = decay * fmath::cosh(t);
        k0 += decay;
        k1 += f1;
        if f1 < 1e-18 * k1 && decay < 1e-18 * k0 {
            break;
        }
    }
    let scale = fmath::exp(-x) * h;
    (scale * k0, scale * k1)
}

/// (K_0(x), K_1(x)) from the ascending series, with harmonic numbers H_k:
/// K_0 = sum H_k a_k - (ln(x/2) + g) I_0 and
/// K_1 = 1/x + ln(x/2) I_1 - x/4 sum (2 H_k - 2g + 1/(k+1)) b_k,
/// where a_k = (x^2/4)^k / k!^2 and b_k = (x^2/4)^k / (k! (k+1)!). nothing
/// in it overflows before the answer does: BESSELK(1E-307,1) is 1E+307.
fn modified_second_series(x: f64) -> (f64, f64) {
    let q = x * x / 4.0;
    let (mut a, mut b, mut harmonic) = (1.0_f64, 1.0_f64, 0.0);
    let (mut i0, mut s0, mut i1, mut s1) = (0.0, 0.0, 0.0, 0.0);
    for k in 1..100 {
        i0 += a;
        s0 += harmonic * a;
        i1 += b;
        s1 += (2.0 * (harmonic - EULER_GAMMA) + 1.0 / k as f64) * b;
        if a < 1e-18 * i0 && b < 1e-18 * i1 {
            break;
        }
        let k = k as f64;
        harmonic += 1.0 / k;
        a *= q / (k * k);
        b *= q / (k * (k + 1.0));
    }
    let log = fmath::ln(x / 2.0);
    let k0 = s0 - (log + EULER_GAMMA) * i0;
    let k1 = 1.0 / x + log * i1 * x / 2.0 - x / 4.0 * s1;
    (k0, k1)
}

/// run a function body and show what it computed.
fn compute(body: impl FnOnce() -> Result<f64, ErrorValue>) -> CellValue {
    match body() {
        Ok(value) => finite(value),
        Err(error) => err(error),
    }
}

/// text a conversion produced, or the error that stopped it.
fn written_text(
    ctx: &EvalContext<'_>,
    body: impl FnOnce() -> Result<String, ErrorValue>,
) -> CellValue {
    match body() {
        Ok(value) => text(ctx, value),
        Err(error) => err(error),
    }
}

fn arity(args: &[Expr], min: usize, max: usize) -> Result<(), ErrorValue> {
    if (min..=max).contains(&args.len()) {
        Ok(())
    } else {
        Err(ErrorValue::Value)
    }
}

/// an argument as a number: text that reads as one is coerced, a blank is
/// zero, and a logical is `#VALUE!` as everywhere in these functions.
fn number_arg(args: &[Expr], ctx: &EvalContext<'_>, index: usize) -> Result<f64, ErrorValue> {
    match evaluate(&args[index], ctx) {
        CellValue::Bool { .. } => Err(ErrorValue::Value),
        value => coerce_number(&value, ctx),
    }
}

/// a trailing argument, none when absent or written as a gap.
fn optional_number(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
) -> Result<Option<f64>, ErrorValue> {
    match args.get(index) {
        Some(arg) if !omitted(arg) => number_arg(args, ctx, index).map(Some),
        _ => Ok(None),
    }
}

/// spend the recalculation budget on a loop of `steps` iterations when it is
/// long.
fn charge(ctx: &EvalContext<'_>, steps: usize) -> Result<(), ErrorValue> {
    if steps <= LONG_LOOP || ctx.consume_cells(steps as u64) {
        Ok(())
    } else {
        Err(ErrorValue::Num)
    }
}
