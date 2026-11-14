//! EUROCONVERT: amounts between the euro and the currencies it replaced, at
//! the conversion rates fixed when each country joined.

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{EvalContext, err, evaluate, to_bool, to_text};
use crate::fmath;
use crate::parser::Expr;

use super::math::round_number;
use super::{finite, nth_number, omitted};

/// (iso code, units per euro, calculation precision) for the currencies
/// Microsoft's reference lists.
const CURRENCIES: [(&str, f64, i32); 14] = [
    ("BEF", 40.3399, 0),
    ("LUF", 40.3399, 0),
    ("DEM", 1.95583, 2),
    ("ESP", 166.386, 0),
    ("FRF", 6.55957, 2),
    ("IEP", 0.787564, 2),
    ("ITL", 1936.27, 0),
    ("NLG", 2.20371, 2),
    ("ATS", 13.7603, 2),
    ("PTE", 200.482, 0),
    ("FIM", 5.94573, 2),
    ("GRD", 340.75, 0),
    ("SIT", 239.64, 2),
    ("EUR", 1.0, 2),
];

/// the most decimals a triangulation rounds to, about all a double holds.
const MAX_TRIANGULATION: f64 = 15.0;

/// `value` rounded to `digits` decimals. a value too large to carry that many
/// decimals is already whole at that precision and is kept as it is, so the
/// rounding cannot overflow.
fn rounded(value: f64, digits: i32) -> f64 {
    let scaled = value * fmath::powi(10.0, digits);
    if scaled.is_finite() && scaled.abs() < 9_007_199_254_740_992.0 {
        round_number(value, digits)
    } else {
        value
    }
}

/// the table row of an iso code, in any letter case; anything else is
/// #VALUE!.
fn currency(value: &CellValue) -> Result<(&'static str, f64, i32), ErrorValue> {
    let code = to_text(value)?;
    CURRENCIES
        .iter()
        .find(|(iso, _, _)| iso.eq_ignore_ascii_case(&code))
        .copied()
        .ok_or(ErrorValue::Value)
}

/// EUROCONVERT(number, source, target, [full_precision],
/// [triangulation_precision]). a member amount becomes euros first, rounded
/// to `triangulation_precision` decimals when one is given (3 or more), and
/// the result is rounded to the target's calculation precision unless
/// `full_precision` is TRUE. the same source and target return `number`
/// unchanged.
pub(crate) fn euroconvert(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let compute = || -> Result<CellValue, ErrorValue> {
        if !(3..=5).contains(&args.len()) {
            return Err(ErrorValue::Value);
        }
        let number = nth_number(args, ctx, 0)?;
        let (source, from_rate, _) = currency(&evaluate(&args[1], ctx))?;
        let (target, to_rate, to_precision) = currency(&evaluate(&args[2], ctx))?;
        let full = match args.get(3) {
            Some(arg) if !omitted(arg) => to_bool(&evaluate(arg, ctx))?,
            _ => false,
        };
        let triangulation = match args.get(4) {
            Some(arg) if !omitted(arg) => {
                let digits = nth_number(args, ctx, 4)?.trunc();
                if digits < 3.0 {
                    return Err(ErrorValue::Value);
                }
                Some(digits.min(MAX_TRIANGULATION) as i32)
            }
            _ => None,
        };
        if source == target {
            return Ok(finite(number));
        }
        let converted = if source == "EUR" {
            number * to_rate
        } else {
            let euros = number / from_rate;
            let euros = triangulation.map_or(euros, |digits| rounded(euros, digits));
            euros * to_rate
        };
        Ok(finite(if full {
            converted
        } else {
            rounded(converted, to_precision)
        }))
    };
    compute().unwrap_or_else(err)
}
