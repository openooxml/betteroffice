//! financial functions on excel's cash-flow sign convention: money paid out is
//! negative, money received is positive.

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{EvalContext, err};
use crate::parser::Expr;

use super::{finite, nth_number, omitted};

/// FV(rate, nper, pmt, [pv], [type]): the future value of a fixed-rate
/// annuity. `type` 1 pays at the start of each period.
pub(crate) fn fv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 3 || args.len() > 5 {
        return err(ErrorValue::Value);
    }
    let (rate, nper, pmt) = match (
        nth_number(args, ctx, 0),
        nth_number(args, ctx, 1),
        nth_number(args, ctx, 2),
    ) {
        (Ok(r), Ok(n), Ok(p)) => (r, n, p),
        (Err(e), _, _) | (_, Err(e), _) | (_, _, Err(e)) => return err(e),
    };
    let pv = match optional(args, ctx, 3) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let due = match optional(args, ctx, 4) {
        Ok(v) => v != 0.0,
        Err(e) => return err(e),
    };
    if rate == 0.0 {
        return finite(-(pv + pmt * nper));
    }
    let growth = (1.0 + rate).powf(nper);
    let annuity = if due { pmt * (1.0 + rate) } else { pmt };
    finite(-(pv * growth + annuity * (growth - 1.0) / rate))
}

/// a trailing numeric argument, defaulting to zero when omitted or blank.
fn optional(args: &[Expr], ctx: &EvalContext<'_>, index: usize) -> Result<f64, ErrorValue> {
    match args.get(index) {
        Some(arg) if !omitted(arg) => match crate::eval::evaluate(arg, ctx) {
            CellValue::Empty => Ok(0.0),
            value => crate::eval::to_number(&value),
        },
        _ => Ok(0.0),
    }
}

/// NPV(rate, value1, ...): cash flows discounted from the end of period one.
/// Referenced cells contribute only their numbers; a value written into the
/// call coerces.
pub(crate) fn npv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 2 {
        return err(ErrorValue::Value);
    }
    let rate = match nth_number(args, ctx, 0) {
        Ok(rate) => rate,
        Err(e) => return err(e),
    };
    if rate == -1.0 {
        return err(ErrorValue::Div0);
    }
    let flows = match super::collect_numbers(&args[1..], ctx) {
        Ok(flows) => flows,
        Err(e) => return err(e),
    };
    let mut total = 0.0;
    let mut discount = 1.0;
    for flow in flows {
        discount *= 1.0 + rate;
        total += flow / discount;
    }
    finite(total)
}
