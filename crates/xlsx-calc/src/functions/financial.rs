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

/// PMT(rate, nper, pv, [fv], [type]): the level payment that settles a loan.
/// negative, because the payment leaves the borrower.
pub(crate) fn pmt(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 3 || args.len() > 5 {
        return err(ErrorValue::Value);
    }
    let (rate, nper, pv) = match (
        nth_number(args, ctx, 0),
        nth_number(args, ctx, 1),
        nth_number(args, ctx, 2),
    ) {
        (Ok(r), Ok(n), Ok(p)) => (r, n, p),
        (Err(e), _, _) | (_, Err(e), _) | (_, _, Err(e)) => return err(e),
    };
    let fv = match optional(args, ctx, 3) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let due = match optional(args, ctx, 4) {
        Ok(v) => v != 0.0,
        Err(e) => return err(e),
    };
    if nper == 0.0 {
        return err(ErrorValue::Num);
    }
    if rate == 0.0 {
        return finite(-(pv + fv) / nper);
    }
    let growth = (1.0 + rate).powf(nper);
    let timing = if due { 1.0 + rate } else { 1.0 };
    finite(-(pv * growth + fv) * rate / ((growth - 1.0) * timing))
}

/// PV(rate, nper, pmt, [fv], [type]): what a fixed-rate annuity is worth now.
pub(crate) fn pv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
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
    let fv = match optional(args, ctx, 3) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let due = match optional(args, ctx, 4) {
        Ok(v) => v != 0.0,
        Err(e) => return err(e),
    };
    if rate == 0.0 {
        return finite(-(fv + pmt * nper));
    }
    let growth = (1.0 + rate).powf(nper);
    let annuity = if due { pmt * (1.0 + rate) } else { pmt };
    finite(-(fv + annuity * (growth - 1.0) / rate) / growth)
}

/// NPER(rate, pmt, pv, [fv], [type]): how many periods the annuity runs for.
pub(crate) fn nper(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 3 || args.len() > 5 {
        return err(ErrorValue::Value);
    }
    let (rate, pmt, pv) = match (
        nth_number(args, ctx, 0),
        nth_number(args, ctx, 1),
        nth_number(args, ctx, 2),
    ) {
        (Ok(r), Ok(p), Ok(v)) => (r, p, v),
        (Err(e), _, _) | (_, Err(e), _) | (_, _, Err(e)) => return err(e),
    };
    let fv = match optional(args, ctx, 3) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let due = match optional(args, ctx, 4) {
        Ok(v) => v != 0.0,
        Err(e) => return err(e),
    };
    if rate == 0.0 {
        if pmt == 0.0 {
            return err(ErrorValue::Num);
        }
        return finite(-(pv + fv) / pmt);
    }
    let annuity = if due { pmt * (1.0 + rate) } else { pmt };
    let numerator = annuity - fv * rate;
    let denominator = pv * rate + annuity;
    if numerator / denominator <= 0.0 {
        return err(ErrorValue::Num);
    }
    finite((numerator / denominator).ln() / (1.0 + rate).ln())
}

/// NPV(rate, value, ...): the present value of cash flows one period apart,
/// the first of them one period out.
pub(crate) fn npv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 2 {
        return err(ErrorValue::Value);
    }
    let rate = match nth_number(args, ctx, 0) {
        Ok(r) => r,
        Err(e) => return err(e),
    };
    if rate == -1.0 {
        return err(ErrorValue::Num);
    }
    let flows = match super::collect_numbers_deep(&args[1..], ctx) {
        Ok(flows) => flows,
        Err(e) => return err(e),
    };
    let mut discount = 1.0;
    let mut total = 0.0;
    for flow in flows {
        discount *= 1.0 + rate;
        total += flow / discount;
    }
    finite(total)
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
