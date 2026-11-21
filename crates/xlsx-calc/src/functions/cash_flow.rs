//! cash-flow financial functions: rates of return, loan schedules and
//! depreciation, on excel's sign convention where money paid out is negative.

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{EvalContext, as_area, err, evaluate};
use crate::fmath;
use crate::parser::Expr;

use super::{argument_cells, collect_numbers_deep, finite, nth_number, omitted};

/// how many secant steps RATE takes before it gives up.
const RATE_STEPS: usize = 200;
/// the step size and residual RATE accepts as converged.
const RATE_TOLERANCE: f64 = 1e-7;
/// how many newton steps IRR and XIRR take before they give up.
const RETURN_STEPS: usize = 100;
/// the last serial excel reads as a date, 9999-12-31.
const LAST_DATE: f64 = 2_958_465.0;

/// RATE(nper, pmt, pv, [fv], [type], [guess]): the per-period rate of an
/// annuity. excel runs a secant search from `guess` and `guess + 0.001`; it
/// stops once a step is under 1e-7 and the balance it leaves is too, and after
/// 200 steps it still answers when the balance is that small. so a zero rate
/// comes back as the search's tiny residue, as excel's does. `nper` must be
/// positive.
pub(crate) fn rate(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 3, 6)?;
        let nper = nth_number(args, ctx, 0)?;
        let pmt = optional(args, ctx, 1, 0.0)?;
        let pv = nth_number(args, ctx, 2)?;
        let fv = optional(args, ctx, 3, 0.0)?;
        let due = optional(args, ctx, 4, 0.0)? != 0.0;
        let guess = optional(args, ctx, 5, 0.1)?;
        if nper <= 0.0 {
            return Err(ErrorValue::Num);
        }
        let balance = |rate: f64| annuity_balance(rate, nper, pmt, pv, fv, due);
        let residual = RATE_TOLERANCE.max(1e-14 * (pv.abs() + fv.abs() + (pmt * nper).abs()));
        let (mut x0, mut x1) = (guess, guess + 0.001);
        let mut f0 = balance(x0);
        let mut f1 = balance(x1);
        let mut x2 = x1;
        let mut f2 = f1;
        for _ in 0..RATE_STEPS {
            if !f1.is_finite() || f1 == f0 {
                return Err(ErrorValue::Num);
            }
            x2 = x1 - f1 * (x1 - x0) / (f1 - f0);
            f2 = balance(x2);
            if !x2.is_finite() || !f2.is_finite() {
                return Err(ErrorValue::Num);
            }
            if (x2 - x1).abs() < RATE_TOLERANCE && f2.abs() < residual {
                return Ok(x2);
            }
            (x0, f0, x1, f1) = (x1, f1, x2, f2);
        }
        if f2.abs() < residual {
            Ok(x2)
        } else {
            Err(ErrorValue::Num)
        }
    })
}

/// what is left of an annuity's balance at `rate`: zero at the rate that
/// settles it.
fn annuity_balance(rate: f64, nper: f64, pmt: f64, pv: f64, fv: f64, due: bool) -> f64 {
    if rate == 0.0 {
        return pv + pmt * nper + fv;
    }
    let growth = fmath::pow(1.0 + rate, nper);
    let mut balance = pv * growth;
    if pmt != 0.0 {
        let timing = if due { 1.0 + rate } else { 1.0 };
        balance += pmt * timing * (growth - 1.0) / rate;
    }
    balance + fv
}

/// IRR(values, [guess]): the rate at which the periodic cash flows have zero
/// net present value. text, logicals and blanks in a range are skipped.
pub(crate) fn irr(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 1, 2)?;
        let flows = collect_numbers_deep(&args[..1], ctx)?;
        let guess = optional(args, ctx, 1, 0.1)?;
        if !has_both_signs(&flows) {
            return Err(ErrorValue::Num);
        }
        solve_return(guess, |rate| {
            let mut value = 0.0;
            let mut slope = 0.0;
            let mut discount = 1.0;
            for (period, flow) in flows.iter().enumerate() {
                value += flow * discount;
                discount /= 1.0 + rate;
                slope -= period as f64 * flow * discount;
            }
            (value, slope)
        })
    })
}

/// MIRR(values, finance_rate, reinvest_rate): the return when outflows are
/// financed at one rate and inflows reinvested at another. `#DIV/0!` without
/// both an inflow and an outflow, and for a rate of -100%, where excel's
/// discounting divides by zero.
pub(crate) fn mirr(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 3, 3)?;
        let flows = collect_numbers_deep(&args[..1], ctx)?;
        let finance = nth_number(args, ctx, 1)?;
        let reinvest = nth_number(args, ctx, 2)?;
        if finance == -1.0 || reinvest == -1.0 {
            return Err(ErrorValue::Div0);
        }
        let n = flows.len() as f64;
        let mut inflows = 0.0;
        let mut outflows = 0.0;
        for (period, flow) in flows.iter().enumerate() {
            let t = period as f64;
            if *flow > 0.0 {
                inflows += flow * fmath::pow(1.0 + reinvest, n - 1.0 - t);
            } else if *flow < 0.0 {
                outflows += flow / fmath::pow(1.0 + finance, t);
            }
        }
        if inflows == 0.0 || outflows == 0.0 {
            return Err(ErrorValue::Div0);
        }
        Ok(fmath::pow(-inflows / outflows, 1.0 / (n - 1.0)) - 1.0)
    })
}

/// XNPV(rate, values, dates): the present value of cash flows on arbitrary
/// dates, discounted in 365-day years from the first date.
pub(crate) fn xnpv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 3, 3)?;
        let rate = nth_number(args, ctx, 0)?;
        let (flows, years) = dated_flows(&args[1], &args[2], ctx)?;
        Ok(flows
            .iter()
            .zip(&years)
            .map(|(flow, years)| flow / fmath::pow(1.0 + rate, *years))
            .sum())
    })
}

/// XIRR(values, dates, [guess]): the annual rate at which dated cash flows
/// have zero net present value. excel stops once its estimate moves by less
/// than 1e-8, so its answers sit within about 1e-8 of the root this finds.
pub(crate) fn xirr(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 2, 3)?;
        let (flows, years) = dated_flows(&args[0], &args[1], ctx)?;
        let guess = optional(args, ctx, 2, 0.1)?;
        if !has_both_signs(&flows) {
            return Err(ErrorValue::Num);
        }
        solve_return(guess, |rate| {
            let mut value = 0.0;
            let mut slope = 0.0;
            for (flow, years) in flows.iter().zip(&years) {
                let discounted = flow / fmath::pow(1.0 + rate, *years);
                value += discounted;
                slope -= years * discounted / (1.0 + rate);
            }
            (value, slope)
        })
    })
}

/// the cash flows and their distance from the first date in 365-day years.
/// dates truncate to whole days; a date before the first or a count that
/// differs from the values is `#NUM!`, anything but a number `#VALUE!`.
fn dated_flows(
    values: &Expr,
    dates: &Expr,
    ctx: &EvalContext<'_>,
) -> Result<(Vec<f64>, Vec<f64>), ErrorValue> {
    let flows = positional_numbers(values, ctx)?;
    let days = positional_numbers(dates, ctx)?;
    if flows.0 != days.0 {
        return Err(ErrorValue::Num);
    }
    let mut flows = flows.1;
    let mut days = days.1;
    let len = flows.len().max(days.len());
    flows.resize(len, 0.0);
    days.resize(len, 0.0);
    let mut years = Vec::with_capacity(len);
    let first = days.first().map_or(0.0, |day| day.trunc());
    for day in &days {
        if !(0.0..LAST_DATE + 1.0).contains(day) {
            return Err(ErrorValue::Value);
        }
        let day = day.trunc();
        if day < first {
            return Err(ErrorValue::Num);
        }
        years.push((day - first) / 365.0);
    }
    Ok((flows, years))
}

/// an argument's cells in order, with the count its reference spans. blanks
/// read as zero; trailing cells past the sheet's used extent are left out.
fn positional_numbers(arg: &Expr, ctx: &EvalContext<'_>) -> Result<(usize, Vec<f64>), ErrorValue> {
    let span = as_area(arg, ctx).map(|area| area.rows.saturating_mul(area.cols));
    let (cells, from_sheet) = argument_cells(arg, ctx)?;
    let mut numbers = Vec::with_capacity(cells.len());
    for cell in &cells {
        numbers.push(match cell {
            CellValue::Number { value } => *value,
            CellValue::Empty => 0.0,
            CellValue::Error { value } => return Err(*value),
            other if !from_sheet => ctx.number(other)?,
            _ => return Err(ErrorValue::Value),
        });
    }
    Ok((span.unwrap_or(numbers.len()), numbers))
}

fn has_both_signs(flows: &[f64]) -> bool {
    flows.iter().any(|flow| *flow > 0.0) && flows.iter().any(|flow| *flow < 0.0)
}

/// newton's method for a rate of return from `guess`; a step past -100% is
/// pulled back halfway towards it, so the discount factor stays defined.
fn solve_return(
    guess: f64,
    value_and_slope: impl Fn(f64) -> (f64, f64),
) -> Result<f64, ErrorValue> {
    let mut rate = guess;
    for _ in 0..RETURN_STEPS {
        let (value, slope) = value_and_slope(rate);
        if !value.is_finite() || !slope.is_finite() || slope == 0.0 {
            return Err(ErrorValue::Num);
        }
        let mut next = rate - value / slope;
        if next <= -1.0 {
            next = (rate - 1.0) / 2.0;
        }
        if (next - rate).abs() <= 1e-12 * next.abs().max(1.0) {
            return Ok(next);
        }
        rate = next;
    }
    Err(ErrorValue::Num)
}

/// IPMT(rate, per, nper, pv, [fv], [type]): the interest share of payment
/// `per`.
pub(crate) fn ipmt(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer(payment_split(args, ctx).map(|(_, interest)| interest))
}

/// PPMT(rate, per, nper, pv, [fv], [type]): the principal share of payment
/// `per`.
pub(crate) fn ppmt(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer(payment_split(args, ctx).map(|(payment, interest)| payment - interest))
}

/// the level payment and the interest part of payment `per`; `per` outside
/// `1..=nper` is `#NUM!`.
fn payment_split(args: &[Expr], ctx: &EvalContext<'_>) -> Result<(f64, f64), ErrorValue> {
    arity(args, 4, 6)?;
    let rate = nth_number(args, ctx, 0)?;
    let per = nth_number(args, ctx, 1)?;
    let nper = nth_number(args, ctx, 2)?;
    let pv = nth_number(args, ctx, 3)?;
    let fv = optional(args, ctx, 4, 0.0)?;
    let due = optional(args, ctx, 5, 0.0)? != 0.0;
    if per < 1.0 || per > nper {
        return Err(ErrorValue::Num);
    }
    let payment = level_payment(rate, nper, pv, fv, due)?;
    Ok((payment, interest_part(rate, per, payment, pv, due)))
}

/// the level payment PMT computes.
fn level_payment(rate: f64, nper: f64, pv: f64, fv: f64, due: bool) -> Result<f64, ErrorValue> {
    if nper == 0.0 {
        return Err(ErrorValue::Num);
    }
    if rate == 0.0 {
        return Ok(-(pv + fv) / nper);
    }
    let growth = fmath::pow(1.0 + rate, nper);
    let timing = if due { 1.0 + rate } else { 1.0 };
    Ok(-(pv * growth + fv) * rate / ((growth - 1.0) * timing))
}

/// the value FV computes: the negated balance after `nper` payments.
fn future_value(rate: f64, nper: f64, pmt: f64, pv: f64, due: bool) -> f64 {
    if rate == 0.0 {
        return -(pv + pmt * nper);
    }
    let growth = fmath::pow(1.0 + rate, nper);
    let timing = if due { 1.0 + rate } else { 1.0 };
    -(pv * growth + pmt * timing * (growth - 1.0) / rate)
}

/// the interest in payment `per`. a payment at the start of a period settles
/// the interest of the period before, so the first one carries none.
fn interest_part(rate: f64, per: f64, payment: f64, pv: f64, due: bool) -> f64 {
    if !due {
        return future_value(rate, per - 1.0, payment, pv, false) * rate;
    }
    if per == 1.0 {
        return 0.0;
    }
    (future_value(rate, per - 2.0, payment, pv, true) - payment) * rate
}

/// CUMIPMT(rate, nper, pv, start_period, end_period, type): the interest paid
/// from `start_period` through `end_period`.
pub(crate) fn cumipmt(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer(cumulative(args, ctx, |_, interest| interest))
}

/// CUMPRINC(rate, nper, pv, start_period, end_period, type): the principal
/// repaid from `start_period` through `end_period`.
pub(crate) fn cumprinc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer(cumulative(args, ctx, |payment, interest| {
        payment - interest
    }))
}

/// sum `part(payment, interest)` over the periods. the periods truncate;
/// rate, nper and pv must be positive and type exactly 0 or 1.
fn cumulative(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    part: impl Fn(f64, f64) -> f64,
) -> Result<f64, ErrorValue> {
    arity(args, 6, 6)?;
    let rate = nth_number(args, ctx, 0)?;
    let nper = nth_number(args, ctx, 1)?;
    let pv = nth_number(args, ctx, 2)?;
    let start = nth_number(args, ctx, 3)?.trunc();
    let end = nth_number(args, ctx, 4)?.trunc();
    let kind = nth_number(args, ctx, 5)?;
    if rate <= 0.0 || nper <= 0.0 || pv <= 0.0 {
        return Err(ErrorValue::Num);
    }
    if start < 1.0 || end < start || end > nper || (kind != 0.0 && kind != 1.0) {
        return Err(ErrorValue::Num);
    }
    let due = kind == 1.0;
    charge(ctx, end - start + 1.0)?;
    let payment = level_payment(rate, nper, pv, 0.0, due)?;
    let mut total = 0.0;
    for offset in 0..=(end - start) as u64 {
        let per = start + offset as f64;
        total += part(payment, interest_part(rate, per, payment, pv, due));
    }
    Ok(total)
}

/// ISPMT(rate, per, nper, pv): the interest in period `per` (counted from
/// zero) of a loan repaid in equal principal parts.
pub(crate) fn ispmt(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 4, 4)?;
        let [rate, per, nper, pv] = numbers(args, ctx)?;
        if nper == 0.0 {
            return Err(ErrorValue::Div0);
        }
        Ok(pv * rate * (per / nper - 1.0))
    })
}

/// SLN(cost, salvage, life): straight-line depreciation per period.
pub(crate) fn sln(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 3, 3)?;
        let [cost, salvage, life] = numbers(args, ctx)?;
        if life == 0.0 {
            return Err(ErrorValue::Div0);
        }
        Ok((cost - salvage) / life)
    })
}

/// SYD(cost, salvage, life, per): sum-of-years'-digits depreciation for
/// period `per`.
pub(crate) fn syd(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 4, 4)?;
        let [cost, salvage, life, per] = numbers(args, ctx)?;
        if life <= 0.0 || per <= 0.0 || per > life {
            return Err(ErrorValue::Num);
        }
        Ok((cost - salvage) * (life - per + 1.0) * 2.0 / (life * (life + 1.0)))
    })
}

/// DB(cost, salvage, life, period, [month]): fixed-declining-balance
/// depreciation. the rate is rounded to three decimals, the first year counts
/// `month` months and, when that is short of twelve, the year after `life`
/// takes the rest. a salvage at or above cost leaves nothing to depreciate.
pub(crate) fn db(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 4, 5)?;
        let cost = nth_number(args, ctx, 0)?;
        let salvage = nth_number(args, ctx, 1)?;
        let life = nth_number(args, ctx, 2)?;
        let period = nth_number(args, ctx, 3)?.trunc();
        let month = optional(args, ctx, 4, 12.0)?.trunc();
        if cost < 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0.0 {
            return Err(ErrorValue::Num);
        }
        let last = if month < 12.0 { life + 1.0 } else { life };
        if !(1.0..=12.0).contains(&month) || period > last {
            return Err(ErrorValue::Num);
        }
        if salvage >= cost {
            return Ok(0.0);
        }
        charge(ctx, period)?;
        let rate = ((1.0 - fmath::pow(salvage / cost, 1.0 / life)) * 1000.0).round() / 1000.0;
        let mut depreciation = cost * rate * month / 12.0;
        let mut total = depreciation;
        for year in 2..=period as u64 {
            depreciation = (cost - total) * rate;
            if year as f64 > life {
                depreciation *= (12.0 - month) / 12.0;
            }
            total += depreciation;
        }
        Ok(depreciation)
    })
}

/// DDB(cost, salvage, life, period, [factor]): declining-balance depreciation
/// at `factor / life` per period, never taking the book value below salvage.
pub(crate) fn ddb(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 4, 5)?;
        let cost = nth_number(args, ctx, 0)?;
        let salvage = nth_number(args, ctx, 1)?;
        let life = nth_number(args, ctx, 2)?;
        let period = nth_number(args, ctx, 3)?;
        let factor = optional(args, ctx, 4, 2.0)?;
        if cost < 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0.0 || factor <= 0.0 {
            return Err(ErrorValue::Num);
        }
        if period > life {
            return Err(ErrorValue::Num);
        }
        Ok(declining(cost, salvage, life, period, factor))
    })
}

/// one period of declining-balance depreciation from the closed-form book
/// values before and after it.
fn declining(cost: f64, salvage: f64, life: f64, period: f64, factor: f64) -> f64 {
    let rate = (factor / life).min(1.0);
    let before = cost * fmath::pow(1.0 - rate, period - 1.0);
    let after = cost * fmath::pow(1.0 - rate, period);
    let depreciation = if after < salvage {
        before - salvage
    } else {
        before - after
    };
    depreciation.max(0.0)
}

/// VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch]):
/// declining-balance depreciation between two points in time, switching to
/// straight-line once that writes off more, unless `no_switch`. a fractional
/// period takes its share of that whole period's depreciation.
pub(crate) fn vdb(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 5, 7)?;
        let [cost, salvage, life, start, end] = numbers(args, ctx)?;
        let factor = optional(args, ctx, 5, 2.0)?;
        let no_switch = optional(args, ctx, 6, 0.0)? != 0.0;
        if cost < 0.0 || salvage < 0.0 || life <= 0.0 || factor <= 0.0 {
            return Err(ErrorValue::Num);
        }
        if start < 0.0 || end < start || end > life {
            return Err(ErrorValue::Num);
        }
        let periods = end.ceil();
        charge(ctx, periods)?;
        let mut schedule = Vec::with_capacity(periods as usize);
        let mut written_off = 0.0;
        let mut straight_line = None;
        for index in 0..periods as usize {
            let period = index as f64 + 1.0;
            let depreciation = match straight_line {
                Some(amount) => amount,
                None => {
                    let declined = declining(cost, salvage, life, period, factor);
                    let remaining = (cost - salvage - written_off) / (life - index as f64);
                    if !no_switch && remaining > declined {
                        straight_line = Some(remaining);
                        remaining
                    } else {
                        declined
                    }
                }
            };
            written_off += depreciation;
            schedule.push(depreciation);
        }
        let through = |time: f64| {
            let whole = time.floor() as usize;
            let mut total: f64 = schedule[..whole].iter().sum();
            if let Some(next) = schedule.get(whole) {
                total += (time - whole as f64) * next;
            }
            total
        };
        Ok(through(end) - through(start))
    })
}

/// EFFECT(nominal_rate, npery): the effective annual rate of a nominal rate
/// compounded `npery` times a year (truncated).
pub(crate) fn effect(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 2, 2)?;
        let nominal = nth_number(args, ctx, 0)?;
        let periods = nth_number(args, ctx, 1)?.trunc();
        if nominal <= 0.0 || periods < 1.0 {
            return Err(ErrorValue::Num);
        }
        Ok(fmath::pow(1.0 + nominal / periods, periods) - 1.0)
    })
}

/// NOMINAL(effect_rate, npery): the nominal rate that compounds to
/// `effect_rate` over `npery` periods a year (truncated).
pub(crate) fn nominal(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 2, 2)?;
        let effective = nth_number(args, ctx, 0)?;
        let periods = nth_number(args, ctx, 1)?.trunc();
        if effective <= 0.0 || periods < 1.0 {
            return Err(ErrorValue::Num);
        }
        Ok(periods * (fmath::pow(1.0 + effective, 1.0 / periods) - 1.0))
    })
}

/// PDURATION(rate, pv, fv): how many periods `pv` takes to grow to `fv`.
pub(crate) fn pduration(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 3, 3)?;
        let [rate, pv, fv] = numbers(args, ctx)?;
        if rate <= 0.0 || pv <= 0.0 || fv <= 0.0 {
            return Err(ErrorValue::Num);
        }
        Ok((fmath::ln(fv) - fmath::ln(pv)) / fmath::ln(1.0 + rate))
    })
}

/// RRI(nper, pv, fv): the per-period growth rate that takes `pv` to `fv`. a
/// non-positive `nper` or `pv`, or a negative `fv`, is `#NUM!`.
pub(crate) fn rri(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 3, 3)?;
        let [nper, pv, fv] = numbers(args, ctx)?;
        if nper <= 0.0 || pv <= 0.0 || fv < 0.0 {
            return Err(ErrorValue::Num);
        }
        Ok(fmath::pow(fv / pv, 1.0 / nper) - 1.0)
    })
}

/// FVSCHEDULE(principal, schedule): `principal` grown by each rate in turn.
/// blank cells in the schedule count as zero, other non-numbers are
/// `#VALUE!`.
pub(crate) fn fvschedule(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 2, 2)?;
        let principal = nth_number(args, ctx, 0)?;
        let (_, rates) = positional_numbers(&args[1], ctx)?;
        Ok(rates
            .iter()
            .fold(principal, |value, rate| value * (1.0 + rate)))
    })
}

/// DOLLARDE(fractional_dollar, fraction): a price written as whole dollars
/// and a numerator over `fraction` (truncated), as a decimal.
pub(crate) fn dollarde(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer(dollar(args, ctx, |part, fraction, scale| {
        part * scale / fraction
    }))
}

/// DOLLARFR(decimal_dollar, fraction): a decimal price written as whole
/// dollars and a numerator over `fraction` (truncated).
pub(crate) fn dollarfr(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer(dollar(args, ctx, |part, fraction, scale| {
        part * fraction / scale
    }))
}

/// the shared shape of DOLLARDE and DOLLARFR: the fractional part is
/// rescaled between `fraction` and the power of ten with as many digits. a
/// negative fraction is `#NUM!`, one that truncates to zero `#DIV/0!`.
fn dollar(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    rescale: impl Fn(f64, f64, f64) -> f64,
) -> Result<f64, ErrorValue> {
    arity(args, 2, 2)?;
    let price = nth_number(args, ctx, 0)?;
    let fraction = nth_number(args, ctx, 1)?;
    if fraction < 0.0 {
        return Err(ErrorValue::Num);
    }
    let fraction = fraction.trunc();
    if fraction == 0.0 {
        return Err(ErrorValue::Div0);
    }
    let mut scale = 1.0;
    while scale < fraction {
        scale *= 10.0;
    }
    let magnitude = price.abs();
    let whole = magnitude.floor();
    let value = whole + rescale(magnitude - whole, fraction, scale);
    Ok(if price < 0.0 { -value } else { value })
}

/// run a function body and show what it computed.
fn compute(body: impl FnOnce() -> Result<f64, ErrorValue>) -> CellValue {
    answer(body())
}

/// a computed result, or the error that stopped it; non-finite is `#NUM!`.
fn answer(result: Result<f64, ErrorValue>) -> CellValue {
    match result {
        Ok(value) => finite(value),
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

/// the leading `N` arguments as numbers.
fn numbers<const N: usize>(args: &[Expr], ctx: &EvalContext<'_>) -> Result<[f64; N], ErrorValue> {
    let mut out = [0.0; N];
    for (i, slot) in out.iter_mut().enumerate() {
        *slot = nth_number(args, ctx, i)?;
    }
    Ok(out)
}

/// a trailing argument, `default` when absent or written as a gap; a blank
/// cell reads as zero.
fn optional(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
    default: f64,
) -> Result<f64, ErrorValue> {
    match args.get(index) {
        Some(arg) if !omitted(arg) => super::coerce_number(&evaluate(arg, ctx), ctx),
        _ => Ok(default),
    }
}

/// spend the recalculation budget on a loop of `steps` iterations.
fn charge(ctx: &EvalContext<'_>, steps: f64) -> Result<(), ErrorValue> {
    if steps.is_finite() && ctx.consume_cells(steps.max(0.0) as u64) {
        Ok(())
    } else {
        Err(ErrorValue::Num)
    }
}
