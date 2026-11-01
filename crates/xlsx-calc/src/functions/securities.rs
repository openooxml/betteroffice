//! securities and day-count financial functions: coupon schedules, bond
//! prices, yields and durations, discount and interest-at-maturity
//! securities, treasury bills and the french depreciation schedules.
//!
//! dates are read as 1900 serials whatever the workbook's date system, with
//! their time of day dropped. coupon dates step back from maturity (or from
//! the first or last interest date of an odd period) by 12 / frequency
//! months, and land on the month's last day when that anchor date does.
//! day-count bases: 0 us 30/360, 1 actual/actual, 2 actual/360, 3
//! actual/365, 4 european 30/360.

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{EvalContext, err, evaluate, num};
use crate::fmath;
use crate::parser::Expr;

use super::datetime::{
    civil_to_serial, days_in_month, from_1900, in_date_range, serial_to_civil, to_1900,
    year_fraction,
};
use super::{coerce_number, finite, nth_number, omitted};

/// how many steps the yield searches take before they give up with `#NUM!`.
const YIELD_STEPS: usize = 100;
/// a newton step this small, relative to the yield, ends a yield search.
const YIELD_TOLERANCE: f64 = 1e-13;
/// the coupon period treasury bills fall back to the bond-equivalent
/// quadratic beyond.
const HALF_YEAR_DAYS: f64 = 182.0;

/// COUPDAYBS(settlement, maturity, frequency, [basis]): days from the start
/// of the coupon period to settlement.
pub(crate) fn coupdaybs(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (settlement, maturity, freq, basis) = coupon_args(args, ctx)?;
        Ok(Coupon::new(settlement, maturity, freq, basis)?.accrued)
    })
}

/// COUPDAYS(settlement, maturity, frequency, [basis]): days in the coupon
/// period that holds settlement; only actual/actual counts them, the other
/// bases divide a 360 or 365 day year by the frequency.
pub(crate) fn coupdays(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (settlement, maturity, freq, basis) = coupon_args(args, ctx)?;
        Ok(Coupon::new(settlement, maturity, freq, basis)?.length)
    })
}

/// COUPDAYSNC(settlement, maturity, frequency, [basis]): days from settlement
/// to the next coupon. the 30/360 bases take the period's nominal length less
/// COUPDAYBS, the others count actual days.
pub(crate) fn coupdaysnc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (settlement, maturity, freq, basis) = coupon_args(args, ctx)?;
        Ok(Coupon::new(settlement, maturity, freq, basis)?.to_next)
    })
}

/// COUPNUM(settlement, maturity, frequency, [basis]): coupons payable
/// between settlement and maturity.
pub(crate) fn coupnum(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (settlement, maturity, freq, _) = coupon_args(args, ctx)?;
        Ok(Period::around(settlement, maturity, freq)?.remaining as f64)
    })
}

/// COUPNCD(settlement, maturity, frequency, [basis]): the next coupon date
/// after settlement.
pub(crate) fn coupncd(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    date_result(ctx, || {
        let (settlement, maturity, freq, _) = coupon_args(args, ctx)?;
        Ok(Period::around(settlement, maturity, freq)?.next)
    })
}

/// COUPPCD(settlement, maturity, frequency, [basis]): the coupon date on or
/// before settlement.
pub(crate) fn couppcd(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    date_result(ctx, || {
        let (settlement, maturity, freq, _) = coupon_args(args, ctx)?;
        Ok(Period::around(settlement, maturity, freq)?.previous)
    })
}

/// ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis],
/// [calc_method]): interest accrued on a periodic-coupon security, summed
/// over the quasi-coupon periods of the first interest date's schedule. a
/// whole period counts one coupon, a part counts its days over the period's
/// length. a false `calc_method` starts at first_interest once settlement is
/// past it; otherwise, and by default, accrual starts at issue. an omitted
/// par is 1000.
pub(crate) fn accrint(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 6, 8)?;
        let [issue, first, settlement, rate] = numbers::<4>(args, ctx)?;
        let par = optional(args, ctx, 4, 1000.0)?;
        let freq = nth_number(args, ctx, 5)?;
        let basis = optional(args, ctx, 6, 0.0)?;
        let from_issue = optional(args, ctx, 7, 1.0)? != 0.0;
        let (issue, first, settlement) =
            (day(issue, ctx)?, day(first, ctx)?, day(settlement, ctx)?);
        let (freq, basis) = (frequency(freq)?, day_basis(basis)?);
        if issue >= settlement || rate <= 0.0 || par <= 0.0 {
            return Err(ErrorValue::Num);
        }
        let start = if from_issue || settlement <= first {
            issue
        } else {
            first
        };
        let quasi = Quasi::new(first, freq, is_month_end(first)?);
        let mut k = quasi.index_of(start)?;
        let mut coupons = 0.0;
        loop {
            let (lo, hi) = (quasi.at(k)?, quasi.at(k - 1)?);
            if lo >= settlement {
                break;
            }
            charge(ctx, 1)?;
            coupons += quasi.share(lo.max(start), hi.min(settlement), lo, hi, basis)?;
            k -= 1;
        }
        Ok(par * rate / freq as f64 * coupons)
    })
}

/// ACCRINTM(issue, settlement, rate, [par], [basis]): interest accrued on a
/// security that pays at maturity, as YEARFRAC over the basis. an omitted
/// par is 1000.
pub(crate) fn accrintm(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        arity(args, 3, 5)?;
        let [issue, settlement, rate] = numbers::<3>(args, ctx)?;
        let par = optional(args, ctx, 3, 1000.0)?;
        let basis = optional(args, ctx, 4, 0.0)?;
        let (issue, settlement) = (day(issue, ctx)?, day(settlement, ctx)?);
        let basis = day_basis(basis)?;
        if issue >= settlement || rate <= 0.0 || par <= 0.0 {
            return Err(ErrorValue::Num);
        }
        Ok(par * rate * fraction(issue, settlement, basis)?)
    })
}

/// PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis]):
/// the clean price per 100 face of a periodic-coupon bond, excel's
/// documented sum over the coupons left, even when one is.
pub(crate) fn price(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (bond, rate, yld, redemption) = bond_args(args, ctx)?;
        if rate < 0.0 || yld < 0.0 || redemption <= 0.0 {
            return Err(ErrorValue::Num);
        }
        Flows::regular(&bond.coupon).price(rate, yld, redemption, bond.freq, ctx)
    })
}

/// YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis]):
/// the yield of a periodic-coupon bond. with one coupon left excel solves
/// its closed form; otherwise it runs newton's method on PRICE, here kept
/// inside a bisection bracket and limited to excel's 100 steps.
pub(crate) fn yield_(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (bond, rate, pr, redemption) = bond_args(args, ctx)?;
        if rate < 0.0 || pr <= 0.0 || redemption <= 0.0 {
            return Err(ErrorValue::Num);
        }
        let c = &bond.coupon;
        let f = bond.freq as f64;
        if c.remaining <= 1.0 {
            let paid = pr / 100.0 + c.accrued / c.length * rate / f;
            let returned = redemption / 100.0 + rate / f;
            return Ok((returned - paid) / paid * f * c.length / c.to_next);
        }
        let flows = Flows::regular(c);
        solve_yield(|y| flows.price(rate, y, redemption, bond.freq, ctx), pr, f)
    })
}

/// DURATION(settlement, maturity, coupon, yld, frequency, [basis]): the
/// macaulay duration in years, each payment timed by COUPDAYSNC over
/// COUPDAYS into its period.
pub(crate) fn duration(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (bond, coupon, yld) = duration_args(args, ctx)?;
        macaulay(&bond, coupon, yld, ctx)
    })
}

/// MDURATION(settlement, maturity, coupon, yld, frequency, [basis]): the
/// modified duration, DURATION over one plus the periodic yield.
pub(crate) fn mduration(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (bond, coupon, yld) = duration_args(args, ctx)?;
        Ok(macaulay(&bond, coupon, yld, ctx)? / (1.0 + yld / bond.freq as f64))
    })
}

/// PRICEDISC(settlement, maturity, discount, redemption, [basis]): the price
/// per 100 face of a discounted security.
pub(crate) fn pricedisc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (years, discount, redemption) = discount_args(args, ctx)?;
        Ok(redemption * (1.0 - discount * years))
    })
}

/// DISC(settlement, maturity, pr, redemption, [basis]): the discount rate of
/// a security.
pub(crate) fn disc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (years, pr, redemption) = discount_args(args, ctx)?;
        Ok((redemption - pr) / redemption / years)
    })
}

/// INTRATE(settlement, maturity, investment, redemption, [basis]): the
/// interest rate of a fully invested security.
pub(crate) fn intrate(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (years, investment, redemption) = discount_args(args, ctx)?;
        Ok((redemption - investment) / investment / years)
    })
}

/// RECEIVED(settlement, maturity, investment, discount, [basis]): what a
/// fully invested security pays at maturity. a discount that consumes the
/// whole term, `discount * years >= 1`, is `#NUM!`.
pub(crate) fn received(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (years, investment, discount) = discount_args(args, ctx)?;
        let kept = 1.0 - discount * years;
        if kept <= 0.0 {
            return Err(ErrorValue::Num);
        }
        Ok(investment / kept)
    })
}

/// YIELDDISC(settlement, maturity, pr, redemption, [basis]): the annual
/// yield of a discounted security.
pub(crate) fn yielddisc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (years, pr, redemption) = discount_args(args, ctx)?;
        Ok((redemption / pr - 1.0) / years)
    })
}

/// PRICEMAT(settlement, maturity, issue, rate, yld, [basis]): the price per
/// 100 face of a security that pays its interest at maturity.
pub(crate) fn pricemat(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (span, rate, yld) = maturity_args(args, ctx)?;
        if rate < 0.0 || yld < 0.0 {
            return Err(ErrorValue::Num);
        }
        Ok((100.0 + span.term * rate * 100.0) / (1.0 + span.left * yld) - span.held * rate * 100.0)
    })
}

/// YIELDMAT(settlement, maturity, issue, rate, pr, [basis]): the annual
/// yield of a security that pays its interest at maturity.
pub(crate) fn yieldmat(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (span, rate, pr) = maturity_args(args, ctx)?;
        if rate < 0.0 || pr <= 0.0 {
            return Err(ErrorValue::Num);
        }
        let paid = pr / 100.0 + span.held * rate;
        Ok((1.0 + span.term * rate - paid) / paid / span.left)
    })
}

/// TBILLPRICE(settlement, maturity, discount): the price per 100 face of a
/// treasury bill on an actual/360 discount; a price at or below zero is
/// `#NUM!`.
pub(crate) fn tbillprice(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (days, discount) = bill_args(args, ctx)?;
        let price = 1.0 - discount * days / 360.0;
        if discount <= 0.0 || price <= 0.0 {
            return Err(ErrorValue::Num);
        }
        Ok(100.0 * price)
    })
}

/// TBILLYIELD(settlement, maturity, pr): the yield of a treasury bill.
pub(crate) fn tbillyield(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (days, pr) = bill_args(args, ctx)?;
        if pr <= 0.0 {
            return Err(ErrorValue::Num);
        }
        Ok((100.0 - pr) / pr * 360.0 / days)
    })
}

/// TBILLEQ(settlement, maturity, discount): the bond-equivalent yield of a
/// treasury bill. up to 182 days it is excel's documented
/// `365 * rate / (360 - rate * days)`; a longer bill solves the quadratic for
/// a semiannual coupon equivalent, as excel does. a discount that prices
/// the bill at or below zero is `#NUM!`.
pub(crate) fn tbilleq(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (days, discount) = bill_args(args, ctx)?;
        let price = 1.0 - discount * days / 360.0;
        if discount <= 0.0 || price <= 0.0 {
            return Err(ErrorValue::Num);
        }
        if days <= HALF_YEAR_DAYS {
            return Ok(365.0 * discount / (360.0 - discount * days));
        }
        let t = days / 365.0;
        let root = (t * t - (2.0 * t - 1.0) * (1.0 - 1.0 / price)).sqrt();
        Ok((root - t) / (t - 0.5))
    })
}

/// ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld,
/// redemption, frequency, [basis]): the price per 100 face of a bond whose
/// first period is short or long, by excel's documented formulas. a
/// redemption at or below zero is `#NUM!`, as it is for PRICE.
pub(crate) fn oddfprice(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (odd, freq, rate, yld, redemption) = odd_first_args(args, ctx)?;
        if rate < 0.0 || yld < 0.0 || redemption <= 0.0 {
            return Err(ErrorValue::Num);
        }
        odd.price(rate, yld, redemption, freq, ctx)
    })
}

/// ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr,
/// redemption, frequency, [basis]): the yield of a bond with an odd first
/// period, by the same bracketed newton search as YIELD on ODDFPRICE.
pub(crate) fn oddfyield(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (odd, freq, rate, pr, redemption) = odd_first_args(args, ctx)?;
        if rate < 0.0 || pr <= 0.0 || redemption <= 0.0 {
            return Err(ErrorValue::Num);
        }
        solve_yield(
            |y| odd.price(rate, y, redemption, freq, ctx),
            pr,
            freq as f64,
        )
    })
}

/// ODDLPRICE(settlement, maturity, last_interest, rate, yld, redemption,
/// frequency, [basis]): the price per 100 face of a bond whose last period
/// is odd, summed over the quasi-coupon periods after last_interest. a
/// redemption at or below zero is `#NUM!`, as it is for PRICE.
pub(crate) fn oddlprice(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (odd, rate, yld, redemption) = odd_last_args(args, ctx)?;
        if rate < 0.0 || yld < 0.0 || redemption <= 0.0 {
            return Err(ErrorValue::Num);
        }
        let f = odd.freq as f64;
        let cash = 100.0 * rate / f;
        Ok((redemption + odd.counted * cash) / (1.0 + odd.left * yld / f) - odd.accrued * cash)
    })
}

/// ODDLYIELD(settlement, maturity, last_interest, rate, pr, redemption,
/// frequency, [basis]): the yield of a bond whose last period is odd, excel's
/// closed form.
pub(crate) fn oddlyield(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let (odd, rate, pr, redemption) = odd_last_args(args, ctx)?;
        if rate < 0.0 || pr <= 0.0 || redemption <= 0.0 {
            return Err(ErrorValue::Num);
        }
        let f = odd.freq as f64;
        let cash = 100.0 * rate / f;
        let paid = pr + odd.accrued * cash;
        Ok((redemption + odd.counted * cash - paid) / paid * f / odd.left)
    })
}

/// AMORLINC(cost, date_purchased, first_period, salvage, period, rate,
/// [basis]): french straight-line depreciation. period 0 is the prorated
/// first period, then full `cost * rate` periods until the last takes what
/// is left above salvage. a first period worth more than cost less salvage
/// takes only that, and nothing is left for later periods. basis 2 is not
/// offered, as in excel.
pub(crate) fn amorlinc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let asset = asset_args(args, ctx)?;
        let full_charge = asset.cost * asset.rate;
        let depreciable = asset.cost - asset.salvage;
        let first = (asset.first_share * asset.rate * asset.cost).min(depreciable);
        let full_periods = ((depreciable - first) / full_charge).trunc();
        let period = asset.period as f64;
        Ok(if period == 0.0 {
            first
        } else if period <= full_periods {
            full_charge
        } else if period == full_periods + 1.0 {
            depreciable - full_charge * full_periods - first
        } else {
            0.0
        })
    })
}

/// AMORDEGRC(cost, date_purchased, first_period, salvage, period, rate,
/// [basis]): french declining-balance depreciation. the rate is scaled by
/// the coefficient for the asset's life `1 / rate` (1.5 for 3 to 4 years,
/// 2 for 5 to 6, 2.5 beyond) and every charge rounds to a whole unit; the
/// period that would cut into salvage, and the one before it, take half the
/// remaining value. a first period worth more than cost less salvage takes
/// only that and leaves nothing for later periods. each later period spends
/// one unit of the recalculation budget. a life between 0 and 1, 1 and 2,
/// 2 and 3 or 4 and 5 years is `#NUM!`, as excel documents.
pub(crate) fn amordegrc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    compute(|| {
        let asset = asset_args(args, ctx)?;
        let life = 1.0 / asset.rate;
        let gap = |lo: f64, hi: f64| life > lo && life < hi;
        if gap(0.0, 1.0) || gap(1.0, 2.0) || gap(2.0, 3.0) || gap(4.0, 5.0) {
            return Err(ErrorValue::Num);
        }
        let coefficient = if life < 3.0 {
            1.0
        } else if life <= 4.0 {
            1.5
        } else if life <= 6.0 {
            2.0
        } else {
            2.5
        };
        let rate = asset.rate * coefficient;
        let depreciable = asset.cost - asset.salvage;
        let mut amount = (asset.first_share * rate * asset.cost).round();
        if amount >= depreciable {
            return Ok(if asset.period == 0 { depreciable } else { 0.0 });
        }
        let mut value = asset.cost - amount;
        let mut rest = value - asset.salvage;
        for n in 0..asset.period {
            charge(ctx, 1)?;
            amount = (rate * value).round();
            if amount == 0.0 {
                return Ok(0.0);
            }
            rest -= amount;
            if rest < 0.0 {
                return Ok(if asset.period - n <= 1 {
                    (value * 0.5).round()
                } else {
                    0.0
                });
            }
            value -= amount;
        }
        Ok(amount)
    })
}

/// the coupon period holding a settlement date, on the schedule that steps
/// back from maturity, and how many coupons are left after settlement.
struct Period {
    previous: i64,
    next: i64,
    remaining: i64,
}

impl Period {
    fn around(settlement: i64, maturity: i64, freq: i64) -> Result<Self, ErrorValue> {
        let quasi = Quasi::new(maturity, freq, is_month_end(maturity)?);
        let k = quasi.index_of(settlement)?;
        Ok(Period {
            previous: quasi.at(k)?,
            next: quasi.at(k - 1)?,
            remaining: k,
        })
    }
}

/// the figures excel's bond formulas are written in: days from the previous
/// coupon to settlement (A), days in the period (E), days from settlement to
/// the next coupon (DSC) and the coupons left (N).
struct Coupon {
    accrued: f64,
    length: f64,
    to_next: f64,
    remaining: f64,
}

impl Coupon {
    fn new(settlement: i64, maturity: i64, freq: i64, basis: i64) -> Result<Self, ErrorValue> {
        let period = Period::around(settlement, maturity, freq)?;
        let accrued = days_between(period.previous, settlement, basis)?;
        let length = period_days(period.previous, period.next, freq, basis);
        let to_next = match basis {
            0 | 4 => length - accrued,
            _ => (period.next - settlement) as f64,
        };
        Ok(Coupon {
            accrued,
            length,
            to_next,
            remaining: period.remaining as f64,
        })
    }
}

/// the quasi-coupon dates of a schedule anchored at one date: `at(k)` is `k`
/// periods before the anchor (after it for negative `k`), on the month's last
/// day when `month_end`.
struct Quasi {
    anchor: i64,
    step: i64,
    month_end: bool,
    freq: i64,
}

impl Quasi {
    fn new(anchor: i64, freq: i64, month_end: bool) -> Self {
        Quasi {
            anchor,
            step: 12 / freq,
            month_end,
            freq,
        }
    }

    fn at(&self, k: i64) -> Result<i64, ErrorValue> {
        add_months(self.anchor, -self.step * k, self.month_end)
    }

    /// the `k` whose period `(at(k), at(k - 1)]` starts on or before `date`
    /// and ends after it.
    fn index_of(&self, date: i64) -> Result<i64, ErrorValue> {
        let (ay, am, _) = ymd(self.anchor)?;
        let (dy, dm, _) = ymd(date)?;
        let months = (ay - dy) * 12 + (am - dm);
        let mut k = months.div_euclid(self.step);
        while self.at(k)? > date {
            k += 1;
        }
        while self.at(k - 1)? <= date {
            k -= 1;
        }
        Ok(k)
    }

    /// the coupons `[from, to)` is worth inside the period `[lo, hi)`: one
    /// for the whole period, its days over the period's length for a part.
    fn share(&self, from: i64, to: i64, lo: i64, hi: i64, basis: i64) -> Result<f64, ErrorValue> {
        if from == lo && to == hi {
            Ok(1.0)
        } else if to > from {
            Ok(days_between(from, to, basis)? / period_days(lo, hi, self.freq, basis))
        } else {
            Ok(0.0)
        }
    }
}

/// a bond's payments in coupon units: the first coupon `lead` periods out
/// and worth `first` coupons, then `regular` whole coupons a period apart
/// with the redemption paid alongside the last, less the `accrued` coupons
/// the buyer owes the seller.
struct Flows {
    lead: f64,
    first: f64,
    regular: f64,
    accrued: f64,
}

impl Flows {
    fn regular(c: &Coupon) -> Self {
        Flows {
            lead: c.to_next / c.length,
            first: 1.0,
            regular: c.remaining - 1.0,
            accrued: c.accrued / c.length,
        }
    }

    /// the clean price at `yld`, each payment discounted charged to the
    /// recalculation budget. a zero coupon adds nothing, not the `0 / 0` of
    /// a discount factor that underflows.
    fn price(
        &self,
        rate: f64,
        yld: f64,
        redemption: f64,
        freq: i64,
        ctx: &EvalContext<'_>,
    ) -> Result<f64, ErrorValue> {
        let regular = self.regular as u64;
        charge(ctx, regular.saturating_add(1))?;
        let f = freq as f64;
        let cash = 100.0 * rate / f;
        let base = 1.0 + yld / f;
        let mut total = redemption / fmath::pow(base, self.lead + self.regular);
        if cash != 0.0 {
            total += cash * self.first / fmath::pow(base, self.lead);
            for j in 1..=regular {
                total += cash / fmath::pow(base, self.lead + j as f64);
            }
            total -= cash * self.accrued;
        }
        Ok(total)
    }
}

/// a settled periodic-coupon bond: its coupon figures and frequency.
struct Bond {
    coupon: Coupon,
    freq: i64,
}

/// PRICE and YIELD's arguments: settlement, maturity, rate, the yield or
/// price, redemption, frequency, [basis].
fn bond_args(args: &[Expr], ctx: &EvalContext<'_>) -> Result<(Bond, f64, f64, f64), ErrorValue> {
    arity(args, 6, 7)?;
    let [settlement, maturity, rate, quote, redemption, freq] = numbers::<6>(args, ctx)?;
    let basis = optional(args, ctx, 6, 0.0)?;
    let bond = settle(settlement, maturity, freq, basis, ctx)?;
    Ok((bond, rate, quote, redemption))
}

/// DURATION's arguments: settlement, maturity, coupon, yld, frequency,
/// [basis]; a negative coupon or yield is `#NUM!`.
fn duration_args(args: &[Expr], ctx: &EvalContext<'_>) -> Result<(Bond, f64, f64), ErrorValue> {
    arity(args, 5, 6)?;
    let [settlement, maturity, coupon, yld, freq] = numbers::<5>(args, ctx)?;
    let basis = optional(args, ctx, 5, 0.0)?;
    let bond = settle(settlement, maturity, freq, basis, ctx)?;
    if coupon < 0.0 || yld < 0.0 {
        return Err(ErrorValue::Num);
    }
    Ok((bond, coupon, yld))
}

/// validate a bond's dates, frequency and basis and find its coupon period.
fn settle(
    settlement: f64,
    maturity: f64,
    freq: f64,
    basis: f64,
    ctx: &EvalContext<'_>,
) -> Result<Bond, ErrorValue> {
    let (settlement, maturity) = (day(settlement, ctx)?, day(maturity, ctx)?);
    let (freq, basis) = (frequency(freq)?, day_basis(basis)?);
    if settlement >= maturity {
        return Err(ErrorValue::Num);
    }
    Ok(Bond {
        coupon: Coupon::new(settlement, maturity, freq, basis)?,
        freq,
    })
}

/// the macaulay duration in years, each payment charged to the
/// recalculation budget.
fn macaulay(bond: &Bond, coupon: f64, yld: f64, ctx: &EvalContext<'_>) -> Result<f64, ErrorValue> {
    let c = &bond.coupon;
    let payments = c.remaining as u64;
    charge(ctx, payments)?;
    let f = bond.freq as f64;
    let cash = 100.0 * coupon / f;
    let base = 1.0 + yld / f;
    let lead = c.to_next / c.length - 1.0;
    let (mut weighted, mut total) = (0.0, 0.0);
    for k in 1..=payments {
        let t = k as f64 + lead;
        let flow = if k == payments { cash + 100.0 } else { cash };
        let present = flow / fmath::pow(base, t);
        weighted += t * present;
        total += present;
    }
    Ok(weighted / total / f)
}

/// the yield at which `price` comes to `target`. prices fall as yields rise,
/// so newton steps run inside a bracket, and bisection takes over whenever a
/// newton step would leave the bracket or fails to halve the step before it,
/// so a slow crawl cannot stall the search. it ends on a step too small to
/// matter; no convergence within excel's 100 steps is `#NUM!`, never the last
/// iterate.
fn solve_yield(
    price: impl Fn(f64) -> Result<f64, ErrorValue>,
    target: f64,
    freq: f64,
) -> Result<f64, ErrorValue> {
    let gap = |y: f64| price(y).map(|p| p - target);
    let mut lo = -freq;
    let mut hi = 1.0;
    let mut widenings = 0;
    while gap(hi)? > 0.0 {
        lo = hi;
        hi *= 2.0;
        widenings += 1;
        if widenings > 60 {
            return Err(ErrorValue::Num);
        }
    }
    let mut y = if lo < 0.1 && 0.1 < hi {
        0.1
    } else {
        0.5 * (lo + hi)
    };
    let mut step = hi - lo;
    let mut step_before = step;
    for _ in 0..YIELD_STEPS {
        let g = gap(y)?;
        if g.is_nan() {
            return Err(ErrorValue::Num);
        }
        if g == 0.0 {
            return Ok(y);
        }
        if g > 0.0 {
            lo = y;
        } else {
            hi = y;
        }
        let h = 1e-7 * (1.0 + y.abs());
        let slope = (gap(y + h)? - g) / h;
        let newton = y - g / slope;
        let fast = newton > lo && newton < hi && (2.0 * g).abs() <= (step_before * slope).abs();
        let next = if fast { newton } else { 0.5 * (lo + hi) };
        step_before = step;
        step = next - y;
        if step.abs() <= YIELD_TOLERANCE * (1.0 + next.abs()) {
            return Ok(next);
        }
        y = next;
    }
    Err(ErrorValue::Num)
}

/// PRICEDISC, DISC, INTRATE, RECEIVED and YIELDDISC's arguments:
/// settlement, maturity, two positive amounts, [basis]. the term comes back
/// as YEARFRAC.
fn discount_args(args: &[Expr], ctx: &EvalContext<'_>) -> Result<(f64, f64, f64), ErrorValue> {
    arity(args, 4, 5)?;
    let [settlement, maturity, x, y] = numbers::<4>(args, ctx)?;
    let basis = optional(args, ctx, 4, 0.0)?;
    let (settlement, maturity) = (day(settlement, ctx)?, day(maturity, ctx)?);
    let basis = day_basis(basis)?;
    if settlement >= maturity || x <= 0.0 || y <= 0.0 {
        return Err(ErrorValue::Num);
    }
    Ok((fraction(settlement, maturity, basis)?, x, y))
}

/// an interest-at-maturity security's terms in years: issue to maturity,
/// issue to settlement and settlement to maturity.
struct Span {
    term: f64,
    held: f64,
    left: f64,
}

/// PRICEMAT and YIELDMAT's arguments: settlement, maturity, issue, rate, the
/// yield or price, [basis]. settlement must fall after issue.
fn maturity_args(args: &[Expr], ctx: &EvalContext<'_>) -> Result<(Span, f64, f64), ErrorValue> {
    arity(args, 5, 6)?;
    let [settlement, maturity, issue, rate, quote] = numbers::<5>(args, ctx)?;
    let basis = optional(args, ctx, 5, 0.0)?;
    let (settlement, maturity, issue) =
        (day(settlement, ctx)?, day(maturity, ctx)?, day(issue, ctx)?);
    let basis = day_basis(basis)?;
    if settlement >= maturity || issue >= settlement {
        return Err(ErrorValue::Num);
    }
    let span = Span {
        term: fraction(issue, maturity, basis)?,
        held: fraction(issue, settlement, basis)?,
        left: fraction(settlement, maturity, basis)?,
    };
    Ok((span, rate, quote))
}

/// the treasury-bill functions' arguments: settlement, maturity and one
/// amount. the bill may run at most one calendar year; the days come back
/// actual.
fn bill_args(args: &[Expr], ctx: &EvalContext<'_>) -> Result<(f64, f64), ErrorValue> {
    arity(args, 3, 3)?;
    let [settlement, maturity, amount] = numbers::<3>(args, ctx)?;
    let (settlement, maturity) = (day(settlement, ctx)?, day(maturity, ctx)?);
    if settlement >= maturity || maturity > add_months(settlement, 12, false)? {
        return Err(ErrorValue::Num);
    }
    Ok(((maturity - settlement) as f64, amount))
}

/// a bond with an odd last period, in the coupon units of excel's formula:
/// the coupons the odd period pays, those accrued at settlement and the
/// periods left from settlement to maturity.
struct OddLast {
    counted: f64,
    accrued: f64,
    left: f64,
    freq: i64,
}

/// ODDLPRICE and ODDLYIELD's arguments: settlement, maturity, last_interest,
/// rate, the yield or price, redemption, frequency, [basis]. they need
/// maturity > settlement > last_interest.
fn odd_last_args(
    args: &[Expr],
    ctx: &EvalContext<'_>,
) -> Result<(OddLast, f64, f64, f64), ErrorValue> {
    arity(args, 7, 8)?;
    let [settlement, maturity, last, rate, quote, redemption, freq] = numbers::<7>(args, ctx)?;
    let basis = optional(args, ctx, 7, 0.0)?;
    let (settlement, maturity, last) =
        (day(settlement, ctx)?, day(maturity, ctx)?, day(last, ctx)?);
    let (freq, basis) = (frequency(freq)?, day_basis(basis)?);
    if !(maturity > settlement && settlement > last) {
        return Err(ErrorValue::Num);
    }
    let quasi = Quasi::new(last, freq, is_month_end(last)?);
    let mut odd = OddLast {
        counted: 0.0,
        accrued: 0.0,
        left: 0.0,
        freq,
    };
    let mut k = 0;
    loop {
        let lo = quasi.at(k)?;
        if lo >= maturity {
            break;
        }
        charge(ctx, 1)?;
        let hi = quasi.at(k - 1)?;
        odd.counted += quasi.share(lo, hi.min(maturity), lo, hi, basis)?;
        odd.accrued += quasi.share(lo, hi.min(settlement), lo, hi, basis)?;
        odd.left += quasi.share(lo.max(settlement), hi.min(maturity), lo, hi, basis)?;
        k -= 1;
    }
    Ok((odd, rate, quote, redemption))
}

/// ODDFPRICE and ODDFYIELD's arguments: settlement, maturity, issue,
/// first_coupon, rate, the yield or price, redemption, frequency, [basis].
/// they need maturity > first_coupon > settlement > issue. the payments come
/// back in coupon units; a first period starting after its quasi-coupon
/// date is short, one starting before it long.
fn odd_first_args(
    args: &[Expr],
    ctx: &EvalContext<'_>,
) -> Result<(Flows, i64, f64, f64, f64), ErrorValue> {
    arity(args, 8, 9)?;
    let [
        settlement,
        maturity,
        issue,
        first,
        rate,
        quote,
        redemption,
        freq,
    ] = numbers::<8>(args, ctx)?;
    let basis = optional(args, ctx, 8, 0.0)?;
    let (settlement, maturity) = (day(settlement, ctx)?, day(maturity, ctx)?);
    let (issue, first) = (day(issue, ctx)?, day(first, ctx)?);
    let (freq, basis) = (frequency(freq)?, day_basis(basis)?);
    if !(maturity > first && first > settlement && settlement > issue) {
        return Err(ErrorValue::Num);
    }
    let after = Period::around(first, maturity, freq)?.remaining as f64;
    let quasi = Quasi::new(first, freq, is_month_end(maturity)?);
    let start = quasi.at(1)?;
    if issue >= start {
        let length = period_days(start, first, freq, basis);
        let flows = Flows {
            lead: days_between(settlement, first, basis)? / length,
            first: days_between(issue, first, basis)? / length,
            regular: after,
            accrued: days_between(issue, settlement, basis)? / length,
        };
        return Ok((flows, freq, rate, quote, redemption));
    }
    let (mut counted, mut accrued) = (0.0, 0.0);
    let mut k = 0;
    loop {
        let hi = quasi.at(k)?;
        if hi <= issue {
            break;
        }
        charge(ctx, 1)?;
        let lo = quasi.at(k + 1)?;
        counted += quasi.share(lo.max(issue), hi, lo, hi, basis)?;
        accrued += quasi.share(lo.max(issue), hi.min(settlement), lo, hi, basis)?;
        k += 1;
    }
    let whole = quasi.index_of(settlement)? - 1;
    let (lo, hi) = (quasi.at(whole + 1)?, quasi.at(whole)?);
    let length = period_days(lo, hi, freq, basis);
    let to_next = match basis {
        0 | 4 => length - days_between(lo, settlement, basis)?,
        _ => days_between(settlement, hi, basis)?,
    };
    let flows = Flows {
        lead: whole as f64 + to_next / length,
        first: counted,
        regular: after,
        accrued,
    };
    Ok((flows, freq, rate, quote, redemption))
}

/// a depreciating asset: cost, the first period's share of a year, salvage,
/// the period asked for and the rate.
struct Asset {
    cost: f64,
    first_share: f64,
    salvage: f64,
    period: u64,
    rate: f64,
}

/// AMORLINC and AMORDEGRC's arguments: cost, date_purchased, first_period,
/// salvage, period, rate, [basis].
fn asset_args(args: &[Expr], ctx: &EvalContext<'_>) -> Result<Asset, ErrorValue> {
    arity(args, 6, 7)?;
    let [cost, purchased, first, salvage, period, rate] = numbers::<6>(args, ctx)?;
    let basis = optional(args, ctx, 6, 0.0)?;
    let (purchased, first) = (day(purchased, ctx)?, day(first, ctx)?);
    let basis = day_basis(basis)?;
    if basis == 2
        || purchased > first
        || rate <= 0.0
        || cost < 0.0
        || salvage < 0.0
        || salvage > cost
        || period < 0.0
    {
        return Err(ErrorValue::Num);
    }
    Ok(Asset {
        cost,
        first_share: fraction(purchased, first, basis)?,
        salvage,
        period: period.trunc() as u64,
        rate,
    })
}

/// the arguments every COUP function takes: settlement, maturity,
/// frequency, [basis].
fn coupon_args(args: &[Expr], ctx: &EvalContext<'_>) -> Result<(i64, i64, i64, i64), ErrorValue> {
    arity(args, 3, 4)?;
    let [settlement, maturity, freq] = numbers::<3>(args, ctx)?;
    let basis = optional(args, ctx, 3, 0.0)?;
    let (settlement, maturity) = (day(settlement, ctx)?, day(maturity, ctx)?);
    let (freq, basis) = (frequency(freq)?, day_basis(basis)?);
    if settlement >= maturity {
        return Err(ErrorValue::Num);
    }
    Ok((settlement, maturity, freq, basis))
}

/// days from `start` to `end` under the basis: 30/360 for 0 and 4, actual
/// otherwise.
fn days_between(start: i64, end: i64, basis: i64) -> Result<f64, ErrorValue> {
    match basis {
        0 => days_360(start, end, false),
        4 => days_360(start, end, true),
        _ => Ok((end - start) as f64),
    }
}

/// days on a 30/360 calendar, for dates past either end of the range too.
/// `european` clamps both days to 30; the us rule folds the end of february
/// and the 31st onto the 30th the way YEARFRAC does.
fn days_360(start: i64, end: i64, european: bool) -> Result<f64, ErrorValue> {
    let (y1, m1, mut d1) = ymd(start)?;
    let (y2, m2, mut d2) = ymd(end)?;
    if european {
        d1 = d1.min(30);
        d2 = d2.min(30);
    } else {
        let february_end = |y: i64, m: i64, d: i64| m == 2 && d == days_in_month(y, 2);
        if february_end(y1, m1, d1) {
            if february_end(y2, m2, d2) {
                d2 = 30;
            }
            d1 = 30;
        }
        if d1 == 31 {
            d1 = 30;
        }
        if d2 == 31 && d1 == 30 {
            d2 = 30;
        }
    }
    Ok(((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) as f64)
}

/// the length of the coupon period `[previous, next)`: actual days for
/// basis 1, a 365 day year for basis 3 and a 360 day one otherwise, divided
/// by the frequency.
fn period_days(previous: i64, next: i64, freq: i64, basis: i64) -> f64 {
    match basis {
        1 => (next - previous) as f64,
        3 => 365.0 / freq as f64,
        _ => 360.0 / freq as f64,
    }
}

/// YEARFRAC from `lo` to `hi`.
fn fraction(lo: i64, hi: i64, basis: i64) -> Result<f64, ErrorValue> {
    year_fraction(lo, hi, basis).ok_or(ErrorValue::Num)
}

/// `serial` moved by whole months: onto the month's last day when
/// `month_end`, otherwise onto the same day clamped to the month's length.
/// the result may lie outside the date range; only dates a function returns
/// or reads as arguments have to be inside it.
fn add_months(serial: i64, months: i64, month_end: bool) -> Result<i64, ErrorValue> {
    let (y, m, d) = ymd(serial)?;
    let total = y * 12 + (m - 1) + months;
    let (year, month) = (total.div_euclid(12), total.rem_euclid(12) + 1);
    let last = days_in_month(year, month);
    let day = if month_end { last } else { d.min(last) };
    civil_to_serial(year, month, day).ok_or(ErrorValue::Num)
}

/// whether `serial` is the last day of its month.
fn is_month_end(serial: i64) -> Result<bool, ErrorValue> {
    let (y, m, d) = ymd(serial)?;
    Ok(d == days_in_month(y, m))
}

fn ymd(serial: i64) -> Result<(i64, i64, i64), ErrorValue> {
    serial_to_civil(serial).ok_or(ErrorValue::Num)
}

/// a date argument as the 1900 serial of its day; a time of day is dropped
/// and a date excel cannot hold is `#NUM!`.
fn day(value: f64, ctx: &EvalContext<'_>) -> Result<i64, ErrorValue> {
    if !(0.0..=f64::from(u32::MAX)).contains(&value) {
        return Err(ErrorValue::Num);
    }
    let serial = to_1900(value.trunc() as i64, ctx.date_system);
    if serial < 1 || !in_date_range(serial) {
        return Err(ErrorValue::Num);
    }
    Ok(serial)
}

/// a coupon frequency: 1, 2 or 4 payments a year after truncation.
fn frequency(value: f64) -> Result<i64, ErrorValue> {
    let freq = value.trunc();
    if freq == 1.0 || freq == 2.0 || freq == 4.0 {
        Ok(freq as i64)
    } else {
        Err(ErrorValue::Num)
    }
}

/// a day-count basis 0..=4 after truncation; a negative one is `#NUM!`
/// before truncation could make it zero.
fn day_basis(value: f64) -> Result<i64, ErrorValue> {
    let basis = value.trunc();
    if value >= 0.0 && basis <= 4.0 {
        Ok(basis as i64)
    } else {
        Err(ErrorValue::Num)
    }
}

fn compute(body: impl FnOnce() -> Result<f64, ErrorValue>) -> CellValue {
    match body() {
        Ok(value) => finite(value),
        Err(error) => err(error),
    }
}

/// a computed 1900 serial as the workbook's date serial; a date outside the
/// range is `#NUM!`.
fn date_result(ctx: &EvalContext<'_>, body: impl FnOnce() -> Result<i64, ErrorValue>) -> CellValue {
    let shown = |serial: i64| {
        (serial >= 1 && in_date_range(serial))
            .then(|| from_1900(serial, ctx.date_system))
            .flatten()
            .ok_or(ErrorValue::Num)
    };
    match body().and_then(shown) {
        Ok(serial) => num(serial as f64),
        Err(error) => err(error),
    }
}

/// spend the recalculation budget on `steps` units of work.
fn charge(ctx: &EvalContext<'_>, steps: u64) -> Result<(), ErrorValue> {
    if ctx.consume_cells(steps) {
        Ok(())
    } else {
        Err(ErrorValue::Num)
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

/// a trailing argument, `default` when it is left out.
fn optional(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
    default: f64,
) -> Result<f64, ErrorValue> {
    match args.get(index) {
        Some(arg) if !omitted(arg) => coerce_number(&evaluate(arg, ctx), ctx),
        _ => Ok(default),
    }
}
