//! discrete distributions and the remaining statistics: the binomial,
//! negative binomial, hypergeometric and poisson families, the exponential
//! and weibull distributions, fisher's transformation, STANDARDIZE, the
//! permutation counts, STEYX, population skewness, the `-A` variances, the
//! exclusive percentiles and quartiles, the percent ranks, MODE.MULT, PROB
//! and Z.TEST. semantics follow microsoft's function reference: counts such
//! as trials and successes are truncated, out-of-domain arguments give
//! `#NUM!`, non-numeric arguments give `#VALUE!`, and the compatibility names
//! share the modern implementations.

use std::collections::HashMap;
use std::f64::consts::TAU;

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{EvalContext, err, evaluate, to_bool};
use crate::parser::Expr;

use super::special;
use super::stats;
use super::{
    argument_cells, collect_numbers, collect_numbers_anytype, collect_numbers_deep, finite,
    nth_number, omitted,
};

fn answer(value: Result<f64, ErrorValue>) -> CellValue {
    match value {
        Ok(value) => finite(value),
        Err(e) => err(e),
    }
}

fn arity(args: &[Expr], min: usize, max: usize) -> Result<(), ErrorValue> {
    if args.len() < min || args.len() > max {
        return Err(ErrorValue::Value);
    }
    Ok(())
}

fn numbers(args: &[Expr], ctx: &EvalContext<'_>, count: usize) -> Result<Vec<f64>, ErrorValue> {
    (0..count).map(|i| nth_number(args, ctx, i)).collect()
}

fn flag(args: &[Expr], ctx: &EvalContext<'_>, i: usize) -> Result<bool, ErrorValue> {
    to_bool(&evaluate(&args[i], ctx))
}

/// ln n! less its stirling approximation, for n = 1 to 15 (mpmath).
const STIRLING_ERRORS: [f64; 15] = [
    0.08106146679532726,
    0.0413406959554093,
    0.02767792568499834,
    0.020790672103765093,
    0.016644691189821193,
    0.013876128823070748,
    0.01189670994589177,
    0.010411265261972096,
    0.009255462182712733,
    0.00833056343336287,
    0.007573675487951841,
    0.00694284010720953,
    0.006408994188004207,
    0.0059513701127588475,
    0.005554733551962801,
];

/// ln n! - ((n + 1/2) ln n - n + ln sqrt(2 pi)) for a whole n >= 1: a table up
/// to 15, the asymptotic series beyond.
fn stirling_error(n: f64) -> f64 {
    if n <= 15.0 {
        return STIRLING_ERRORS[n as usize - 1];
    }
    let inv = 1.0 / n;
    let inv2 = inv * inv;
    inv * (1.0 / 12.0
        - inv2 * (1.0 / 360.0 - inv2 * (1.0 / 1260.0 - inv2 * (1.0 / 1680.0 - inv2 / 1188.0))))
}

/// x ln(x / mean) + mean - x, by its series in (x - mean) / (x + mean) when x
/// is near the mean so the terms do not cancel.
fn deviance(x: f64, mean: f64) -> f64 {
    let half_sum = 0.5 * x + 0.5 * mean;
    if (x - mean).abs() < 0.2 * half_sum {
        let v = 0.5 * (x - mean) / half_sum;
        let v2 = v * v;
        let mut sum = (x - mean) * v;
        let mut power = x * (2.0 * v);
        let mut odd = 1.0;
        for _ in 0..64 {
            power *= v2;
            odd += 2.0;
            let next = sum + power / odd;
            if next == sum || !next.is_finite() {
                break;
            }
            sum = next;
        }
        return sum;
    }
    x * (x / mean).ln() + mean - x
}

/// counts up to this many items have their binomial coefficients computed
/// exactly in integers.
const EXACT_COUNTS: f64 = 120.0;

/// C(n, k) exactly, for n <= [`EXACT_COUNTS`].
fn exact_choose(n: f64, k: f64) -> f64 {
    let (n, k) = (n as u128, k.min(n - k) as u128);
    let mut c: u128 = 1;
    for i in 0..k {
        c = c * (n - i) / (i + 1);
    }
    c as f64
}

/// the probability of `k` successes in `n` trials: exact coefficients for
/// small counts, else the saddle point form that keeps full relative
/// precision for large ones.
fn binomial_pmf(k: f64, n: f64, p: f64) -> f64 {
    let q = 1.0 - p;
    if p == 0.0 {
        return if k == 0.0 { 1.0 } else { 0.0 };
    }
    if q == 0.0 {
        return if k == n { 1.0 } else { 0.0 };
    }
    if n <= EXACT_COUNTS {
        let direct = exact_choose(n, k) * p.powi(k as i32) * q.powi((n - k) as i32);
        if direct > 1e-280 {
            return direct;
        }
    }
    if k == 0.0 {
        return (n * (-p).ln_1p()).exp();
    }
    if k == n {
        return (n * p.ln()).exp();
    }
    let lc = stirling_error(n)
        - stirling_error(k)
        - stirling_error(n - k)
        - deviance(k, n * p)
        - deviance(n - k, n * q);
    let lf = TAU.ln() + k.ln() + (-k / n).ln_1p();
    (lc - 0.5 * lf).exp()
}

/// the poisson probability of `k` events at rate `mean`.
fn poisson_pmf(k: f64, mean: f64) -> f64 {
    if mean == 0.0 {
        return if k == 0.0 { 1.0 } else { 0.0 };
    }
    if k == 0.0 {
        return (-mean).exp();
    }
    (-stirling_error(k) - deviance(k, mean)).exp() / TAU.sqrt() / k.sqrt()
}

/// the probability of `failures` failures before the `successes`th success.
fn negbinom_pmf(failures: f64, successes: f64, p: f64) -> f64 {
    successes / (failures + successes) * binomial_pmf(successes, failures + successes, p)
}

/// the probability of `k` successes in a sample of `n` drawn without
/// replacement from `total` items of which `good` are successes, as a ratio of
/// binomial probabilities at p = n / total.
fn hypgeom_pmf(k: f64, n: f64, good: f64, total: f64) -> f64 {
    if total <= EXACT_COUNTS {
        return exact_choose(good, k) * exact_choose(total - good, n - k) / exact_choose(total, n);
    }
    let p = n / total;
    binomial_pmf(k, good, p) * binomial_pmf(n - k, total - good, p) / binomial_pmf(n, total, p)
}

/// the longest walk one tail sum may take.
const MAX_STEPS: u64 = 20_000_000;

/// steps a [`Meter`] takes between charges to the recalculation budget.
const CHARGE_EVERY: u64 = 4096;

/// the inner steps of one call, charged to the shared recalculation budget
/// every [`CHARGE_EVERY`] steps and settled at the end, so the work of many
/// calls is bounded together and a refused charge stops the walk with `#NUM!`.
struct Meter<'c, 'a> {
    ctx: &'c EvalContext<'a>,
    steps: u64,
}

impl<'c, 'a> Meter<'c, 'a> {
    /// a meter for a call; `#NUM!` once this evaluation has run out of budget.
    fn new(ctx: &'c EvalContext<'a>) -> Result<Self, ErrorValue> {
        ctx.within_budget()?;
        Ok(Self { ctx, steps: 0 })
    }

    fn step(&mut self) -> Result<(), ErrorValue> {
        self.steps += 1;
        if self.steps.is_multiple_of(CHARGE_EVERY) {
            self.pay(CHARGE_EVERY)?;
        }
        Ok(())
    }

    fn pay(&self, steps: u64) -> Result<(), ErrorValue> {
        if self.ctx.consume_cells(crate::array::step_units(steps)) {
            Ok(())
        } else {
            Err(ErrorValue::Num)
        }
    }

    /// charge the steps since the last full block and pass `value` on.
    fn settle(self, value: f64) -> Result<f64, ErrorValue> {
        self.pay(self.steps % CHARGE_EVERY)?;
        Ok(value)
    }
}

/// whether a tail can be summed term by term: a spread below 1e5 standard
/// deviations keeps the walk short, and a mode below 1e15 keeps every whole
/// number the walk visits exact in f64.
fn summable(variance: f64, mode: f64) -> bool {
    variance < 1e10 && mode < 1e15
}

/// a count argument: a negative value is `#NUM!` before its fraction is
/// dropped, as excel checks it.
fn count(value: f64) -> Result<f64, ErrorValue> {
    if value < 0.0 {
        return Err(ErrorValue::Num);
    }
    Ok(value.trunc())
}

/// the mass on the whole numbers `a..=b`, term by term.
fn mass_between(
    a: f64,
    b: f64,
    pmf: impl Fn(f64) -> f64,
    meter: &mut Meter<'_, '_>,
) -> Result<f64, ErrorValue> {
    let steps = (b - a) as u64;
    let mut total = 0.0;
    for i in 0..=steps {
        meter.step()?;
        total += pmf(a + i as f64);
    }
    Ok(total)
}

/// a unimodal probability mass on the whole numbers `lo..=hi`. `ratio(j,
/// down)` is pmf(j - 1) / pmf(j) when `down`, else pmf(j + 1) / pmf(j).
struct Discrete<F: Fn(f64) -> f64, R: Fn(f64, bool) -> f64> {
    pmf: F,
    ratio: R,
    lo: f64,
    hi: f64,
    mode: f64,
}

impl<F: Fn(f64) -> f64, R: Fn(f64, bool) -> f64> Discrete<F, R> {
    /// the mass from `from` stepping away from the mode to `to`, stopping once
    /// the terms no longer move the total. terms come from the ratio, with an
    /// exact mass every 256 steps so rounding cannot drift, and a compensated
    /// total so a long walk keeps its last digits.
    fn tail(&self, from: f64, to: f64, meter: &mut Meter<'_, '_>) -> Result<f64, ErrorValue> {
        let down = to < from;
        let steps = (to - from).abs();
        let at = |taken: u64| {
            if down {
                from - taken as f64
            } else {
                from + taken as f64
            }
        };
        let mut term = (self.pmf)(from);
        let mut sum = term;
        let mut lost = 0.0;
        let mut taken: u64 = 0;
        while (taken as f64) < steps && taken < MAX_STEPS && term > sum * 1e-17 {
            meter.step()?;
            let j = at(taken);
            taken += 1;
            term = if taken.is_multiple_of(256) {
                (self.pmf)(at(taken))
            } else {
                term * (self.ratio)(j, down)
            };
            let total = sum + term;
            lost += if sum >= term {
                (sum - total) + term
            } else {
                (term - total) + sum
            };
            sum = total;
        }
        Ok(sum + lost)
    }

    /// P(X <= k), summed over whichever side of the mode `k` falls on.
    fn lower(&self, k: f64, meter: &mut Meter<'_, '_>) -> Result<f64, ErrorValue> {
        Ok(if k < self.lo {
            0.0
        } else if k >= self.hi {
            1.0
        } else if k <= self.mode {
            self.tail(k, self.lo, meter)?
        } else {
            1.0 - self.tail(k + 1.0, self.hi, meter)?
        })
    }

    /// P(X >= k).
    fn upper(&self, k: f64, meter: &mut Meter<'_, '_>) -> Result<f64, ErrorValue> {
        Ok(if k <= self.lo {
            1.0
        } else if k > self.hi {
            0.0
        } else if k > self.mode {
            self.tail(k, self.hi, meter)?
        } else {
            1.0 - self.tail(k - 1.0, self.lo, meter)?
        })
    }

    /// P(a <= X <= b) from the tails, for a range too wide to sum directly.
    fn between(&self, a: f64, b: f64, meter: &mut Meter<'_, '_>) -> Result<f64, ErrorValue> {
        Ok(if a > self.mode {
            self.upper(a, meter)? - self.upper(b + 1.0, meter)?
        } else {
            self.lower(b, meter)? - self.lower(a - 1.0, meter)?
        })
    }
}

/// the binomial distribution for 0 < p < 1.
fn binomial(n: f64, p: f64) -> Discrete<impl Fn(f64) -> f64, impl Fn(f64, bool) -> f64> {
    let q = 1.0 - p;
    Discrete {
        pmf: move |k| binomial_pmf(k, n, p),
        ratio: move |j, down| {
            if down {
                j * q / ((n - j + 1.0) * p)
            } else {
                (n - j) * p / ((j + 1.0) * q)
            }
        },
        lo: 0.0,
        hi: n,
        mode: binomial_mode(n, p),
    }
}

fn binomial_mode(n: f64, p: f64) -> f64 {
    ((n + 1.0) * p).floor().min(n)
}

/// whether the binomial tails can be summed; beyond that the cumulative
/// functions give `#NUM!`.
fn binomial_summable(n: f64, p: f64) -> Result<(), ErrorValue> {
    if !summable(n * p * (1.0 - p), binomial_mode(n, p)) {
        return Err(ErrorValue::Num);
    }
    Ok(())
}

/// P(X <= k) for X binomial over `n` trials.
fn binomial_cdf(k: f64, n: f64, p: f64, meter: &mut Meter<'_, '_>) -> Result<f64, ErrorValue> {
    if k < 0.0 {
        return Ok(0.0);
    }
    if k >= n || p == 0.0 {
        return Ok(1.0);
    }
    if p == 1.0 {
        return Ok(0.0);
    }
    binomial_summable(n, p)?;
    binomial(n, p).lower(k, meter)
}

/// trials and success probability, checked as every binomial function checks
/// them.
fn binomial_domain(n: f64, p: f64) -> Result<(), ErrorValue> {
    if n < 0.0 || !(0.0..=1.0).contains(&p) {
        return Err(ErrorValue::Num);
    }
    Ok(())
}

/// BINOM.DIST and BINOMDIST(number_s, trials, probability_s, cumulative).
pub(crate) fn binom_dist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 4, 4)?;
        let v = numbers(args, ctx, 3)?;
        let cumulative = flag(args, ctx, 3)?;
        let (k, n, p) = (count(v[0])?, count(v[1])?, v[2]);
        binomial_domain(n, p)?;
        if k > n {
            return Err(ErrorValue::Num);
        }
        if cumulative {
            let mut meter = Meter::new(ctx)?;
            let value = binomial_cdf(k, n, p, &mut meter)?;
            meter.settle(value)
        } else {
            Ok(binomial_pmf(k, n, p))
        }
    })())
}

/// BINOM.DIST.RANGE(trials, probability_s, number_s, [number_s2]): the
/// probability of between `number_s` and `number_s2` successes inclusive, or
/// of exactly `number_s` when the upper bound is omitted.
pub(crate) fn binom_dist_range(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 3, 4)?;
        let v = numbers(args, ctx, 3)?;
        let (n, p, a) = (count(v[0])?, v[1], count(v[2])?);
        let b = if args.len() == 4 && !omitted(&args[3]) {
            count(nth_number(args, ctx, 3)?)?
        } else {
            a
        };
        binomial_domain(n, p)?;
        if a > n || b < a || b > n {
            return Err(ErrorValue::Num);
        }
        let mut meter = Meter::new(ctx)?;
        let value = if b - a <= 1e5 {
            mass_between(a, b, |j| binomial_pmf(j, n, p), &mut meter)?
        } else if p == 0.0 || p == 1.0 {
            binomial_cdf(b, n, p, &mut meter)? - binomial_cdf(a - 1.0, n, p, &mut meter)?
        } else {
            binomial_summable(n, p)?;
            binomial(n, p).between(a, b, &mut meter)?
        };
        meter.settle(value)
    })())
}

/// the largest count a BINOM.INV search visits: past every mode that can be
/// summed, and exact in f64.
const SEARCH_TOP: f64 = 9e15;

/// BINOM.INV and CRITBINOM(trials, probability_s, alpha): the smallest
/// number of successes whose cumulative probability reaches `alpha`. an
/// alpha of one is reached only by every trial; above one half the upper
/// tail is compared with 1 - alpha, which is exact there, so a cumulative
/// that rounds up to alpha does not end the search early. the search starts
/// at the normal approximation and widens until it brackets the answer.
pub(crate) fn binom_inv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 3, 3)?;
        let v = numbers(args, ctx, 3)?;
        let (n, p, alpha) = (count(v[0])?, v[1], v[2]);
        binomial_domain(n, p)?;
        if !(0.0..=1.0).contains(&alpha) {
            return Err(ErrorValue::Num);
        }
        if alpha == 0.0 || p == 0.0 {
            return Ok(0.0);
        }
        if alpha == 1.0 || p == 1.0 {
            return Ok(n);
        }
        binomial_summable(n, p)?;
        let mut meter = Meter::new(ctx)?;
        let dist = binomial(n, p);
        let reaches = |k: u64, meter: &mut Meter<'_, '_>| -> Result<bool, ErrorValue> {
            let k = k as f64;
            Ok(if alpha > 0.5 {
                dist.upper(k + 1.0, meter)? <= 1.0 - alpha
            } else {
                dist.lower(k, meter)? >= alpha
            })
        };
        let q = 1.0 - p;
        let z = if alpha > 0.5 {
            -special::norm_inv_lower(1.0 - alpha)
        } else {
            special::norm_inv_lower(alpha)
        };
        meter.pay(special::take_steps())?;
        let top = n.min(SEARCH_TOP) as u64;
        let guess = (n * p + z * (n * p * q).sqrt() + (q - p) * (z * z - 1.0) / 6.0 - 0.5)
            .round()
            .clamp(0.0, top as f64) as u64;
        let mut step: u64 = 1;
        let (mut lo, mut hi);
        if reaches(guess, &mut meter)? {
            hi = guess;
            lo = loop {
                if hi == 0 {
                    break 0;
                }
                let probe = hi.saturating_sub(step);
                if !reaches(probe, &mut meter)? {
                    break probe + 1;
                }
                hi = probe;
                step = step.saturating_mul(2);
            };
        } else {
            lo = guess + 1;
            hi = loop {
                let probe = lo.saturating_add(step - 1).min(top);
                if reaches(probe, &mut meter)? {
                    break probe;
                }
                if probe == top {
                    return Err(ErrorValue::Num);
                }
                lo = probe + 1;
                step = step.saturating_mul(2);
            };
        }
        while lo < hi {
            let mid = lo + (hi - lo) / 2;
            if reaches(mid, &mut meter)? {
                hi = mid;
            } else {
                lo = mid + 1;
            }
        }
        meter.settle(lo as f64)
    })())
}

fn negbinom_impl(
    ctx: &EvalContext<'_>,
    failures: f64,
    successes: f64,
    p: f64,
    cumulative: bool,
) -> Result<f64, ErrorValue> {
    if !(0.0..=1.0).contains(&p) || successes < 1.0 {
        return Err(ErrorValue::Num);
    }
    if p == 0.0 {
        return Ok(0.0);
    }
    if !cumulative {
        return Ok(negbinom_pmf(failures, successes, p));
    }
    if p == 1.0 {
        return Ok(1.0);
    }
    let q = 1.0 - p;
    let mode = ((successes - 1.0) * q / p).floor().max(0.0);
    if !summable(successes * q / (p * p), mode) {
        return Err(ErrorValue::Num);
    }
    let dist = Discrete {
        pmf: |f| negbinom_pmf(f, successes, p),
        ratio: |f, down| {
            if down {
                f / ((f + successes - 1.0) * q)
            } else {
                (f + successes) * q / (f + 1.0)
            }
        },
        lo: 0.0,
        hi: f64::INFINITY,
        mode,
    };
    let mut meter = Meter::new(ctx)?;
    let value = dist.lower(failures, &mut meter)?;
    meter.settle(value)
}

/// NEGBINOM.DIST(number_f, number_s, probability_s, cumulative).
pub(crate) fn negbinom_dist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 4, 4)?;
        let v = numbers(args, ctx, 3)?;
        let cumulative = flag(args, ctx, 3)?;
        negbinom_impl(ctx, count(v[0])?, count(v[1])?, v[2], cumulative)
    })())
}

/// NEGBINOMDIST(number_f, number_s, probability_s): the mass function only.
pub(crate) fn negbinomdist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 3, 3)?;
        let v = numbers(args, ctx, 3)?;
        negbinom_impl(ctx, count(v[0])?, count(v[1])?, v[2], false)
    })())
}

fn hypgeom_impl(
    ctx: &EvalContext<'_>,
    k: f64,
    n: f64,
    good: f64,
    total: f64,
    cumulative: bool,
) -> Result<f64, ErrorValue> {
    if total <= 0.0 || n <= 0.0 || n > total || good <= 0.0 || good > total {
        return Err(ErrorValue::Num);
    }
    let lo = (n - total + good).max(0.0);
    let hi = n.min(good);
    if k < lo || k > hi {
        return Err(ErrorValue::Num);
    }
    if !cumulative {
        return Ok(hypgeom_pmf(k, n, good, total));
    }
    if k >= hi {
        return Ok(1.0);
    }
    let bad = total - good;
    let variance = n * good * bad * (total - n) / (total * total * (total - 1.0));
    let mode = ((n + 1.0) * (good + 1.0) / (total + 2.0))
        .floor()
        .clamp(lo, hi);
    if !summable(variance, mode) {
        return Err(ErrorValue::Num);
    }
    let dist = Discrete {
        pmf: |j| hypgeom_pmf(j, n, good, total),
        ratio: |j, down| {
            if down {
                j * (bad - n + j) / ((good - j + 1.0) * (n - j + 1.0))
            } else {
                (good - j) * (n - j) / ((j + 1.0) * (bad - n + j + 1.0))
            }
        },
        lo,
        hi,
        mode,
    };
    let mut meter = Meter::new(ctx)?;
    let value = dist.lower(k, &mut meter)?;
    meter.settle(value)
}

/// HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop,
/// cumulative).
pub(crate) fn hypgeom_dist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 5, 5)?;
        let v = numbers(args, ctx, 4)?;
        let cumulative = flag(args, ctx, 4)?;
        hypgeom_impl(
            ctx,
            count(v[0])?,
            count(v[1])?,
            count(v[2])?,
            count(v[3])?,
            cumulative,
        )
    })())
}

/// HYPGEOMDIST(sample_s, number_sample, population_s, number_pop): the mass
/// function only.
pub(crate) fn hypgeomdist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 4, 4)?;
        let v = numbers(args, ctx, 4)?;
        hypgeom_impl(
            ctx,
            count(v[0])?,
            count(v[1])?,
            count(v[2])?,
            count(v[3])?,
            false,
        )
    })())
}

/// POISSON.DIST and POISSON(x, mean, cumulative).
pub(crate) fn poisson_dist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 3, 3)?;
        let v = numbers(args, ctx, 2)?;
        let cumulative = flag(args, ctx, 2)?;
        let (k, mean) = (count(v[0])?, v[1]);
        if mean < 0.0 {
            return Err(ErrorValue::Num);
        }
        if !cumulative {
            return Ok(poisson_pmf(k, mean));
        }
        if !summable(mean, mean) {
            return Err(ErrorValue::Num);
        }
        let dist = Discrete {
            pmf: |j| poisson_pmf(j, mean),
            ratio: |j, down| if down { j / mean } else { mean / (j + 1.0) },
            lo: 0.0,
            hi: f64::INFINITY,
            mode: mean.floor(),
        };
        let mut meter = Meter::new(ctx)?;
        let value = dist.lower(k, &mut meter)?;
        meter.settle(value)
    })())
}

/// EXPON.DIST and EXPONDIST(x, lambda, cumulative).
pub(crate) fn expon_dist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 3, 3)?;
        let v = numbers(args, ctx, 2)?;
        let cumulative = flag(args, ctx, 2)?;
        let (x, lambda) = (v[0], v[1]);
        if x < 0.0 || lambda <= 0.0 {
            return Err(ErrorValue::Num);
        }
        Ok(if cumulative {
            -(-lambda * x).exp_m1()
        } else {
            lambda * (-lambda * x).exp()
        })
    })())
}

/// WEIBULL.DIST and WEIBULL(x, alpha, beta, cumulative). when x / beta or
/// the density leaves the normal range the terms are taken as logarithms, so
/// a density that underflows is zero rather than infinity times zero; at
/// x = 0 the density is infinite for alpha < 1 (`#NUM!`), 1 / beta for
/// alpha = 1 and zero beyond.
pub(crate) fn weibull_dist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 4, 4)?;
        let v = numbers(args, ctx, 3)?;
        let cumulative = flag(args, ctx, 3)?;
        let (x, alpha, beta) = (v[0], v[1], v[2]);
        if x < 0.0 || alpha <= 0.0 || beta <= 0.0 {
            return Err(ErrorValue::Num);
        }
        let ratio = x / beta;
        if ratio.is_normal() {
            let scaled = ratio.powf(alpha);
            if cumulative {
                return Ok(-(-scaled).exp_m1());
            }
            let density = alpha / beta * ratio.powf(alpha - 1.0) * (-scaled).exp();
            if density.is_normal() {
                return Ok(density);
            }
        }
        if x == 0.0 {
            return match (cumulative, alpha.total_cmp(&1.0)) {
                (true, _) | (false, std::cmp::Ordering::Greater) => Ok(0.0),
                (false, std::cmp::Ordering::Equal) => Ok(1.0 / beta),
                (false, std::cmp::Ordering::Less) => Err(ErrorValue::Num),
            };
        }
        let log_ratio = x.ln() - beta.ln();
        let scaled = (alpha * log_ratio).exp();
        if cumulative {
            return Ok(-(-scaled).exp_m1());
        }
        if scaled.is_infinite() {
            return Ok(0.0);
        }
        Ok((alpha.ln() - beta.ln() + (alpha - 1.0) * log_ratio - scaled).exp())
    })())
}

/// FISHER(x): atanh(x) for -1 < x < 1.
pub(crate) fn fisher(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 1, 1)?;
        let x = nth_number(args, ctx, 0)?;
        if x <= -1.0 || x >= 1.0 {
            return Err(ErrorValue::Num);
        }
        Ok(x.atanh())
    })())
}

/// FISHERINV(y): tanh(y).
pub(crate) fn fisherinv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 1, 1)?;
        Ok(nth_number(args, ctx, 0)?.tanh())
    })())
}

/// STANDARDIZE(x, mean, standard_dev).
pub(crate) fn standardize(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 3, 3)?;
        let v = numbers(args, ctx, 3)?;
        if v[2] <= 0.0 {
            return Err(ErrorValue::Num);
        }
        Ok((v[0] - v[1]) / v[2])
    })())
}

/// PERMUT(number, number_chosen): n! / (n - k)!; microsoft requires n > 0.
pub(crate) fn permut(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 2, 2)?;
        let v = numbers(args, ctx, 2)?;
        let (n, k) = (count(v[0])?, count(v[1])?);
        if n <= 0.0 || n < k {
            return Err(ErrorValue::Num);
        }
        let mut meter = Meter::new(ctx)?;
        let mut product = 1.0_f64;
        for i in 0..k as u64 {
            if !product.is_finite() {
                break;
            }
            meter.step()?;
            product *= n - i as f64;
        }
        meter.settle(product)
    })())
}

/// PERMUTATIONA(number, number_chosen): n^k, repetition allowed.
pub(crate) fn permutationa(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 2, 2)?;
        let v = numbers(args, ctx, 2)?;
        Ok(count(v[0])?.powf(count(v[1])?))
    })())
}

/// the largest power of two at or below the largest magnitude in `values`,
/// so dividing by it is exact; one when every value is zero or subnormal.
fn power_scale(values: &[f64]) -> f64 {
    let largest = values.iter().fold(0.0_f64, |m, v| m.max(v.abs()));
    if largest.is_normal() {
        f64::from_bits(largest.to_bits() & 0x7ff0_0000_0000_0000)
    } else {
        1.0
    }
}

/// STEYX(known_y's, known_x's): the standard error of the predicted y,
/// pairing the arguments as CORREL does; fewer than three pairs is
/// `#DIV/0!`. the residuals are taken from the fitted line rather than as a
/// difference of sums, so large nearly collinear values keep their error,
/// and both samples are scaled by powers of two so squares neither overflow
/// nor underflow.
pub(crate) fn steyx(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        let (ys, xs) = stats::pairs(args, ctx)?;
        if ys.len() < 3 {
            return Err(ErrorValue::Div0);
        }
        let (sx, sy) = (power_scale(&xs), power_scale(&ys));
        let us: Vec<f64> = xs.iter().map(|x| x / sx).collect();
        let vs: Vec<f64> = ys.iter().map(|y| y / sy).collect();
        let (n, mu, mv, suu, _, suv) = stats::moments(&vs, &us);
        if suu == 0.0 {
            return Err(ErrorValue::Div0);
        }
        let slope = suv / suu;
        let residuals: f64 = us
            .iter()
            .zip(&vs)
            .map(|(u, v)| {
                let r = (v - mv) - slope * (u - mu);
                r * r
            })
            .sum();
        Ok(sy * (residuals / (n - 2.0)).sqrt())
    })())
}

/// SKEW.P(number, ...): the population skewness; microsoft gives `#DIV/0!`
/// for fewer than three values or no spread.
pub(crate) fn skew_p(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        let nums = collect_numbers_deep(args, ctx)?;
        if nums.len() < 3 {
            return Err(ErrorValue::Div0);
        }
        let n = nums.len() as f64;
        let mean = nums.iter().sum::<f64>() / n;
        let variance = nums.iter().map(|x| (x - mean) * (x - mean)).sum::<f64>() / n;
        if variance <= 0.0 {
            return Err(ErrorValue::Div0);
        }
        let sd = variance.sqrt();
        Ok(nums.iter().map(|x| ((x - mean) / sd).powi(3)).sum::<f64>() / n)
    })())
}

/// the variance of the `-A` values: text in references as zero, logicals as
/// one or zero.
fn variance_a(args: &[Expr], ctx: &EvalContext<'_>, sample: bool) -> Result<f64, ErrorValue> {
    let nums = collect_numbers_anytype(args, ctx)?;
    let lost = usize::from(sample);
    if nums.len() <= lost {
        return Err(ErrorValue::Div0);
    }
    let n = nums.len() as f64;
    let mean = nums.iter().sum::<f64>() / n;
    let total: f64 = nums.iter().map(|x| (x - mean) * (x - mean)).sum();
    Ok(total / (nums.len() - lost) as f64)
}

/// VARA(value, ...): the sample variance with text as zero.
pub(crate) fn vara(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer(variance_a(args, ctx, true))
}

/// VARPA(value, ...): the population variance with text as zero.
pub(crate) fn varpa(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer(variance_a(args, ctx, false))
}

/// STDEVPA(value, ...): the population standard deviation with text as zero.
pub(crate) fn stdevpa(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer(variance_a(args, ctx, false).map(f64::sqrt))
}

/// the sorted numbers of an array argument; none at all is `#NUM!`.
fn sorted_numbers(arg: &Expr, ctx: &EvalContext<'_>) -> Result<Vec<f64>, ErrorValue> {
    let mut nums = collect_numbers(std::slice::from_ref(arg), ctx)?;
    if nums.is_empty() {
        return Err(ErrorValue::Num);
    }
    nums.sort_by(f64::total_cmp);
    Ok(nums)
}

/// the exclusive percentile: rank k (n + 1), interpolated between order
/// statistics; a rank outside 1..=n is `#NUM!`.
fn percentile_exc(sorted: &[f64], k: f64) -> Result<f64, ErrorValue> {
    let n = sorted.len() as f64;
    let rank = k * (n + 1.0);
    if k <= 0.0 || k >= 1.0 || rank < 1.0 || rank > n {
        return Err(ErrorValue::Num);
    }
    let lower = rank.floor() as usize;
    let fraction = rank - rank.floor();
    if lower == sorted.len() {
        return Ok(sorted[lower - 1]);
    }
    Ok(sorted[lower - 1] + fraction * (sorted[lower] - sorted[lower - 1]))
}

/// PERCENTILE.EXC(array, k) for 0 < k < 1.
pub(crate) fn percentile_exc_fn(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 2, 2)?;
        let sorted = sorted_numbers(&args[0], ctx)?;
        percentile_exc(&sorted, nth_number(args, ctx, 1)?)
    })())
}

/// QUARTILE.EXC(array, quart): quart truncates and must be 1, 2 or 3.
pub(crate) fn quartile_exc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 2, 2)?;
        let sorted = sorted_numbers(&args[0], ctx)?;
        let quart = nth_number(args, ctx, 1)?.trunc();
        if quart <= 0.0 || quart >= 4.0 {
            return Err(ErrorValue::Num);
        }
        percentile_exc(&sorted, quart / 4.0)
    })())
}

/// cuts `value` down (never rounds) to `digits` decimals; a product that
/// lands within rounding of a whole number counts as that number, so 0.7 keeps
/// three digits as 0.7.
fn truncate_digits(value: f64, digits: f64) -> f64 {
    let scale = 10f64.powi(digits.min(308.0) as i32);
    let scaled = value * scale;
    if !scaled.is_finite() {
        return value;
    }
    let nearest = scaled.round();
    let whole = if (scaled - nearest).abs() <= 1e-12 * nearest.abs().max(1.0) {
        nearest
    } else {
        scaled.floor()
    };
    whole / scale
}

fn percent_rank(args: &[Expr], ctx: &EvalContext<'_>, exclusive: bool) -> CellValue {
    answer((|| {
        arity(args, 2, 3)?;
        let sorted = sorted_numbers(&args[0], ctx)?;
        let x = nth_number(args, ctx, 1)?;
        let digits = if args.len() == 3 && !omitted(&args[2]) {
            nth_number(args, ctx, 2)?.trunc()
        } else {
            3.0
        };
        if digits < 1.0 {
            return Err(ErrorValue::Num);
        }
        let n = sorted.len();
        if x < sorted[0] || x > sorted[n - 1] {
            return Err(ErrorValue::NA);
        }
        let below = sorted.partition_point(|v| *v < x);
        let rank = if sorted[below] == x {
            below as f64 + 1.0
        } else {
            let (lo, hi) = (sorted[below - 1], sorted[below]);
            below as f64 + (x - lo) / (hi - lo)
        };
        let value = if exclusive {
            rank / (n as f64 + 1.0)
        } else if n == 1 {
            1.0
        } else {
            (rank - 1.0) / (n as f64 - 1.0)
        };
        Ok(truncate_digits(value, digits))
    })())
}

/// PERCENTRANK and PERCENTRANK.INC(array, x, [significance]): the rank of x
/// as (position - 1) / (n - 1), interpolated between neighbours and cut to
/// `significance` digits (3 by default).
pub(crate) fn percentrank_inc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    percent_rank(args, ctx, false)
}

/// PERCENTRANK.EXC(array, x, [significance]): position / (n + 1).
pub(crate) fn percentrank_exc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    percent_rank(args, ctx, true)
}

/// MODE.MULT's answer: every value that occurs most often, in order of first
/// appearance; nothing repeated is `#N/A`.
pub(crate) fn modes(args: &[Expr], ctx: &EvalContext<'_>) -> Result<Vec<f64>, ErrorValue> {
    let nums = collect_numbers(args, ctx)?;
    let mut counts: HashMap<u64, usize> = HashMap::new();
    for x in &nums {
        *counts.entry((x + 0.0).to_bits()).or_default() += 1;
    }
    let best = counts.values().copied().max().unwrap_or(0);
    if best < 2 {
        return Err(ErrorValue::NA);
    }
    let mut out = Vec::new();
    for x in nums {
        let key = (x + 0.0).to_bits();
        if counts.get(&key) == Some(&best) {
            counts.remove(&key);
            out.push(x);
        }
    }
    Ok(out)
}

/// the numbers of one PROB range by position, `None` where a cell holds no
/// number.
fn positioned(arg: &Expr, ctx: &EvalContext<'_>) -> Result<Vec<Option<f64>>, ErrorValue> {
    let (cells, _) = argument_cells(arg, ctx)?;
    cells
        .into_iter()
        .map(|cell| match cell {
            CellValue::Number { value } => Ok(Some(value)),
            CellValue::Error { value } => Err(value),
            _ => Ok(None),
        })
        .collect()
}

/// PROB(x_range, prob_range, lower_limit, [upper_limit]): the probability
/// mass between the limits inclusive, or at `lower_limit` alone. every
/// probability must lie in (0, 1] and they must sum to one.
pub(crate) fn prob(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 3, 4)?;
        let xs = positioned(&args[0], ctx)?;
        let ps = positioned(&args[1], ctx)?;
        if xs.len() != ps.len() {
            return Err(ErrorValue::NA);
        }
        let lower = nth_number(args, ctx, 2)?;
        let upper = if args.len() == 4 && !omitted(&args[3]) {
            nth_number(args, ctx, 3)?
        } else {
            lower
        };
        let pairs: Vec<(f64, f64)> = xs
            .into_iter()
            .zip(ps)
            .filter_map(|(x, p)| Some((x?, p?)))
            .collect();
        if pairs.iter().any(|(_, p)| *p <= 0.0 || *p > 1.0) {
            return Err(ErrorValue::Num);
        }
        let total: f64 = pairs.iter().map(|(_, p)| p).sum();
        if (total - 1.0).abs() > 1e-9 {
            return Err(ErrorValue::Num);
        }
        Ok(pairs
            .iter()
            .filter(|(x, _)| (lower..=upper).contains(x))
            .map(|(_, p)| p)
            .sum())
    })())
}

/// Z.TEST and ZTEST(array, x, [sigma]): the one-tailed p-value
/// 1 - NORM.S.DIST((mean - x) / (sigma / sqrt(n))), with the sample
/// standard deviation when sigma is omitted.
pub(crate) fn z_test(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    answer((|| {
        arity(args, 2, 3)?;
        let nums = collect_numbers(&args[..1], ctx)?;
        let x = nth_number(args, ctx, 1)?;
        let sigma = if args.len() == 3 && !omitted(&args[2]) {
            Some(nth_number(args, ctx, 2)?)
        } else {
            None
        };
        if nums.is_empty() {
            return Err(ErrorValue::NA);
        }
        let n = nums.len() as f64;
        let mean = nums.iter().sum::<f64>() / n;
        let sd = match sigma {
            Some(sigma) if sigma <= 0.0 => return Err(ErrorValue::Num),
            Some(sigma) => sigma,
            None => {
                if nums.len() < 2 {
                    return Err(ErrorValue::Div0);
                }
                let total: f64 = nums.iter().map(|v| (v - mean) * (v - mean)).sum();
                (total / (n - 1.0)).sqrt()
            }
        };
        if sd == 0.0 {
            return Err(ErrorValue::Div0);
        }
        Ok(special::norm_cdf(-(mean - x) / (sd / n.sqrt())))
    })())
}
