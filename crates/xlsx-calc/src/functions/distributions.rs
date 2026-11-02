//! continuous distributions and special functions: gamma, beta, error,
//! normal, lognormal, chi-square, t and f families, their inverses, the
//! chi-square, t and f hypothesis tests, and the confidence intervals.
//! semantics follow microsoft's function reference: integer arguments such
//! as degrees of freedom are truncated, out-of-domain arguments give
//! `#NUM!`, non-numeric arguments give `#VALUE!`, and the compatibility
//! names share the modern implementations (the old beta and lognormal
//! distribution names have no cumulative flag and are always cumulative).

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{Area, EvalContext, as_area, err, evaluate, to_bool};
use crate::fmath;
use crate::parser::Expr;

use super::criteria;
use super::special;
use super::{finite, nth_int, nth_number, omitted};

fn numbers(args: &[Expr], ctx: &EvalContext<'_>, count: usize) -> Result<Vec<f64>, ErrorValue> {
    if args.len() != count {
        return Err(ErrorValue::Value);
    }
    first_numbers(args, ctx, count)
}

fn first_numbers(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    count: usize,
) -> Result<Vec<f64>, ErrorValue> {
    (0..count).map(|i| nth_number(args, ctx, i)).collect()
}

fn number_range(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    min: usize,
    max: usize,
) -> Result<Vec<f64>, ErrorValue> {
    if args.len() < min || args.len() > max {
        return Err(ErrorValue::Value);
    }
    (0..args.len()).map(|i| nth_number(args, ctx, i)).collect()
}

fn flag(args: &[Expr], ctx: &EvalContext<'_>, i: usize) -> Result<bool, ErrorValue> {
    to_bool(&evaluate(&args[i], ctx))
}

/// the result, after charging the iterations the special functions ran to
/// the recalculation budget.
fn finish(ctx: &EvalContext<'_>, value: f64) -> CellValue {
    if !ctx.consume_cells(crate::array::step_units(special::take_steps())) {
        return err(ErrorValue::Num);
    }
    finite(value)
}

/// GAMMA(number): the gamma function; zero and negative integers are poles.
pub(crate) fn gamma(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let v = match numbers(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let z = v[0];
    if z <= 0.0 && z == z.trunc() {
        return err(ErrorValue::Num);
    }
    finish(ctx, special::gamma_fn(z))
}

/// GAMMALN and GAMMALN.PRECISE(x): ln Gamma(x) for x > 0.
pub(crate) fn gammaln(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let v = match numbers(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if v[0] <= 0.0 {
        return err(ErrorValue::Num);
    }
    finish(ctx, special::lngamma(v[0]))
}

fn gamma_dist_impl(x: f64, alpha: f64, beta: f64, cumulative: bool) -> Result<f64, ErrorValue> {
    if x < 0.0 || alpha <= 0.0 || beta <= 0.0 {
        return Err(ErrorValue::Num);
    }
    Ok(if cumulative {
        special::gamma_cdf(x, alpha, beta)
    } else {
        special::gamma_pdf(x, alpha, beta)
    })
}

/// GAMMA.DIST and GAMMADIST(x, alpha, beta, cumulative).
pub(crate) fn gamma_dist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 4 {
        return err(ErrorValue::Value);
    }
    let v = match first_numbers(args, ctx, 3) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let cumulative = match flag(args, ctx, 3) {
        Ok(b) => b,
        Err(e) => return err(e),
    };
    match gamma_dist_impl(v[0], v[1], v[2], cumulative) {
        Ok(value) => finish(ctx, value),
        Err(e) => err(e),
    }
}

/// GAMMA.INV and GAMMAINV(probability, alpha, beta).
pub(crate) fn gamma_inv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let v = match numbers(args, ctx, 3) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let (p, alpha, beta) = (v[0], v[1], v[2]);
    if !(0.0..=1.0).contains(&p) || alpha <= 0.0 || beta <= 0.0 {
        return err(ErrorValue::Num);
    }
    if p == 1.0 {
        return err(ErrorValue::Num);
    }
    finish(ctx, beta * special::gamma_quantile(alpha, p, 1.0 - p))
}

fn beta_bounds(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    first: usize,
) -> Result<(f64, f64), ErrorValue> {
    let lo = if args.len() > first && !omitted(&args[first]) {
        nth_number(args, ctx, first)?
    } else {
        0.0
    };
    let hi = if args.len() > first + 1 && !omitted(&args[first + 1]) {
        nth_number(args, ctx, first + 1)?
    } else {
        1.0
    };
    Ok((lo, hi))
}

fn beta_dist_impl(
    x: f64,
    alpha: f64,
    beta: f64,
    cumulative: bool,
    lo: f64,
    hi: f64,
) -> Result<f64, ErrorValue> {
    if alpha <= 0.0 || beta <= 0.0 || hi <= lo || x < lo || x > hi {
        return Err(ErrorValue::Num);
    }
    Ok(if cumulative {
        special::beta_cdf(x, alpha, beta, lo, hi)
    } else {
        special::beta_pdf(x, alpha, beta, lo, hi)
    })
}

/// BETA.DIST(x, alpha, beta, cumulative, [lower], [upper]).
pub(crate) fn beta_dist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 4 || args.len() > 6 {
        return err(ErrorValue::Value);
    }
    let v = match first_numbers(args, ctx, 3) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let cumulative = match flag(args, ctx, 3) {
        Ok(b) => b,
        Err(e) => return err(e),
    };
    let (lo, hi) = match beta_bounds(args, ctx, 4) {
        Ok(b) => b,
        Err(e) => return err(e),
    };
    match beta_dist_impl(v[0], v[1], v[2], cumulative, lo, hi) {
        Ok(value) => finish(ctx, value),
        Err(e) => err(e),
    }
}

/// BETADIST(x, alpha, beta, [lower], [upper]): always cumulative.
pub(crate) fn betadist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 3 || args.len() > 5 {
        return err(ErrorValue::Value);
    }
    let v = match first_numbers(args, ctx, 3) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let (lo, hi) = match beta_bounds(args, ctx, 3) {
        Ok(b) => b,
        Err(e) => return err(e),
    };
    match beta_dist_impl(v[0], v[1], v[2], true, lo, hi) {
        Ok(value) => finish(ctx, value),
        Err(e) => err(e),
    }
}

/// BETA.INV and BETAINV(probability, alpha, beta, [lower], [upper]).
pub(crate) fn beta_inv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 3 || args.len() > 5 {
        return err(ErrorValue::Value);
    }
    let v = match first_numbers(args, ctx, 3) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let (lo, hi) = match beta_bounds(args, ctx, 3) {
        Ok(b) => b,
        Err(e) => return err(e),
    };
    let (p, alpha, beta) = (v[0], v[1], v[2]);
    if !(0.0..=1.0).contains(&p) || alpha <= 0.0 || beta <= 0.0 || hi <= lo {
        return err(ErrorValue::Num);
    }
    let t = special::beta_quantile(alpha, beta, p, 1.0 - p);
    let span = hi - lo;
    finish(
        ctx,
        if span.is_finite() {
            lo + span * t
        } else {
            2.0 * (0.5 * lo + (0.5 * hi - 0.5 * lo) * t)
        },
    )
}

/// ERF(lower, [upper]): the error function, or its difference over a range.
pub(crate) fn erf(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let v = match number_range(args, ctx, 1, 2) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if v.len() == 1 {
        finish(ctx, special::erf_fn(v[0]))
    } else {
        finish(ctx, special::erf_fn(v[1]) - special::erf_fn(v[0]))
    }
}

/// ERF.PRECISE(x): the error function between 0 and x.
pub(crate) fn erf_precise(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let v = match numbers(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    finish(ctx, special::erf_fn(v[0]))
}

/// ERFC and ERFC.PRECISE(x): the complementary error function, for any x.
pub(crate) fn erfc(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let v = match numbers(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    finish(ctx, special::erfc_fn(v[0]))
}

/// NORM.S.DIST(z, cumulative).
pub(crate) fn norm_s_dist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let z = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let cumulative = match flag(args, ctx, 1) {
        Ok(b) => b,
        Err(e) => return err(e),
    };
    finish(
        ctx,
        if cumulative {
            special::norm_cdf(z)
        } else {
            special::norm_pdf(z)
        },
    )
}

/// NORMSDIST(z): the standard normal cdf; no cumulative flag.
pub(crate) fn normsdist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let v = match numbers(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    finish(ctx, special::norm_cdf(v[0]))
}

/// PHI(x): the standard normal density.
pub(crate) fn phi(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let v = match numbers(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    finish(ctx, special::norm_pdf(v[0]))
}

/// GAUSS(z): the probability that a standard normal value falls between the
/// mean and z, NORM.S.DIST(z, TRUE) - 0.5.
pub(crate) fn gauss(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let v = match numbers(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    finish(
        ctx,
        0.5 * special::erf_fn(v[0] * std::f64::consts::FRAC_1_SQRT_2),
    )
}

/// NORM.S.INV and NORMSINV(probability).
pub(crate) fn norm_s_inv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let v = match numbers(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if v[0] <= 0.0 || v[0] >= 1.0 {
        return err(ErrorValue::Num);
    }
    finish(ctx, special::norm_inv_cdf(v[0]))
}

/// NORM.INV and NORMINV(probability, mean, standard_dev).
pub(crate) fn norm_inv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let v = match numbers(args, ctx, 3) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if v[0] <= 0.0 || v[0] >= 1.0 || v[2] <= 0.0 {
        return err(ErrorValue::Num);
    }
    finish(ctx, v[1] + v[2] * special::norm_inv_cdf(v[0]))
}

fn lognorm_cdf(x: f64, mean: f64, sd: f64) -> f64 {
    special::norm_cdf((x.ln() - mean) / sd)
}

fn lognorm_pdf(x: f64, mean: f64, sd: f64) -> f64 {
    let z = (x.ln() - mean) / sd;
    (-0.5 * z * z).exp() / (x * sd * (2.0 * std::f64::consts::PI).sqrt())
}

/// LOGNORM.DIST(x, mean, standard_dev, cumulative).
pub(crate) fn lognorm_dist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 4 {
        return err(ErrorValue::Value);
    }
    let v = match first_numbers(args, ctx, 3) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let cumulative = match flag(args, ctx, 3) {
        Ok(b) => b,
        Err(e) => return err(e),
    };
    if v[0] <= 0.0 || v[2] <= 0.0 {
        return err(ErrorValue::Num);
    }
    finish(
        ctx,
        if cumulative {
            lognorm_cdf(v[0], v[1], v[2])
        } else {
            lognorm_pdf(v[0], v[1], v[2])
        },
    )
}

/// LOGNORMDIST(x, mean, standard_dev): always cumulative.
pub(crate) fn lognormdist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let v = match numbers(args, ctx, 3) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if v[0] <= 0.0 || v[2] <= 0.0 {
        return err(ErrorValue::Num);
    }
    finish(ctx, lognorm_cdf(v[0], v[1], v[2]))
}

/// LOGNORM.INV and LOGINV(probability, mean, standard_dev).
pub(crate) fn lognorm_inv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let v = match numbers(args, ctx, 3) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if v[0] <= 0.0 || v[0] >= 1.0 || v[2] <= 0.0 {
        return err(ErrorValue::Num);
    }
    finish(ctx, fmath::exp(v[1] + v[2] * special::norm_inv_cdf(v[0])))
}

/// a degrees-of-freedom argument, truncated toward zero.
fn degrees_of_freedom(args: &[Expr], ctx: &EvalContext<'_>, i: usize) -> Result<f64, ErrorValue> {
    Ok(nth_number(args, ctx, i)?.trunc())
}

/// the chi-square functions reject more than 10^10 degrees of freedom.
const CHISQ_MAX_DF: f64 = 1e10;

/// the f functions reject more than 10^10 degrees of freedom in either
/// argument, like the chi-square ones.
fn f_df_out_of_range(d1: f64, d2: f64) -> bool {
    d1 > F_MAX_DF || d2 > F_MAX_DF
}

const F_MAX_DF: f64 = 1e10;

/// the t distribution functions are `#NUM!` once `x` squared overflows.
fn t_out_of_range(x: f64) -> bool {
    !(x * x).is_finite()
}

/// CHISQ.DIST(x, deg_freedom, cumulative).
pub(crate) fn chisq_dist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let x = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let df = match degrees_of_freedom(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let cumulative = match flag(args, ctx, 2) {
        Ok(b) => b,
        Err(e) => return err(e),
    };
    if x < 0.0 || df < 1.0 || df > CHISQ_MAX_DF {
        return err(ErrorValue::Num);
    }
    finish(
        ctx,
        if cumulative {
            special::chisq_cdf(x, df)
        } else {
            special::chisq_pdf(x, df)
        },
    )
}

/// CHISQ.DIST.RT and CHIDIST(x, deg_freedom): the right tail.
pub(crate) fn chisq_dist_rt(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let x = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let df = match degrees_of_freedom(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if x < 0.0 || df < 1.0 || df > CHISQ_MAX_DF {
        return err(ErrorValue::Num);
    }
    finish(ctx, special::chisq_tail(x, df))
}

/// CHISQ.INV(probability, deg_freedom): the left-tailed inverse.
pub(crate) fn chisq_inv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let p = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let df = match degrees_of_freedom(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if !(0.0..=1.0).contains(&p) || df < 1.0 || df > CHISQ_MAX_DF || p == 1.0 {
        return err(ErrorValue::Num);
    }
    finish(ctx, 2.0 * special::gamma_quantile(0.5 * df, p, 1.0 - p))
}

/// CHISQ.INV.RT and CHIINV(probability, deg_freedom): the right-tailed inverse.
pub(crate) fn chisq_inv_rt(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let p = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let df = match degrees_of_freedom(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if !(0.0..=1.0).contains(&p) || df < 1.0 || df > CHISQ_MAX_DF || p == 0.0 {
        return err(ErrorValue::Num);
    }
    finish(ctx, 2.0 * special::gamma_quantile(0.5 * df, 1.0 - p, p))
}

/// T.DIST(x, deg_freedom, cumulative).
pub(crate) fn t_dist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let x = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let df = match degrees_of_freedom(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let cumulative = match flag(args, ctx, 2) {
        Ok(b) => b,
        Err(e) => return err(e),
    };
    if df < 1.0 || t_out_of_range(x) {
        return err(ErrorValue::Num);
    }
    finish(
        ctx,
        if cumulative {
            special::t_cdf(x, df)
        } else {
            special::t_pdf(x, df)
        },
    )
}

/// T.DIST.RT(x, deg_freedom).
pub(crate) fn t_dist_rt(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let x = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let df = match degrees_of_freedom(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if df < 1.0 || t_out_of_range(x) {
        return err(ErrorValue::Num);
    }
    finish(ctx, special::t_right(x, df))
}

/// T.DIST.2T(x, deg_freedom): both tails; x must be non-negative.
pub(crate) fn t_dist_2t(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let x = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let df = match degrees_of_freedom(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if x < 0.0 || df < 1.0 || t_out_of_range(x) {
        return err(ErrorValue::Num);
    }
    finish(ctx, 2.0 * special::t_tail(x, df))
}

/// TDIST(x, deg_freedom, tails) with tails truncated to 1 or 2.
pub(crate) fn tdist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let x = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let df = match degrees_of_freedom(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let tails = match nth_int(args, ctx, 2) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if x < 0.0 || df < 1.0 || (tails != 1 && tails != 2) || t_out_of_range(x) {
        return err(ErrorValue::Num);
    }
    let rt = special::t_tail(x, df);
    finish(ctx, if tails == 1 { rt } else { 2.0 * rt })
}

/// T.INV(probability, deg_freedom): the left-tailed inverse.
pub(crate) fn t_inv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let p = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let df = match degrees_of_freedom(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if p <= 0.0 || p >= 1.0 || df < 1.0 {
        return err(ErrorValue::Num);
    }
    finish(
        ctx,
        if p < 0.5 {
            -special::t_upper_quantile(p, df)
        } else {
            special::t_upper_quantile(1.0 - p, df)
        },
    )
}

/// T.INV.2T and TINV(probability, deg_freedom): the two-tailed inverse,
/// returned as a positive value.
pub(crate) fn t_inv_2t(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let p = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let df = match degrees_of_freedom(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if p <= 0.0 || p > 1.0 || df < 1.0 {
        return err(ErrorValue::Num);
    }
    if p == 1.0 {
        return finish(ctx, 0.0);
    }
    finish(ctx, special::t_upper_quantile(0.5 * p, df))
}

/// F.DIST(x, deg_freedom1, deg_freedom2, cumulative).
pub(crate) fn f_dist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 4 {
        return err(ErrorValue::Value);
    }
    let x = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let d1 = match degrees_of_freedom(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let d2 = match degrees_of_freedom(args, ctx, 2) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let cumulative = match flag(args, ctx, 3) {
        Ok(b) => b,
        Err(e) => return err(e),
    };
    if x < 0.0 || d1 < 1.0 || d2 < 1.0 || f_df_out_of_range(d1, d2) {
        return err(ErrorValue::Num);
    }
    finish(
        ctx,
        if cumulative {
            special::f_tails(x, d1, d2).0
        } else {
            special::f_pdf(x, d1, d2)
        },
    )
}

/// F.DIST.RT and FDIST(x, deg_freedom1, deg_freedom2): the right tail.
pub(crate) fn f_dist_rt(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let x = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let d1 = match degrees_of_freedom(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let d2 = match degrees_of_freedom(args, ctx, 2) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if x < 0.0 || d1 < 1.0 || d2 < 1.0 || f_df_out_of_range(d1, d2) {
        return err(ErrorValue::Num);
    }
    finish(ctx, special::f_tails(x, d1, d2).1)
}

/// F.INV(probability, deg_freedom1, deg_freedom2): the left-tailed inverse.
pub(crate) fn f_inv(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let p = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let d1 = match degrees_of_freedom(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let d2 = match degrees_of_freedom(args, ctx, 2) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if !(0.0..=1.0).contains(&p) || d1 < 1.0 || d2 < 1.0 || f_df_out_of_range(d1, d2) || p == 1.0 {
        return err(ErrorValue::Num);
    }
    finish(ctx, special::f_quantile(p, 1.0 - p, d1, d2))
}

/// F.INV.RT and FINV(probability, deg_freedom1, deg_freedom2): the
/// right-tailed inverse.
pub(crate) fn f_inv_rt(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let p = match nth_number(args, ctx, 0) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let d1 = match degrees_of_freedom(args, ctx, 1) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let d2 = match degrees_of_freedom(args, ctx, 2) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if !(0.0..=1.0).contains(&p) || d1 < 1.0 || d2 < 1.0 || f_df_out_of_range(d1, d2) || p == 0.0 {
        return err(ErrorValue::Num);
    }
    finish(ctx, special::f_quantile(1.0 - p, p, d1, d2))
}

/// CONFIDENCE.NORM and CONFIDENCE(alpha, standard_dev, size).
pub(crate) fn confidence_norm(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let v = match first_numbers(args, ctx, 2) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let size = match nth_int(args, ctx, 2) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let (alpha, sd) = (v[0], v[1]);
    if alpha <= 0.0 || alpha >= 1.0 || sd <= 0.0 || size < 1 {
        return err(ErrorValue::Num);
    }
    finish(
        ctx,
        -special::norm_inv_lower(0.5 * alpha) * sd / (size as f64).sqrt(),
    )
}

/// CONFIDENCE.T(alpha, standard_dev, size).
pub(crate) fn confidence_t(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let v = match first_numbers(args, ctx, 2) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let size = match nth_int(args, ctx, 2) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let (alpha, sd) = (v[0], v[1]);
    if alpha <= 0.0 || alpha >= 1.0 || sd <= 0.0 || size < 1 {
        return err(ErrorValue::Num);
    }
    if size == 1 {
        return err(ErrorValue::Div0);
    }
    let critical = special::t_upper_quantile(0.5 * alpha, (size - 1) as f64);
    finish(ctx, critical * sd / (size as f64).sqrt())
}

/// one test argument by position: `None` where a cell is not a number. the
/// shape is the authored one, the cells are read from the extent the two
/// arguments were cut to together.
struct Grid {
    cells: Vec<Option<f64>>,
    rows: usize,
    cols: usize,
}

impl Grid {
    fn numbers(&self) -> Vec<f64> {
        self.cells.iter().flatten().copied().collect()
    }
}

fn grid(
    arg: &Expr,
    authored: Option<Area>,
    cut: Option<Area>,
    ctx: &EvalContext<'_>,
) -> Result<Grid, ErrorValue> {
    let cell = |value: &CellValue| match value {
        CellValue::Number { value } => Ok(Some(*value)),
        CellValue::Error { value } => Err(*value),
        _ => Ok(None),
    };
    if let (Some(authored), Some(cut)) = (authored, cut) {
        let cells = cut
            .cells_ref(ctx)?
            .iter()
            .map(|value| cell(value.as_ref()))
            .collect::<Result<_, _>>()?;
        return Ok(Grid {
            cells,
            rows: authored.rows,
            cols: authored.cols,
        });
    }
    match crate::array::evaluate_array(arg, ctx) {
        crate::array::Value::Array(array) => {
            let cells = array.cells().iter().map(cell).collect::<Result<_, _>>()?;
            Ok(Grid {
                cells,
                rows: array.rows(),
                cols: array.cols(),
            })
        }
        value => Ok(Grid {
            cells: vec![cell(&value.into_scalar())?],
            rows: 1,
            cols: 1,
        }),
    }
}

/// the two sample arguments of a test, references cut to one extent so a
/// whole column on each side still lines up row for row.
fn grids(args: &[Expr], ctx: &EvalContext<'_>) -> Result<(Grid, Grid), ErrorValue> {
    let authored = [as_area(&args[0], ctx), as_area(&args[1], ctx)];
    let mut cut = authored;
    criteria::cut_references(&mut cut, ctx);
    Ok((
        grid(&args[0], authored[0], cut[0], ctx)?,
        grid(&args[1], authored[1], cut[1], ctx)?,
    ))
}

/// the positions where both arguments hold a number, as Excel pairs them.
fn paired(first: &Grid, second: &Grid) -> Vec<(f64, f64)> {
    first
        .cells
        .iter()
        .zip(&second.cells)
        .filter_map(|(a, b)| Some(((*a)?, (*b)?)))
        .collect()
}

/// a power of two near the largest magnitude in both samples. the t and f
/// statistics do not depend on scale, so dividing by it keeps squares of
/// huge or tiny values finite and nonzero without changing any digit.
fn common_scale(first: &Grid, second: &Grid) -> f64 {
    let max = first
        .cells
        .iter()
        .chain(&second.cells)
        .flatten()
        .fold(0.0f64, |m, v| m.max(v.abs()));
    if max > 0.0 && max.is_finite() {
        2f64.powi((max.log2().floor() as i32).clamp(-1000, 1000))
    } else {
        1.0
    }
}

fn scaled(values: Vec<f64>, scale: f64) -> Vec<f64> {
    values.into_iter().map(|v| v / scale).collect()
}

fn sample_stats(values: &[f64]) -> (f64, f64) {
    let n = values.len() as f64;
    let mean = values.iter().sum::<f64>() / n;
    let var = if values.len() > 1 {
        values.iter().map(|v| (v - mean) * (v - mean)).sum::<f64>() / (n - 1.0)
    } else {
        0.0
    };
    (mean, var)
}

/// CHISQ.TEST and CHITEST(actual_range, expected_range). the ranges must
/// share a shape (`#N/A`); a position where either side is not a number is
/// skipped, and the degrees of freedom follow the shape.
pub(crate) fn chisq_test(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let (actual, expected) = match grids(args, ctx) {
        Ok(g) => g,
        Err(e) => return err(e),
    };
    if actual.rows != expected.rows
        || actual.cols != expected.cols
        || actual.cells.len() != expected.cells.len()
    {
        return err(ErrorValue::NA);
    }
    let pairs = paired(&actual, &expected);
    if pairs.iter().any(|(a, e)| *a < 0.0 || *e < 0.0) {
        return err(ErrorValue::Num);
    }
    if pairs.iter().any(|(_, e)| *e == 0.0) {
        return err(ErrorValue::Div0);
    }
    if pairs.is_empty() {
        return err(ErrorValue::NA);
    }
    let df = if actual.rows == 1 || actual.cols == 1 {
        actual.rows * actual.cols - 1
    } else {
        (actual.rows - 1) * (actual.cols - 1)
    };
    if df == 0 {
        return err(ErrorValue::NA);
    }
    let stat: f64 = pairs.iter().map(|(a, e)| (a - e) * (a - e) / e).sum();
    finish(ctx, special::gamma_q(0.5 * df as f64, 0.5 * stat))
}

/// T.TEST and TTEST(array1, array2, tails, type). the paired test pairs the
/// arrays by position, keeping a pair only where both cells are numbers;
/// arrays of different sizes give `#N/A`.
pub(crate) fn t_test(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 4 {
        return err(ErrorValue::Value);
    }
    let tails = match nth_int(args, ctx, 2) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let kind = match nth_int(args, ctx, 3) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    if (tails != 1 && tails != 2) || !(1..=3).contains(&kind) {
        return err(ErrorValue::Num);
    }
    let (t, df) = match t_statistic(args, ctx, kind) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let rt = special::t_tail(t.abs(), df);
    finish(ctx, if tails == 1 { rt } else { 2.0 * rt })
}

fn t_statistic(args: &[Expr], ctx: &EvalContext<'_>, kind: i64) -> Result<(f64, f64), ErrorValue> {
    let (first, second) = grids(args, ctx)?;
    if kind == 1 {
        if first.rows * first.cols != second.rows * second.cols
            || first.cells.len() != second.cells.len()
        {
            return Err(ErrorValue::NA);
        }
        let scale = common_scale(&first, &second);
        let diffs: Vec<f64> = paired(&first, &second)
            .iter()
            .map(|(a, b)| a / scale - b / scale)
            .collect();
        if diffs.len() < 2 {
            return Err(ErrorValue::Div0);
        }
        let (mean, var) = sample_stats(&diffs);
        if var <= 0.0 {
            return Err(ErrorValue::Div0);
        }
        let n = diffs.len() as f64;
        Ok((mean / (var / n).sqrt(), n - 1.0))
    } else {
        let scale = common_scale(&first, &second);
        let first = scaled(first.numbers(), scale);
        let second = scaled(second.numbers(), scale);
        if first.len() < 2 || second.len() < 2 {
            return Err(ErrorValue::Div0);
        }
        let (m1, v1) = sample_stats(&first);
        let (m2, v2) = sample_stats(&second);
        let n1 = first.len() as f64;
        let n2 = second.len() as f64;
        if kind == 2 {
            let pooled = ((n1 - 1.0) * v1 + (n2 - 1.0) * v2) / (n1 + n2 - 2.0);
            let se = (pooled * (1.0 / n1 + 1.0 / n2)).sqrt();
            if se <= 0.0 {
                return Err(ErrorValue::Div0);
            }
            Ok(((m1 - m2) / se, n1 + n2 - 2.0))
        } else {
            let a = v1 / n1;
            let b = v2 / n2;
            let se = (a + b).sqrt();
            if se <= 0.0 {
                return Err(ErrorValue::Div0);
            }
            let df = (a + b) * (a + b) / (a * a / (n1 - 1.0) + b * b / (n2 - 1.0));
            Ok(((m1 - m2) / se, df))
        }
    }
}

/// F.TEST and FTEST(array1, array2): the two-tailed probability that the
/// variances of two samples are not significantly different.
pub(crate) fn f_test(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let (first, second) = match grids(args, ctx) {
        Ok((first, second)) => {
            let scale = common_scale(&first, &second);
            (
                scaled(first.numbers(), scale),
                scaled(second.numbers(), scale),
            )
        }
        Err(e) => return err(e),
    };
    if first.len() < 2 || second.len() < 2 {
        return err(ErrorValue::Div0);
    }
    let (_, v1) = sample_stats(&first);
    let (_, v2) = sample_stats(&second);
    if v1 <= 0.0 || v2 <= 0.0 {
        return err(ErrorValue::Div0);
    }
    let f = v1 / v2;
    let (cdf, tail) = special::f_tails(f, (first.len() - 1) as f64, (second.len() - 1) as f64);
    finish(ctx, 2.0 * cdf.min(tail))
}
