//! special functions shared by the distribution functions: log-gamma, the
//! regularized incomplete gamma and beta functions, the error function, the
//! normal, t and f distributions, and inversion of monotone tail functions.
//! algorithms are written from the mathematics: a lanczos approximation and
//! the stirling series for log-gamma, power series and continued fractions
//! (modified lentz) for the incomplete functions, temme's uniform asymptotic
//! expansion for the incomplete gamma function at large shape near its
//! mean, and safeguarded newton iteration for inverses. every tail is
//! computed directly rather than as one minus the other tail, so small
//! probabilities keep their relative precision.

use std::cell::Cell;
use std::f64::consts::{FRAC_1_SQRT_2, FRAC_2_SQRT_PI, PI};

const EPS: f64 = 1e-15;
const TINY: f64 = 1e-300;
const MAX_ITER: usize = 1000;
const BETA_MAX_ITER: usize = 20_000;
const SOLVE_MAX_ITER: usize = 500;
const LN_SQRT_2PI: f64 = 0.918_938_533_204_672_8;

thread_local! {
    static STEPS: Cell<u64> = const { Cell::new(0) };
}

/// count `n` iterations of the series, fractions and solvers, so the caller
/// can charge them to the recalculation budget.
fn tick(n: usize) {
    STEPS.with(|steps| steps.set(steps.get().saturating_add(n as u64)));
}

/// the iterations counted since the last call.
pub(crate) fn take_steps() -> u64 {
    STEPS.with(|steps| steps.replace(0))
}

/// lanczos coefficients for g = 7.
const LANCZOS: [f64; 9] = [
    0.999_999_999_999_809_9,
    676.520_368_121_885_1,
    -1_259.139_216_722_402_8,
    771.323_428_777_653_1,
    -176.615_029_162_140_6,
    12.507_343_278_686_905,
    -0.138_571_095_265_720_12,
    9.984_369_578_019_572e-6,
    1.505_632_735_149_311_6e-7,
];

/// taylor coefficients in eta of the first five coefficient functions of
/// temme's expansion of the incomplete gamma ratio.
const TEMME: [[f64; 14]; 5] = [
    [
        -0.3333333333333333,
        0.08333333333333333,
        -0.014814814814814815,
        0.0011574074074074073,
        0.0003527336860670194,
        -0.0001787551440329218,
        3.919263178522438e-05,
        -2.185448510679992e-06,
        -1.85406221071516e-06,
        8.296711340953087e-07,
        -1.7665952736826078e-07,
        6.707853543401498e-09,
        1.0261809784240309e-08,
        -4.382036018453353e-09,
    ],
    [
        -0.001851851851851852,
        -0.003472222222222222,
        0.0026455026455026454,
        -0.0009902263374485596,
        0.00020576131687242798,
        -4.018775720164609e-07,
        -1.8098550334489977e-05,
        7.64916091608111e-06,
        -1.6120900894563446e-06,
        4.647127802807434e-09,
        1.378633446915721e-07,
        -5.752545603517705e-08,
        1.1951628599778148e-08,
        -1.7543241719747647e-11,
    ],
    [
        0.004133597883597883,
        -0.0026813271604938273,
        0.0007716049382716049,
        2.0093878600823047e-06,
        -0.0001073665322636516,
        5.2923448829120125e-05,
        -1.2760635188618728e-05,
        3.423578734096138e-08,
        1.3721957309062934e-06,
        -6.298992138380055e-07,
        1.4280614206064242e-07,
        -2.0477098421990866e-10,
        -1.409252991086752e-08,
        6.228974084922022e-09,
    ],
    [
        0.0006494341563786008,
        0.00022947209362139917,
        -0.0004691894943952557,
        0.00026772063206283885,
        -7.561801671883977e-05,
        -2.396505113867297e-07,
        1.1082654115347302e-05,
        -5.6749528269915965e-06,
        1.4230900732435883e-06,
        -2.7861080291528143e-11,
        -1.6958404091930278e-07,
        8.099464905388083e-08,
        -1.9111168485973655e-08,
        2.3928620439808118e-12,
    ],
    [
        -0.0008618882909167117,
        0.0007840392217200666,
        -0.0002990724803031902,
        -1.4638452578843418e-06,
        6.641498215465122e-05,
        -3.968365047179435e-05,
        1.1375726970678419e-05,
        2.507497226237533e-10,
        -1.6954149536558305e-06,
        8.907507532205309e-07,
        -2.292934834000805e-07,
        2.956794137544049e-11,
        2.8865829742708783e-08,
        -1.4189739437803219e-08,
    ],
];
/// the shape from which temme's expansion replaces the series and the
/// continued fraction near the mean, and the relative distance from the
/// mean it covers.
const TEMME_MIN_SHAPE: f64 = 1000.0;
const TEMME_SPAN: f64 = 0.3;

/// the stirling series remainder ln Gamma(z) - ((z - 1/2) ln z - z + ln
/// sqrt(2 pi)), for z >= 10.
fn stirling_tail(z: f64) -> f64 {
    let z2 = 1.0 / (z * z);
    (1.0 / 12.0
        + z2 * (-1.0 / 360.0
            + z2 * (1.0 / 1260.0
                + z2 * (-1.0 / 1680.0
                    + z2 * (1.0 / 1188.0 + z2 * (-691.0 / 360_360.0 + z2 / 156.0))))))
        / z
}

fn lanczos(z: f64) -> f64 {
    let z = z - 1.0;
    let mut x = LANCZOS[0];
    for (i, c) in LANCZOS.iter().enumerate().skip(1) {
        x += c / (z + i as f64);
    }
    let t = z + 7.5;
    LN_SQRT_2PI + (z + 0.5) * t.ln() - t + x.ln()
}

/// ln Gamma(z) for z > 0: lanczos below 10, the stirling series above, and
/// the recurrence below 0.5 so subnormal arguments stay finite.
pub(crate) fn lngamma(z: f64) -> f64 {
    if z < 0.5 {
        lngamma(z + 1.0) - z.ln()
    } else if z >= 10.0 {
        (z - 0.5) * z.ln() - z + LN_SQRT_2PI + stirling_tail(z)
    } else {
        lanczos(z)
    }
}

/// ln Gamma(z + d) - ln Gamma(z), without the cancellation of two large
/// log-gammas when z is large.
fn lngamma_shift(z: f64, d: f64) -> f64 {
    if z >= 10.0 && z + d >= 10.0 {
        d * (z + d).ln() + (z - 0.5) * (d / z).ln_1p() - d + stirling_tail(z + d) - stirling_tail(z)
    } else {
        lngamma(z + d) - lngamma(z)
    }
}

/// Gamma(z) for any z that is not a non-positive integer (the caller checks
/// the poles). negative arguments go through the reflection formula.
pub(crate) fn gamma_fn(z: f64) -> f64 {
    if z < 0.5 && z != z.trunc() {
        PI / ((PI * z).sin() * gamma_fn(1.0 - z))
    } else if z < 0.5 {
        f64::NAN
    } else {
        lngamma(z).exp()
    }
}

/// ln B(a, b).
fn lnbeta(a: f64, b: f64) -> f64 {
    let (small, big) = if a < b { (a, b) } else { (b, a) };
    lngamma(small) - lngamma_shift(big, small)
}

/// ln(1 + u) - u, accurate for small u.
fn log1pmx(u: f64) -> f64 {
    if u.abs() > 0.5 {
        return u.ln_1p() - u;
    }
    let r = u / (2.0 + u);
    let r2 = r * r;
    let mut term = r * r2;
    let mut sum = 0.0;
    let mut k = 3.0;
    for _ in 0..MAX_ITER {
        let t = term / k;
        sum += t;
        if t.abs() <= 1e-17 * sum.abs() {
            break;
        }
        term *= r2;
        k += 2.0;
    }
    2.0 * sum - u * r
}

/// a ln(x / a) - (x - a), without cancellation near x = a.
fn scaled_deviation(a: f64, x: f64) -> f64 {
    let u = (x - a) / a;
    if u.abs() <= 0.5 {
        a * log1pmx(u)
    } else {
        a * (x / a).ln() - (x - a)
    }
}

/// x^a e^-x / Gamma(a), through the stirling form for large a so the
/// exponent does not cancel.
fn gamma_prefactor(a: f64, x: f64) -> f64 {
    if a < 10.0 {
        (a * x.ln() - x - lngamma(a)).exp()
    } else {
        (scaled_deviation(a, x) - stirling_tail(a)).exp() * (a / (2.0 * PI)).sqrt()
    }
}

fn gamma_p_series(a: f64, x: f64) -> f64 {
    let mut sum = 1.0 / a;
    let mut del = sum;
    for i in 1..MAX_ITER {
        del *= x / (a + i as f64);
        sum += del;
        if del.abs() < sum.abs() * EPS {
            tick(i);
            return sum * gamma_prefactor(a, x);
        }
    }
    tick(MAX_ITER);
    f64::NAN
}

fn gamma_q_cf(a: f64, x: f64) -> f64 {
    let first = x + 1.0 - a;
    let mut c = 1.0 / TINY;
    let mut d = 1.0 / if first.abs() < TINY { TINY } else { first };
    let mut h = d;
    for i in 1..MAX_ITER {
        let fi = i as f64;
        let an = -fi * (fi - a);
        let b = first + 2.0 * fi;
        d = an * d + b;
        if d.abs() < TINY {
            d = TINY;
        }
        c = b + an / c;
        if c.abs() < TINY {
            c = TINY;
        }
        d = 1.0 / d;
        let del = d * c;
        h *= del;
        if (del - 1.0).abs() < EPS {
            tick(i);
            return gamma_prefactor(a, x) * h;
        }
    }
    tick(MAX_ITER);
    f64::NAN
}

fn polynomial(coefficients: &[f64], x: f64) -> f64 {
    coefficients.iter().rev().fold(0.0, |acc, c| acc * x + c)
}

/// temme's uniform expansion Q(a, x) = erfc(eta sqrt(a / 2)) / 2 + R, for
/// large a with x near a; returns (P, Q).
fn gamma_temme(a: f64, x: f64) -> (f64, f64) {
    let u = (x - a) / a;
    let eta = (-2.0 * log1pmx(u)).sqrt().copysign(u);
    let y = eta * (0.5 * a).sqrt();
    let inv = 1.0 / a;
    let series = TEMME
        .iter()
        .rev()
        .fold(0.0, |acc, row| acc * inv + polynomial(row, eta));
    let r = (-0.5 * a * eta * eta).exp() / (2.0 * PI * a).sqrt() * series;
    if eta > 0.0 {
        let q = 0.5 * erfc_fn(y) + r;
        (1.0 - q, q)
    } else {
        let p = 0.5 * erfc_fn(-y) - r;
        (p, 1.0 - p)
    }
}

/// the regularized incomplete gamma ratios (P(a, x), Q(a, x)) for a > 0 and
/// x >= 0; the smaller one is computed directly.
pub(crate) fn gamma_pq(a: f64, x: f64) -> (f64, f64) {
    if x <= 0.0 {
        (0.0, 1.0)
    } else if x == f64::INFINITY {
        (1.0, 0.0)
    } else if a >= TEMME_MIN_SHAPE && (x - a).abs() <= TEMME_SPAN * a {
        gamma_temme(a, x)
    } else if x < a + 1.0 {
        let p = gamma_p_series(a, x);
        (p, 1.0 - p)
    } else {
        let q = gamma_q_cf(a, x);
        (1.0 - q, q)
    }
}

/// regularized lower incomplete gamma P(a, x).
pub(crate) fn gamma_p(a: f64, x: f64) -> f64 {
    gamma_pq(a, x).0
}

/// regularized upper incomplete gamma Q(a, x).
pub(crate) fn gamma_q(a: f64, x: f64) -> f64 {
    gamma_pq(a, x).1
}

/// the standard gamma density z^(a-1) e^-z / Gamma(a) for z > 0.
fn gamma_density(a: f64, z: f64) -> f64 {
    if a < 10.0 {
        ((a - 1.0) * z.ln() - z - lngamma(a)).exp()
    } else {
        gamma_prefactor(a, z) / z
    }
}

/// erf(x) through the incomplete gamma function: erf(x) = sign(x) P(1/2, x^2).
pub(crate) fn erf_fn(x: f64) -> f64 {
    if x.abs() < 1e-8 {
        return if x == 0.0 { 0.0 } else { x * FRAC_2_SQRT_PI };
    }
    let mag = gamma_p(0.5, x * x);
    if x < 0.0 { -mag } else { mag }
}

/// erfc(x), without cancellation in either tail.
pub(crate) fn erfc_fn(x: f64) -> f64 {
    if x < 0.0 {
        2.0 - erfc_fn(-x)
    } else {
        gamma_q(0.5, x * x)
    }
}

/// the standard normal density.
pub(crate) fn norm_pdf(z: f64) -> f64 {
    (-0.5 * z * z).exp() / (2.0 * PI).sqrt()
}

/// the standard normal cdf.
pub(crate) fn norm_cdf(z: f64) -> f64 {
    0.5 * erfc_fn(-z * FRAC_1_SQRT_2)
}

/// the lower-tail standard normal quantile for 0 < p <= 1/2, by acklam's
/// rational approximation polished with newton steps against [`norm_cdf`],
/// which is accurate in the lower tail.
pub(crate) fn norm_inv_lower(p: f64) -> f64 {
    const A: [f64; 6] = [
        -3.969_683_028_665_376e1,
        2.209_460_984_245_205e2,
        -2.759_285_984_469_534e2,
        1.383_577_518_672_69e2,
        -3.066_479_806_614_716e1,
        2.506_628_277_459_239e0,
    ];
    const B: [f64; 5] = [
        -5.447_609_879_822_406e1,
        1.615_858_368_580_409e2,
        -1.556_989_798_598_866e2,
        6.680_131_188_771_972e1,
        -1.328_068_155_288_572e1,
    ];
    const C: [f64; 6] = [
        -7.784_894_002_430_293e-3,
        -3.223_964_580_411_365e-1,
        -2.400_758_277_161_838e0,
        -2.549_732_539_343_734e0,
        4.374_664_141_464_968e0,
        2.938_163_982_698_783e0,
    ];
    const D: [f64; 4] = [
        7.784_695_709_041_462e-3,
        3.224_671_290_700_398e-1,
        2.445_134_137_142_996e0,
        3.754_408_661_907_416e0,
    ];
    let mut z = if p < 0.02425 {
        let q = (-2.0 * p.ln()).sqrt();
        (((((C[0] * q + C[1]) * q + C[2]) * q + C[3]) * q + C[4]) * q + C[5])
            / ((((D[0] * q + D[1]) * q + D[2]) * q + D[3]) * q + 1.0)
    } else {
        let q = p - 0.5;
        let r = q * q;
        (((((A[0] * r + A[1]) * r + A[2]) * r + A[3]) * r + A[4]) * r + A[5]) * q
            / (((((B[0] * r + B[1]) * r + B[2]) * r + B[3]) * r + B[4]) * r + 1.0)
    };
    for _ in 0..3 {
        let pdf = norm_pdf(z);
        if pdf <= 0.0 {
            break;
        }
        z -= (norm_cdf(z) - p) / pdf;
    }
    z
}

/// the inverse standard normal cdf for 0 < p < 1; above 1/2 it inverts the
/// exact complement 1 - p in the lower tail.
pub(crate) fn norm_inv_cdf(p: f64) -> f64 {
    if p > 0.5 {
        -norm_inv_lower(1.0 - p)
    } else {
        norm_inv_lower(p)
    }
}

fn midpoint(lo: f64, hi: f64) -> f64 {
    if lo <= 0.0 {
        if hi > 1e-300 { hi * 1e-3 } else { 0.5 * hi }
    } else if hi > 4.0 * lo {
        lo.sqrt() * hi.sqrt()
    } else {
        lo + 0.5 * (hi - lo)
    }
}

/// the root of an increasing `f` on `[lo, hi]` with `lo >= 0`, where `f(lo)
/// <= 0 <= f(hi)`: newton steps with the supplied slope, falling back to
/// bisection (geometric across wide brackets) and stopping on a step that is
/// small relative to the root, so small quantiles keep their digits.
fn solve(f: impl Fn(f64) -> f64, slope: impl Fn(f64) -> f64, mut lo: f64, mut hi: f64) -> f64 {
    let mut x = midpoint(lo, hi);
    for _ in 0..SOLVE_MAX_ITER {
        tick(1);
        let fx = f(x);
        if fx.is_nan() {
            return f64::NAN;
        }
        if fx == 0.0 {
            return x;
        }
        if fx < 0.0 {
            lo = x;
        } else {
            hi = x;
        }
        let s = slope(x);
        let newton = x - fx / s;
        let next = if s > 0.0 && s.is_finite() && newton > lo && newton < hi {
            newton
        } else {
            midpoint(lo, hi)
        };
        if (next - x).abs() <= EPS * next.abs() {
            return next;
        }
        x = next;
    }
    x
}

/// the smallest power-of-`factor` multiple of `start` (capped at the
/// largest finite value) where the increasing `f` is no longer negative, or
/// `None` when the root lies past every finite value.
fn upper_bracket(f: &impl Fn(f64) -> f64, start: f64, factor: f64) -> Option<f64> {
    let mut hi = start;
    loop {
        tick(1);
        if f(hi) >= 0.0 {
            return Some(hi);
        }
        if hi >= f64::MAX {
            return None;
        }
        hi = (hi * factor).min(f64::MAX);
    }
}

/// the standard gamma quantile: z with P(a, z) = p, where q = 1 - p is
/// passed exactly; the smaller tail is inverted.
pub(crate) fn gamma_quantile(a: f64, p: f64, q: f64) -> f64 {
    if p <= 0.0 {
        return 0.0;
    }
    let f = |z: f64| {
        let (lower, upper) = gamma_pq(a, z);
        if p <= q { lower - p } else { q - upper }
    };
    match upper_bracket(&f, a.max(1.0), 4.0) {
        Some(hi) => solve(f, |z| gamma_density(a, z), 0.0, hi),
        None => f64::INFINITY,
    }
}

/// the density of the gamma distribution with shape `a` and scale `beta`.
pub(crate) fn gamma_pdf(x: f64, a: f64, beta: f64) -> f64 {
    if x < 0.0 {
        0.0
    } else if x == 0.0 {
        if a < 1.0 {
            f64::INFINITY
        } else if a == 1.0 {
            1.0 / beta
        } else {
            0.0
        }
    } else {
        gamma_density(a, x / beta) / beta
    }
}

/// the cdf of the gamma distribution with shape `a` and scale `beta`.
pub(crate) fn gamma_cdf(x: f64, a: f64, beta: f64) -> f64 {
    gamma_p(a, x / beta)
}

fn beta_cf(a: f64, b: f64, x: f64) -> f64 {
    let qab = a + b;
    let qap = a + 1.0;
    let qam = a - 1.0;
    let mut c = 1.0;
    let mut d = 1.0 - qab * x / qap;
    if d.abs() < TINY {
        d = TINY;
    }
    d = 1.0 / d;
    let mut h = d;
    for m in 1..BETA_MAX_ITER {
        let fm = m as f64;
        let m2 = 2.0 * fm;
        let aa = fm * (b - fm) * x / ((qam + m2) * (a + m2));
        d = 1.0 + aa * d;
        if d.abs() < TINY {
            d = TINY;
        }
        c = 1.0 + aa / c;
        if c.abs() < TINY {
            c = TINY;
        }
        d = 1.0 / d;
        h *= d * c;
        let aa = -(a + fm) * (qab + fm) * x / ((a + m2) * (qap + m2));
        d = 1.0 + aa * d;
        if d.abs() < TINY {
            d = TINY;
        }
        c = 1.0 + aa / c;
        if c.abs() < TINY {
            c = TINY;
        }
        d = 1.0 / d;
        let del = d * c;
        h *= del;
        if (del - 1.0).abs() < EPS {
            tick(m);
            return h;
        }
    }
    tick(BETA_MAX_ITER);
    f64::NAN
}

/// (ln x, ln y) for y = 1 - x, each taken from the smaller of the two.
fn ln_pair(x: f64, y: f64) -> (f64, f64) {
    let lx = if x > 0.5 { (-y).ln_1p() } else { x.ln() };
    let ly = if y > 0.5 { (-x).ln_1p() } else { y.ln() };
    (lx, ly)
}

/// a ln(x s / a) - (x s - a) for s = a + b, the stirling-form exponent of
/// one side of the beta prefactor.
fn beta_side(a: f64, x: f64, lx: f64, s: f64) -> f64 {
    let u = (x * s - a) / a;
    if u.abs() <= 0.5 {
        a * log1pmx(u)
    } else {
        a * (lx + (s / a).ln()) - a * u
    }
}

/// x^a y^b / (B(a, b) d) with y = 1 - x supplied: the stirling form when
/// both parameters are large, a stirling difference for B when one is. the
/// divisor d enters the exponent there, so a tiny shape that makes B huge
/// does not pass through a subnormal.
fn beta_prefactor_over(a: f64, b: f64, x: f64, y: f64, d: f64) -> f64 {
    let (lx, ly) = ln_pair(x, y);
    if a >= 10.0 && b >= 10.0 {
        let s = a + b;
        let exponent =
            beta_side(a, x, lx, s) + beta_side(b, y, ly, s) - stirling_tail(a) - stirling_tail(b)
                + stirling_tail(s);
        exponent.exp() * (a * b / (2.0 * PI * s)).sqrt() / d
    } else {
        (a * lx + b * ly - lnbeta(a, b) - d.ln()).exp()
    }
}

/// x^a y^b / B(a, b).
fn beta_prefactor(a: f64, b: f64, x: f64, y: f64) -> f64 {
    beta_prefactor_over(a, b, x, y, 1.0)
}

/// the regularized incomplete beta ratios (I_x(a, b), 1 - I_x(a, b)) with
/// y = 1 - x supplied; the smaller one is computed directly.
pub(crate) fn beta_pq(a: f64, b: f64, x: f64, y: f64) -> (f64, f64) {
    if x <= 0.0 {
        return (0.0, 1.0);
    }
    if y <= 0.0 {
        return (1.0, 0.0);
    }
    if x < (a + 1.0) / (a + b + 2.0) {
        let w = beta_prefactor_over(a, b, x, y, a) * beta_cf(a, b, x);
        (w, 1.0 - w)
    } else {
        let w = beta_prefactor_over(a, b, x, y, b) * beta_cf(b, a, y);
        (1.0 - w, w)
    }
}

/// (x - lo, hi - x) / (hi - lo), halving first when the span overflows.
fn unit_position(x: f64, lo: f64, hi: f64) -> (f64, f64) {
    let span = hi - lo;
    if span.is_finite() {
        ((x - lo) / span, (hi - x) / span)
    } else {
        let half = 0.5 * hi - 0.5 * lo;
        ((0.5 * x - 0.5 * lo) / half, (0.5 * hi - 0.5 * x) / half)
    }
}

/// the density of the beta distribution on `[lo, hi]`.
pub(crate) fn beta_pdf(x: f64, a: f64, b: f64, lo: f64, hi: f64) -> f64 {
    let (t, y) = unit_position(x, lo, hi);
    if !(0.0..=1.0).contains(&t) {
        return 0.0;
    }
    let edge = |shape: f64, other: f64| {
        if shape < 1.0 {
            f64::INFINITY
        } else if shape == 1.0 {
            other / (hi - lo)
        } else {
            0.0
        }
    };
    if t == 0.0 {
        return edge(a, b);
    }
    if t == 1.0 {
        return edge(b, a);
    }
    beta_prefactor(a, b, t, y) / (t * y * (hi - lo))
}

/// the cdf of the beta distribution on `[lo, hi]`.
pub(crate) fn beta_cdf(x: f64, a: f64, b: f64, lo: f64, hi: f64) -> f64 {
    let (t, y) = unit_position(x, lo, hi);
    beta_pq(a, b, t, y).0
}

/// the standard beta quantile on `[0, 1]`: t with I_t(a, b) = p, where
/// q = 1 - p is passed exactly; the smaller tail is inverted.
pub(crate) fn beta_quantile(a: f64, b: f64, p: f64, q: f64) -> f64 {
    if p <= 0.0 {
        return 0.0;
    }
    if q <= 0.0 {
        return 1.0;
    }
    let f = |t: f64| {
        let (lower, upper) = beta_pq(a, b, t, 1.0 - t);
        if p <= q { lower - p } else { q - upper }
    };
    let slope = |t: f64| beta_prefactor(a, b, t, 1.0 - t) / (t * (1.0 - t));
    solve(f, slope, 0.0, 1.0)
}

/// (df / (df + t^2), t^2 / (df + t^2)) without forming t^2 + df.
fn t_parts(t: f64, df: f64) -> (f64, f64) {
    let q = t / df.sqrt();
    let r = q * q;
    if r.is_infinite() {
        return (0.0, 1.0);
    }
    (1.0 / (1.0 + r), r / (1.0 + r))
}

/// the right tail P(T > t) of student's t distribution, for t >= 0, as
/// I_w(df / 2, 1 / 2) / 2.
pub(crate) fn t_tail(t: f64, df: f64) -> f64 {
    if t == 0.0 {
        return 0.5;
    }
    let q = t / df.sqrt();
    if q > 1e16 {
        let a = 0.5 * df;
        return (-df * q.ln() - a.ln() - lnbeta(a, 0.5)).exp() * 0.5;
    }
    let (w, y) = t_parts(t, df);
    0.5 * beta_pq(0.5 * df, 0.5, w, y).0
}

/// the right tail P(T > t) for any t.
pub(crate) fn t_right(t: f64, df: f64) -> f64 {
    if t < 0.0 {
        1.0 - t_tail(-t, df)
    } else {
        t_tail(t, df)
    }
}

/// the cdf of student's t distribution.
pub(crate) fn t_cdf(t: f64, df: f64) -> f64 {
    if t < 0.0 {
        t_tail(-t, df)
    } else {
        1.0 - t_tail(t, df)
    }
}

/// the density of student's t distribution with `df` degrees of freedom.
pub(crate) fn t_pdf(t: f64, df: f64) -> f64 {
    let q = t / df.sqrt();
    (lngamma_shift(0.5 * df, 0.5) - 0.5 * (df * PI).ln() - 0.5 * (df + 1.0) * (q * q).ln_1p()).exp()
}

/// the t value whose right tail is q, for 0 < q <= 1/2.
pub(crate) fn t_upper_quantile(q: f64, df: f64) -> f64 {
    if q >= 0.5 {
        return 0.0;
    }
    let f = |t: f64| q - t_tail(t, df);
    match upper_bracket(&f, 1.0, 16.0) {
        Some(hi) => solve(f, |t| t_pdf(t, df), 0.0, hi),
        None => f64::INFINITY,
    }
}

/// (d1 x / (d1 x + d2), d2 / (d1 x + d2)) for x > 0, formed from the ratio
/// d2 / (d1 x) so neither overflows nor rounds to 1 too early.
fn f_parts(x: f64, d1: f64, d2: f64) -> (f64, f64) {
    let r = d2 / d1 / x;
    if r.is_infinite() {
        return (0.0, 1.0);
    }
    (1.0 / (1.0 + r), r / (1.0 + r))
}

/// the density of the f distribution.
pub(crate) fn f_pdf(x: f64, d1: f64, d2: f64) -> f64 {
    if x < 0.0 {
        0.0
    } else if x == 0.0 {
        if d1 < 2.0 {
            f64::INFINITY
        } else if d1 == 2.0 {
            1.0
        } else {
            0.0
        }
    } else {
        let (u, v) = f_parts(x, d1, d2);
        beta_prefactor(0.5 * d1, 0.5 * d2, u, v) / x
    }
}

/// the (cdf, right tail) pair of the f distribution.
pub(crate) fn f_tails(x: f64, d1: f64, d2: f64) -> (f64, f64) {
    if x <= 0.0 {
        return (0.0, 1.0);
    }
    if x == 1.0 && d1 == d2 {
        return (0.5, 0.5);
    }
    let (u, v) = f_parts(x, d1, d2);
    beta_pq(0.5 * d1, 0.5 * d2, u, v)
}

/// the f quantile: x with cdf p, where q = 1 - p is passed exactly; the
/// smaller tail is inverted. with equal degrees of freedom F and 1/F share a
/// distribution, so the median is exactly 1, as Excel gives at 10^9 and
/// 10^10 degrees of freedom.
pub(crate) fn f_quantile(p: f64, q: f64, d1: f64, d2: f64) -> f64 {
    if p <= 0.0 {
        return 0.0;
    }
    if p == q && d1 == d2 {
        return 1.0;
    }
    let f = |x: f64| {
        let (lower, upper) = f_tails(x, d1, d2);
        if p <= q { lower - p } else { q - upper }
    };
    match upper_bracket(&f, 1.0, 16.0) {
        Some(hi) => solve(f, |x| f_pdf(x, d1, d2), 0.0, hi),
        None => f64::INFINITY,
    }
}

/// the chi-square cdf with `df` degrees of freedom.
pub(crate) fn chisq_cdf(x: f64, df: f64) -> f64 {
    gamma_p(0.5 * df, 0.5 * x)
}

/// the chi-square right tail.
pub(crate) fn chisq_tail(x: f64, df: f64) -> f64 {
    gamma_q(0.5 * df, 0.5 * x)
}

/// the chi-square density.
pub(crate) fn chisq_pdf(x: f64, df: f64) -> f64 {
    gamma_pdf(x, 0.5 * df, 2.0)
}
