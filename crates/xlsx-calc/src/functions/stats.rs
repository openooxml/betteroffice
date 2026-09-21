//! statistical functions. numeric aggregations ignore text/bool/blank inside
//! references, coerce literal arguments, propagate errors.

use std::collections::HashMap;

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{Area, EvalContext, as_area, err, evaluate, num};
use crate::parser::Expr;

use super::criteria::{self, Criterion};
use super::{collect_numbers, finite, nth_int, nth_number};

pub(crate) fn average(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match collect_numbers(args, ctx) {
        Ok(nums) if nums.is_empty() => err(ErrorValue::Div0),
        Ok(nums) => num(nums.iter().sum::<f64>() / nums.len() as f64),
        Err(e) => err(e),
    }
}

pub(crate) fn min(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match collect_numbers(args, ctx) {
        Ok(nums) if nums.is_empty() => num(0.0),
        Ok(nums) => num(nums.iter().copied().fold(f64::INFINITY, f64::min)),
        Err(e) => err(e),
    }
}

pub(crate) fn max(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match collect_numbers(args, ctx) {
        Ok(nums) if nums.is_empty() => num(0.0),
        Ok(nums) => num(nums.iter().copied().fold(f64::NEG_INFINITY, f64::max)),
        Err(e) => err(e),
    }
}

/// paired numeric samples from two areas, dropping positions where either side
/// is not a number. excel requires matching counts.
fn pairs(args: &[Expr], ctx: &EvalContext<'_>) -> Result<(Vec<f64>, Vec<f64>), ErrorValue> {
    if args.len() != 2 {
        return Err(ErrorValue::Value);
    }
    let ys = positioned(&args[0], ctx)?;
    let xs = positioned(&args[1], ctx)?;
    if xs.len() != ys.len() {
        return Err(ErrorValue::NA);
    }
    // a coordinate counts only where both sides are numeric, so dropping one
    // side's blank cannot slide every later pair onto the wrong partner
    let (ys, xs): (Vec<f64>, Vec<f64>) = ys
        .into_iter()
        .zip(xs)
        .filter_map(|(y, x)| Some((y?, x?)))
        .unzip();
    if xs.is_empty() {
        return Err(ErrorValue::NA);
    }
    Ok((ys, xs))
}

/// every cell of an argument in order, `None` where it is not a number, so two
/// ranges stay aligned by position.
fn positioned(arg: &Expr, ctx: &EvalContext<'_>) -> Result<Vec<Option<f64>>, ErrorValue> {
    match as_area(arg, ctx) {
        Some(area) => area
            .values_ref(ctx)?
            .into_iter()
            .map(|value| match value.as_ref() {
                CellValue::Number { value } => Ok(Some(*value)),
                CellValue::Error { value } => Err(*value),
                _ => Ok(None),
            })
            .collect(),
        None => match evaluate(arg, ctx) {
            CellValue::Error { value } => Err(value),
            CellValue::Number { value } => Ok(vec![Some(value)]),
            _ => Ok(vec![None]),
        },
    }
}

/// sums a linear fit needs: n, mean x, mean y, Sxx, Syy, Sxy.
fn moments(ys: &[f64], xs: &[f64]) -> (f64, f64, f64, f64, f64, f64) {
    let n = xs.len() as f64;
    let mx = xs.iter().sum::<f64>() / n;
    let my = ys.iter().sum::<f64>() / n;
    let (mut sxx, mut syy, mut sxy) = (0.0, 0.0, 0.0);
    for (x, y) in xs.iter().zip(ys) {
        sxx += (x - mx) * (x - mx);
        syy += (y - my) * (y - my);
        sxy += (x - mx) * (y - my);
    }
    (n, mx, my, sxx, syy, sxy)
}

/// CORREL(y, x): Pearson's r. zero variance on either side is `#DIV/0!`.
pub(crate) fn correl(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let (ys, xs) = match pairs(args, ctx) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let (_, _, _, sxx, syy, sxy) = moments(&ys, &xs);
    if sxx == 0.0 || syy == 0.0 {
        return err(ErrorValue::Div0);
    }
    finite(sxy / (sxx * syy).sqrt())
}

/// SLOPE(y, x) of the least-squares line.
pub(crate) fn slope(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let (ys, xs) = match pairs(args, ctx) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let (_, _, _, sxx, _, sxy) = moments(&ys, &xs);
    if sxx == 0.0 {
        return err(ErrorValue::Div0);
    }
    finite(sxy / sxx)
}

/// INTERCEPT(y, x) of the least-squares line.
pub(crate) fn intercept(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let (ys, xs) = match pairs(args, ctx) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let (_, mx, my, sxx, _, sxy) = moments(&ys, &xs);
    if sxx == 0.0 {
        return err(ErrorValue::Div0);
    }
    finite(my - (sxy / sxx) * mx)
}

/// COVARIANCE.P / COVARIANCE.S over paired samples.
pub(crate) fn covariance_p(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    covariance(args, ctx, 0.0)
}

pub(crate) fn covariance_s(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    covariance(args, ctx, 1.0)
}

fn covariance(args: &[Expr], ctx: &EvalContext<'_>, lost: f64) -> CellValue {
    let (ys, xs) = match pairs(args, ctx) {
        Ok(v) => v,
        Err(e) => return err(e),
    };
    let (n, _, _, _, _, sxy) = moments(&ys, &xs);
    if n - lost <= 0.0 {
        return err(ErrorValue::Div0);
    }
    finite(sxy / (n - lost))
}

/// PERCENTILE.INC(array, k): linear interpolation between order statistics,
/// which is what QUARTILE.INC divides into quarters.
fn percentile_inc(sorted: &[f64], k: f64) -> Result<f64, ErrorValue> {
    if sorted.is_empty() || !(0.0..=1.0).contains(&k) {
        return Err(ErrorValue::Num);
    }
    let position = k * (sorted.len() - 1) as f64;
    let lower = position.floor() as usize;
    let upper = position.ceil() as usize;
    Ok(sorted[lower] + (position - lower as f64) * (sorted[upper] - sorted[lower]))
}

pub(crate) fn percentile(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let mut nums = match collect_numbers(&args[..1], ctx) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    nums.sort_by(f64::total_cmp);
    match nth_number(args, ctx, 1).and_then(|k| percentile_inc(&nums, k)) {
        Ok(v) => finite(v),
        Err(e) => err(e),
    }
}

pub(crate) fn quartile(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let mut nums = match collect_numbers(&args[..1], ctx) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    nums.sort_by(f64::total_cmp);
    let quart = match nth_number(args, ctx, 1) {
        Ok(q) => q.trunc(),
        Err(e) => return err(e),
    };
    if !(0.0..=4.0).contains(&quart) {
        return err(ErrorValue::Num);
    }
    match percentile_inc(&nums, quart / 4.0) {
        Ok(v) => finite(v),
        Err(e) => err(e),
    }
}

pub(crate) fn median(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match collect_numbers(args, ctx) {
        Ok(nums) if nums.is_empty() => err(ErrorValue::Num),
        Ok(mut nums) => {
            nums.sort_by(f64::total_cmp);
            let n = nums.len();
            if n % 2 == 1 {
                num(nums[n / 2])
            } else {
                num((nums[n / 2 - 1] + nums[n / 2]) / 2.0)
            }
        }
        Err(e) => err(e),
    }
}

/// MODE.SNGL: the most frequent value; the earliest-appearing one wins ties.
/// no repeats -> #N/A.
pub(crate) fn mode(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let nums = match collect_numbers(args, ctx) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let mut counts: HashMap<u64, (f64, usize, usize)> = HashMap::new();
    for (i, &x) in nums.iter().enumerate() {
        if !x.is_finite() {
            return err(ErrorValue::Num);
        }
        let key = if x == 0.0 { 0 } else { x.to_bits() };
        counts
            .entry(key)
            .and_modify(|(_, count, _)| *count += 1)
            .or_insert((x, 1, i));
    }
    let mut best: Option<(f64, usize, usize)> = None;
    for (_, (value, count, first)) in counts {
        if count < 2 {
            continue;
        }
        match best {
            Some((_, best_count, best_first))
                if best_count > count || (best_count == count && best_first <= first) => {}
            _ => best = Some((value, count, first)),
        }
    }
    match best {
        Some((v, _, _)) => num(v),
        None => err(ErrorValue::NA),
    }
}

pub(crate) fn stdev_s(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    variance(args, ctx, true).map(f64::sqrt).into_cell()
}

pub(crate) fn stdev_p(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    variance(args, ctx, false).map(f64::sqrt).into_cell()
}

pub(crate) fn var_s(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    variance(args, ctx, true).into_cell()
}

pub(crate) fn var_p(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    variance(args, ctx, false).into_cell()
}

/// LARGE(array, k): the kth largest value (k = 1 is the maximum).
pub(crate) fn large(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    nth_order(args, ctx, true)
}

/// SMALL(array, k): the kth smallest value.
pub(crate) fn small(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    nth_order(args, ctx, false)
}

/// RANK(number, ref, [order]): position of `number` among the numbers in `ref`;
/// order omitted/0 ranks descending, nonzero ascending, ties share the best rank.
pub(crate) fn rank(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 && args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let target = match nth_number(args, ctx, 0) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let area = match as_area(&args[1], ctx) {
        Some(a) => a,
        None => return err(ErrorValue::Value),
    };
    let ascending = if args.len() == 3 {
        match nth_number(args, ctx, 2) {
            Ok(o) => o != 0.0,
            Err(e) => return err(e),
        }
    } else {
        false
    };
    let values = match area.values_ref(ctx) {
        Ok(values) => values,
        Err(error) => return err(error),
    };
    let nums: Vec<f64> = values
        .iter()
        .filter_map(|v| match v.as_ref() {
            CellValue::Number { value } => Some(*value),
            _ => None,
        })
        .collect();
    if !nums.contains(&target) {
        return err(ErrorValue::NA);
    }
    let better = nums
        .iter()
        .filter(|&&x| if ascending { x < target } else { x > target })
        .count();
    num(better as f64 + 1.0)
}

/// COUNT: numeric values only; errors and non-numerics are ignored (excel).
pub(crate) fn count(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let mut count = 0i64;
    for arg in args {
        match as_area(arg, ctx) {
            Some(area) => {
                let values = match area.values_ref(ctx) {
                    Ok(values) => values,
                    Err(error) => return err(error),
                };
                count += values
                    .iter()
                    .filter(|v| matches!(v.as_ref(), CellValue::Number { .. }))
                    .count() as i64;
            }
            None => match evaluate(arg, ctx) {
                CellValue::Number { .. } | CellValue::Bool { .. } => count += 1,
                CellValue::Text { value } if crate::eval::parse_num(&value).is_some() => count += 1,
                _ => {}
            },
        }
    }
    num(count as f64)
}

/// COUNTA: every non-empty value (text and errors included).
pub(crate) fn counta(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let mut count = 0i64;
    for arg in args {
        match as_area(arg, ctx) {
            Some(area) => {
                let values = match area.values_ref(ctx) {
                    Ok(values) => values,
                    Err(error) => return err(error),
                };
                count += values
                    .iter()
                    .filter(|v| !matches!(v.as_ref(), CellValue::Empty))
                    .count() as i64;
            }
            None => {
                if !matches!(evaluate(arg, ctx), CellValue::Empty) {
                    count += 1;
                }
            }
        }
    }
    num(count as f64)
}

/// COUNTBLANK(range): empty cells and empty strings.
pub(crate) fn countblank(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    let area = match as_area(&args[0], ctx) {
        Some(a) => a,
        None => return err(ErrorValue::Value),
    };
    let values = match area.values_ref(ctx) {
        Ok(values) => values,
        Err(error) => return err(error),
    };
    let n = values
        .iter()
        .filter(|v| {
            matches!(v.as_ref(), CellValue::Empty)
                || matches!(v.as_ref(), CellValue::Text { value } if value.is_empty())
        })
        .count();
    num(n as f64)
}

pub(crate) fn countif(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let area = match as_area(&args[0], ctx) {
        Some(a) => a,
        None => return err(ErrorValue::Value),
    };
    let criterion = criteria::criterion_from_arg(&args[1], ctx);
    let pairs = [(area, criterion)];
    match criteria::matching_indices(&pairs, ctx) {
        Ok(indices) => num(indices.len() as f64),
        Err(error) => err(error),
    }
}

pub(crate) fn countifs(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match criteria::collect_pairs(args, ctx) {
        Some(pairs) => match criteria::matching_indices(&pairs, ctx) {
            Ok(indices) => num(indices.len() as f64),
            Err(error) => err(error),
        },
        None => err(ErrorValue::Value),
    }
}

/// AVERAGEIF(range, criteria, [average_range]).
pub(crate) fn averageif(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 && args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let crit_area = match as_area(&args[0], ctx) {
        Some(a) => a,
        None => return err(ErrorValue::Value),
    };
    let value_spec = if args.len() == 3 { &args[2] } else { &args[0] };
    let value_area = match as_area(value_spec, ctx) {
        Some(a) => a,
        None => return err(ErrorValue::Value),
    };
    let criterion = criteria::criterion_from_arg(&args[1], ctx);
    average_of(&[(crit_area, criterion)], &value_area, ctx)
}

/// AVERAGEIFS(average_range, crit_range1, crit1, ...).
pub(crate) fn averageifs(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 3 {
        return err(ErrorValue::Value);
    }
    let value_area = match as_area(&args[0], ctx) {
        Some(a) => a,
        None => return err(ErrorValue::Value),
    };
    match criteria::collect_pairs(&args[1..], ctx) {
        Some(pairs) if pairs[0].0.rows == value_area.rows && pairs[0].0.cols == value_area.cols => {
            average_of(&pairs, &value_area, ctx)
        }
        _ => err(ErrorValue::Value),
    }
}

/// MAXIFS/MINIFS(values, range1, criteria1, ...): the extreme of the values
/// whose row satisfies every criterion; no match yields 0, as excel does.
pub(crate) fn maxifs(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    extreme_ifs(args, ctx, true)
}

pub(crate) fn minifs(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    extreme_ifs(args, ctx, false)
}

fn extreme_ifs(args: &[Expr], ctx: &EvalContext<'_>, largest: bool) -> CellValue {
    if args.len() < 3 {
        return err(ErrorValue::Value);
    }
    let value_area = match as_area(&args[0], ctx) {
        Some(area) => area,
        None => return err(ErrorValue::Value),
    };
    let pairs = match criteria::collect_pairs(&args[1..], ctx) {
        Some(pairs) if pairs[0].0.rows == value_area.rows && pairs[0].0.cols == value_area.cols => {
            pairs
        }
        _ => return err(ErrorValue::Value),
    };
    let nums = match matching_numbers(&pairs, &value_area, ctx) {
        Ok(nums) => nums,
        Err(error) => return err(error),
    };
    let picked = nums.into_iter().fold(None, |best: Option<f64>, value| {
        Some(match best {
            Some(best) if largest => best.max(value),
            Some(best) => best.min(value),
            None => value,
        })
    });
    num(picked.unwrap_or(0.0))
}

fn average_of(pairs: &[(Area, Criterion)], value_area: &Area, ctx: &EvalContext<'_>) -> CellValue {
    let nums = match matching_numbers(pairs, value_area, ctx) {
        Ok(nums) => nums,
        Err(error) => return err(error),
    };
    if nums.is_empty() {
        err(ErrorValue::Div0)
    } else {
        num(nums.iter().sum::<f64>() / nums.len() as f64)
    }
}

fn matching_numbers(
    pairs: &[(Area, Criterion)],
    value_area: &Area,
    ctx: &EvalContext<'_>,
) -> Result<Vec<f64>, ErrorValue> {
    let cols = pairs[0].0.cols;
    let mut nums = Vec::new();
    for i in criteria::matching_indices(pairs, ctx)? {
        let (r, c) = (i / cols, i % cols);
        if let CellValue::Number { value } = *value_area.get_ref(ctx, r, c)? {
            nums.push(value);
        }
    }
    Ok(nums)
}

/// sample (n-1) or population (n) variance; too few values -> #DIV/0!.
fn variance(args: &[Expr], ctx: &EvalContext<'_>, sample: bool) -> Result<f64, ErrorValue> {
    let nums = collect_numbers(args, ctx)?;
    let n = nums.len();
    let denom_ok = if sample { n >= 2 } else { n >= 1 };
    if !denom_ok {
        return Err(ErrorValue::Div0);
    }
    let mean = nums.iter().sum::<f64>() / n as f64;
    let ss: f64 = nums.iter().map(|x| (x - mean).powi(2)).sum();
    let denom = if sample { n as f64 - 1.0 } else { n as f64 };
    Ok(ss / denom)
}

fn nth_order(args: &[Expr], ctx: &EvalContext<'_>, largest: bool) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let mut nums = match collect_numbers(&args[..1], ctx) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let k = match nth_int(args, ctx, 1) {
        Ok(k) => k,
        Err(e) => return err(e),
    };
    if k < 1 || k as usize > nums.len() {
        return err(ErrorValue::Num);
    }
    nums.sort_by(f64::total_cmp);
    let idx = if largest {
        nums.len() - k as usize
    } else {
        k as usize - 1
    };
    num(nums[idx])
}

/// tiny helper so the variance/stdev entry points read as one expression.
trait IntoCell {
    fn into_cell(self) -> CellValue;
}

impl IntoCell for Result<f64, ErrorValue> {
    fn into_cell(self) -> CellValue {
        match self {
            Ok(v) => num(v),
            Err(e) => err(e),
        }
    }
}
