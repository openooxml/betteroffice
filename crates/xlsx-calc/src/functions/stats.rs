//! statistical functions. numeric aggregations ignore text/bool/blank inside
//! references, coerce literal arguments, propagate errors.

use std::collections::HashMap;

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{Area, EvalContext, as_area, bound_area, err, evaluate, num};
use crate::parser::Expr;

use super::criteria::{self, Criterion};
use super::{
    collect_numbers, collect_numbers_anytype, collect_numbers_deep, finite, nth_int, nth_number,
};

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
        None => match crate::array::evaluate_array(arg, ctx) {
            crate::array::Value::Array(array) => (0..array.rows())
                .flat_map(|row| (0..array.cols()).map(move |col| (row, col)))
                .map(|(row, col)| match array.at(row, col) {
                    CellValue::Number { value } => Ok(Some(value)),
                    CellValue::Error { value } => Err(value),
                    _ => Ok(None),
                })
                .collect(),
            value => match value.into_scalar() {
                CellValue::Error { value } => Err(value),
                CellValue::Number { value } => Ok(vec![Some(value)]),
                _ => Ok(vec![None]),
            },
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

/// SUBTOTAL(code, ...): one of eleven aggregates, chosen by code. 1-11 and
/// 101-111 select the same function; the 100 series also skips manually
/// hidden rows, which we do not model, so both behave alike here.
pub(crate) fn subtotal(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 2 {
        return err(ErrorValue::Value);
    }
    let code = match nth_int(args, ctx, 0) {
        Ok(code) => code,
        Err(e) => return err(e),
    };
    let rest = &args[1..];
    match code {
        1 | 101 => average(rest, ctx),
        2 | 102 => count(rest, ctx),
        3 | 103 => counta(rest, ctx),
        4 | 104 => max(rest, ctx),
        5 | 105 => min(rest, ctx),
        6 | 106 => super::math::product(rest, ctx),
        7 | 107 => stdev_s(rest, ctx),
        8 | 108 => stdev_p(rest, ctx),
        9 | 109 => super::math::sum(rest, ctx),
        10 | 110 => var_s(rest, ctx),
        11 | 111 => var_p(rest, ctx),
        _ => err(ErrorValue::Value),
    }
}

/// AGGREGATE(code, options, ...): the SUBTOTAL aggregates plus LARGE, SMALL,
/// PERCENTILE and QUARTILE, with an options flag that says what to ignore.
/// Options 2,3,6,7 ignore errors, which is the whole point of the function in
/// most sheets; nested subtotals and hidden rows are not modelled.
pub(crate) fn aggregate(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 3 {
        return err(ErrorValue::Value);
    }
    let code = match nth_int(args, ctx, 0) {
        Ok(code) => code,
        Err(e) => return err(e),
    };
    let options = match nth_int(args, ctx, 1) {
        Ok(options) => options,
        Err(e) => return err(e),
    };
    if !(1..=19).contains(&code) || !(0..=7).contains(&options) {
        return err(ErrorValue::Value);
    }
    let rest = &args[2..];
    let skip_errors = matches!(options, 2 | 3 | 6 | 7);
    let value = match code {
        1 => average(rest, ctx),
        2 => count(rest, ctx),
        3 => counta(rest, ctx),
        4 => max(rest, ctx),
        5 => min(rest, ctx),
        6 => super::math::product(rest, ctx),
        7 => stdev_s(rest, ctx),
        8 => stdev_p(rest, ctx),
        9 => super::math::sum(rest, ctx),
        10 => var_s(rest, ctx),
        11 => var_p(rest, ctx),
        12 => median(rest, ctx),
        13 => mode(rest, ctx),
        14 => large(rest, ctx),
        15 => small(rest, ctx),
        16 => percentile(rest, ctx),
        17 => quartile(rest, ctx),
        _ => return err(ErrorValue::Value),
    };
    match (&value, skip_errors) {
        (CellValue::Error { .. }, true) => aggregate_ignoring_errors(code, rest, ctx),
        _ => value,
    }
}

/// the numeric cells of an argument list, skipping errors, as a literal list
/// the scalar aggregates can be re-run over.
fn numeric_literals(args: &[Expr], ctx: &EvalContext<'_>) -> Vec<f64> {
    let mut nums = Vec::new();
    for arg in args {
        match as_area(arg, ctx) {
            Some(area) => {
                if let Ok(values) = area.values_ref(ctx) {
                    for value in values {
                        if let CellValue::Number { value } = value.as_ref() {
                            nums.push(*value);
                        }
                    }
                }
            }
            // `ROW(range)/(range=key)` is the nth-match idiom: a computed
            // block whose misses are #DIV/0!, which the skip-errors options
            // exist to drop
            None => match crate::array::evaluate_array(arg, ctx) {
                crate::array::Value::Array(array) => {
                    for row in 0..array.rows() {
                        for col in 0..array.cols() {
                            if let CellValue::Number { value } = array.at(row, col) {
                                nums.push(value);
                            }
                        }
                    }
                }
                value => {
                    if let CellValue::Number { value } = value.into_scalar() {
                        nums.push(value);
                    }
                }
            },
        }
    }
    nums
}

/// re-run an aggregate over only the cells that are not errors. every code
/// `AGGREGATE` accepts is handled here, so an error in the input cannot turn a
/// supported code into `#VALUE!`.
fn aggregate_ignoring_errors(code: i64, args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let take = match code {
        14..=17 => args.len().saturating_sub(1),
        _ => args.len(),
    };
    let mut nums = numeric_literals(&args[..take], ctx);
    if nums.is_empty() {
        return err(ErrorValue::Div0);
    }
    let n = nums.len() as f64;
    let mean = nums.iter().sum::<f64>() / n;
    let squares = nums.iter().map(|v| (v - mean).powi(2)).sum::<f64>();
    let variance = |lost: f64| -> CellValue {
        if n - lost <= 0.0 {
            return err(ErrorValue::Div0);
        }
        finite(squares / (n - lost))
    };
    let deviation = |lost: f64| -> CellValue {
        if n - lost <= 0.0 {
            return err(ErrorValue::Div0);
        }
        finite((squares / (n - lost)).sqrt())
    };
    // 14-17 take a rank or quantile as their final argument
    let k = match code {
        14..=17 => match nth_number(args, ctx, args.len() - 1) {
            Ok(k) => k,
            Err(e) => return err(e),
        },
        _ => 0.0,
    };
    nums.sort_by(f64::total_cmp);
    match code {
        1 => num(mean),
        2 | 3 => num(n),
        4 => num(nums[nums.len() - 1]),
        5 => num(nums[0]),
        6 => num(nums.iter().product()),
        7 => deviation(1.0),
        8 => deviation(0.0),
        9 => num(nums.iter().sum()),
        10 => variance(1.0),
        11 => variance(0.0),
        12 => num(if nums.len() % 2 == 1 {
            nums[nums.len() / 2]
        } else {
            (nums[nums.len() / 2 - 1] + nums[nums.len() / 2]) / 2.0
        }),
        13 => mode_of(&nums),
        14 => nth_from_end(&nums, k, true),
        15 => nth_from_end(&nums, k, false),
        16 | 17 => match percentile_inc(&nums, if code == 16 { k } else { k.trunc() / 4.0 }) {
            Ok(v) => finite(v),
            Err(e) => err(e),
        },
        _ => err(ErrorValue::Value),
    }
}

/// LARGE/SMALL over an already-sorted list.
fn nth_from_end(sorted: &[f64], k: f64, largest: bool) -> CellValue {
    let k = k.trunc();
    if k < 1.0 || k > sorted.len() as f64 {
        return err(ErrorValue::Num);
    }
    let index = k as usize - 1;
    num(if largest {
        sorted[sorted.len() - 1 - index]
    } else {
        sorted[index]
    })
}

/// the most frequent value, or `#N/A` when every value occurs once.
fn mode_of(nums: &[f64]) -> CellValue {
    let mut best: Option<(f64, usize)> = None;
    for value in nums {
        let count = nums.iter().filter(|other| *other == value).count();
        if count > 1 && best.is_none_or(|(_, seen)| count > seen) {
            best = Some((*value, count));
        }
    }
    match best {
        Some((value, _)) => num(value),
        None => err(ErrorValue::NA),
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
        Some(a) => bound_area(a, ctx),
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
        Some(a) => bound_area(a, ctx),
        None => return err(ErrorValue::Value),
    };
    let value_spec = if args.len() == 3 { &args[2] } else { &args[0] };
    let value_area = match criteria::aligned_area(value_spec, ctx, crit_area.rows, crit_area.cols) {
        Some(a) => a,
        None => match as_area(value_spec, ctx) {
            Some(a) => a,
            None => return err(ErrorValue::Value),
        },
    };
    let criterion = criteria::criterion_from_arg(&args[1], ctx);
    average_of(&[(crit_area, criterion)], &value_area, ctx)
}

/// AVERAGEIFS(average_range, crit_range1, crit1, ...).
pub(crate) fn averageifs(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 3 {
        return err(ErrorValue::Value);
    }
    match criteria::collect_pairs(&args[1..], ctx) {
        Some(pairs) => {
            let (rows, cols) = (pairs[0].0.rows, pairs[0].0.cols);
            match criteria::aligned_area(&args[0], ctx, rows, cols) {
                Some(value_area) => average_of(&pairs, &value_area, ctx),
                None => err(ErrorValue::Value),
            }
        }
        None => err(ErrorValue::Value),
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
    let pairs = match criteria::collect_pairs(&args[1..], ctx) {
        Some(pairs) => pairs,
        None => return err(ErrorValue::Value),
    };
    let value_area = match criteria::aligned_area(&args[0], ctx, pairs[0].0.rows, pairs[0].0.cols) {
        Some(area) => area,
        None => return err(ErrorValue::Value),
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

/// NORM.DIST(x, mean, sd, cumulative): the normal distribution's density, or
/// its cumulative probability when `cumulative` is true.
pub(crate) fn norm_dist(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 4 {
        return err(ErrorValue::Value);
    }
    let (x, mean, sd) = match (
        nth_number(args, ctx, 0),
        nth_number(args, ctx, 1),
        nth_number(args, ctx, 2),
    ) {
        (Ok(x), Ok(m), Ok(s)) => (x, m, s),
        (Err(e), _, _) | (_, Err(e), _) | (_, _, Err(e)) => return err(e),
    };
    let cumulative = match crate::eval::to_bool(&evaluate(&args[3], ctx)) {
        Ok(b) => b,
        Err(e) => return err(e),
    };
    if sd <= 0.0 {
        return err(ErrorValue::Num);
    }
    let z = (x - mean) / sd;
    if cumulative {
        finite(0.5 * erfc(-z * std::f64::consts::FRAC_1_SQRT_2))
    } else {
        finite((-0.5 * z * z).exp() / (sd * (2.0 * std::f64::consts::PI).sqrt()))
    }
}

/// the complementary error function, by w. j. cody's rational approximations
/// (algorithm 715); accurate to near machine precision across the real line.
fn erfc(x: f64) -> f64 {
    const A: [f64; 5] = [
        3.161_123_743_870_565_6e0,
        1.138_641_541_510_501_6e2,
        3.774_852_376_853_02e2,
        3.209_377_589_138_469_5e3,
        1.857_777_061_846_031_5e-1,
    ];
    const B: [f64; 4] = [
        2.360_129_095_234_412e1,
        2.440_246_379_344_441_7e2,
        1.282_616_526_077_372_3e3,
        2.844_236_833_439_171e3,
    ];
    const C: [f64; 9] = [
        5.641_884_969_886_701e-1,
        8.883_149_794_388_376e0,
        6.611_919_063_714_163e1,
        2.986_351_381_974_001e2,
        8.819_522_212_417_69e2,
        1.712_047_612_634_070_6e3,
        2.051_078_377_826_071_5e3,
        1.230_339_354_797_997_3e3,
        2.153_115_354_744_038_5e-8,
    ];
    const D: [f64; 8] = [
        1.574_492_611_070_983_5e1,
        1.176_939_508_913_125e2,
        5.371_811_018_620_099e2,
        1.621_389_574_566_690_2e3,
        3.290_799_235_733_459_6e3,
        4.362_619_090_143_247e3,
        3.439_367_674_143_721_5e3,
        1.230_339_354_803_749_4e3,
    ];
    const P: [f64; 6] = [
        3.053_266_349_612_323_4e-1,
        3.603_448_999_498_044_4e-1,
        1.257_817_261_112_292_5e-1,
        1.608_378_514_874_228e-2,
        6.587_491_615_298_378e-4,
        1.631_538_713_730_209_8e-2,
    ];
    const Q: [f64; 5] = [
        2.568_520_192_289_822,
        1.872_952_849_923_460_5e0,
        5.279_051_029_514_284e-1,
        6.051_834_131_244_132e-2,
        2.335_204_976_268_691_8e-3,
    ];
    /// 1/sqrt(pi)
    const SQRT_1_PI: f64 = 5.641_895_835_477_563e-1;

    let y = x.abs();
    let value = if y <= 0.468_75 {
        let z = y * y;
        let mut numerator = A[4] * z;
        let mut denominator = z;
        for index in 0..3 {
            numerator = (numerator + A[index]) * z;
            denominator = (denominator + B[index]) * z;
        }
        return 1.0 - x * (numerator + A[3]) / (denominator + B[3]);
    } else if y <= 4.0 {
        let mut numerator = C[8] * y;
        let mut denominator = y;
        for index in 0..7 {
            numerator = (numerator + C[index]) * y;
            denominator = (denominator + D[index]) * y;
        }
        scaled_exp(y) * (numerator + C[7]) / (denominator + D[7])
    } else if y >= 26.543 {
        0.0
    } else {
        let z = 1.0 / (y * y);
        let mut numerator = P[5] * z;
        let mut denominator = z;
        for index in 0..4 {
            numerator = (numerator + P[index]) * z;
            denominator = (denominator + Q[index]) * z;
        }
        let tail = z * (numerator + P[4]) / (denominator + Q[4]);
        scaled_exp(y) * (SQRT_1_PI - tail) / y
    };
    if x < 0.0 { 2.0 - value } else { value }
}

/// `exp(-y*y)` split around a 1/16-grid point so the squaring loses no bits.
fn scaled_exp(y: f64) -> f64 {
    let grid = (y * 16.0).trunc() / 16.0;
    let delta = (y - grid) * (y + grid);
    (-grid * grid).exp() * (-delta).exp()
}

/// AVERAGEA(value, ...): the mean with text as zero and logicals as one/zero.
pub(crate) fn averagea(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match collect_numbers_anytype(args, ctx) {
        Ok(nums) if nums.is_empty() => err(ErrorValue::Div0),
        Ok(nums) => num(nums.iter().sum::<f64>() / nums.len() as f64),
        Err(e) => err(e),
    }
}

/// MAXA(value, ...): like MAX, but text and logicals count.
pub(crate) fn maxa(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match collect_numbers_anytype(args, ctx) {
        Ok(nums) if nums.is_empty() => num(0.0),
        Ok(nums) => num(nums.iter().copied().fold(f64::NEG_INFINITY, f64::max)),
        Err(e) => err(e),
    }
}

/// MINA(value, ...): like MIN, but text and logicals count.
pub(crate) fn mina(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match collect_numbers_anytype(args, ctx) {
        Ok(nums) if nums.is_empty() => num(0.0),
        Ok(nums) => num(nums.iter().copied().fold(f64::INFINITY, f64::min)),
        Err(e) => err(e),
    }
}

/// STDEVA(value, ...): the sample standard deviation with text as zero.
pub(crate) fn stdeva(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match collect_numbers_anytype(args, ctx) {
        Ok(nums) if nums.len() < 2 => err(ErrorValue::Div0),
        Ok(nums) => finite(centered(&nums).1.sqrt()),
        Err(e) => err(e),
    }
}

/// GEOMEAN(number, ...): the nth root of the product; every value must be
/// strictly positive.
pub(crate) fn geomean(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let nums = match collect_numbers_deep(args, ctx) {
        Ok(nums) => nums,
        Err(e) => return err(e),
    };
    if nums.is_empty() || nums.iter().any(|x| *x <= 0.0) {
        return err(ErrorValue::Num);
    }
    // the logs keep a long product from overflowing before the root shrinks it
    let mean_log = nums.iter().map(|x| x.ln()).sum::<f64>() / nums.len() as f64;
    finite(mean_log.exp())
}

/// HARMEAN(number, ...): the reciprocal of the mean reciprocal; every value
/// must be strictly positive.
pub(crate) fn harmean(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let nums = match collect_numbers_deep(args, ctx) {
        Ok(nums) => nums,
        Err(e) => return err(e),
    };
    if nums.is_empty() || nums.iter().any(|x| *x <= 0.0) {
        return err(ErrorValue::Num);
    }
    finite(nums.len() as f64 / nums.iter().map(|x| 1.0 / x).sum::<f64>())
}

/// AVEDEV(number, ...): the mean absolute deviation from the mean.
pub(crate) fn avedev(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let nums = match collect_numbers_deep(args, ctx) {
        Ok(nums) => nums,
        Err(e) => return err(e),
    };
    if nums.is_empty() {
        return err(ErrorValue::Num);
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let total: f64 = nums.iter().map(|x| (x - mean).abs()).sum();
    finite(total / nums.len() as f64)
}

/// DEVSQ(number, ...): the sum of squared deviations from the mean.
pub(crate) fn devsq(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let nums = match collect_numbers_deep(args, ctx) {
        Ok(nums) => nums,
        Err(e) => return err(e),
    };
    if nums.is_empty() {
        return err(ErrorValue::Num);
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    finite(nums.iter().map(|x| (x - mean) * (x - mean)).sum())
}

/// SKEW(number, ...): the sample skewness; needs at least three values with
/// some spread between them.
pub(crate) fn skew(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let nums = match collect_numbers_deep(args, ctx) {
        Ok(nums) => nums,
        Err(e) => return err(e),
    };
    let n = nums.len() as f64;
    let (mean, variance) = centered(&nums);
    if nums.len() < 3 || variance <= 0.0 {
        return err(ErrorValue::Div0);
    }
    let sd = variance.sqrt();
    let total: f64 = nums.iter().map(|x| ((x - mean) / sd).powi(3)).sum();
    finite(n / ((n - 1.0) * (n - 2.0)) * total)
}

/// KURT(number, ...): the sample excess kurtosis; needs at least four values
/// with some spread between them.
pub(crate) fn kurt(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    let nums = match collect_numbers_deep(args, ctx) {
        Ok(nums) => nums,
        Err(e) => return err(e),
    };
    let n = nums.len() as f64;
    let (mean, variance) = centered(&nums);
    if nums.len() < 4 || variance <= 0.0 {
        return err(ErrorValue::Div0);
    }
    let sd = variance.sqrt();
    let total: f64 = nums.iter().map(|x| ((x - mean) / sd).powi(4)).sum();
    let scale = n * (n + 1.0) / ((n - 1.0) * (n - 2.0) * (n - 3.0));
    let correction = 3.0 * (n - 1.0) * (n - 1.0) / ((n - 2.0) * (n - 3.0));
    finite(scale * total - correction)
}

/// TRIMMEAN(array, percent): the mean after dropping `percent` of the values,
/// split evenly between the two tails and rounded down to a whole pair.
pub(crate) fn trimmean(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let percent = match nth_number(args, ctx, 1) {
        Ok(p) => p,
        Err(e) => return err(e),
    };
    if !(0.0..1.0).contains(&percent) {
        return err(ErrorValue::Num);
    }
    let mut nums = match collect_numbers_deep(&args[..1], ctx) {
        Ok(nums) => nums,
        Err(e) => return err(e),
    };
    if nums.is_empty() {
        return err(ErrorValue::Num);
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    // excel trims whole pairs, so an odd count rounds down
    let trim = ((nums.len() as f64 * percent) / 2.0).floor() as usize;
    let kept = &nums[trim..nums.len() - trim];
    finite(kept.iter().sum::<f64>() / kept.len() as f64)
}

/// the mean and sample variance of a set, in one pass over the deviations.
fn centered(nums: &[f64]) -> (f64, f64) {
    let n = nums.len() as f64;
    let mean = nums.iter().sum::<f64>() / n;
    if nums.len() < 2 {
        return (mean, 0.0);
    }
    let total: f64 = nums.iter().map(|x| (x - mean) * (x - mean)).sum();
    (mean, total / (n - 1.0))
}
