//! statistical functions. numeric aggregations ignore text/bool/blank inside
//! references, coerce literal arguments, propagate errors.

use std::collections::HashMap;

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{Area, EvalContext, as_area, err, evaluate, num};
use crate::parser::Expr;

use super::criteria::{self, Criterion};
use super::{collect_numbers, nth_int, nth_number};

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
    let values = match area.values(ctx) {
        Ok(values) => values,
        Err(error) => return err(error),
    };
    let nums: Vec<f64> = values
        .into_iter()
        .filter_map(|v| match v {
            CellValue::Number { value } => Some(value),
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
                let values = match area.values(ctx) {
                    Ok(values) => values,
                    Err(error) => return err(error),
                };
                count += values
                    .iter()
                    .filter(|v| matches!(v, CellValue::Number { .. }))
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
                let values = match area.values(ctx) {
                    Ok(values) => values,
                    Err(error) => return err(error),
                };
                count += values
                    .iter()
                    .filter(|v| !matches!(v, CellValue::Empty))
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
    let values = match area.values(ctx) {
        Ok(values) => values,
        Err(error) => return err(error),
    };
    let n = values
        .iter()
        .filter(|v| {
            matches!(v, CellValue::Empty)
                || matches!(v, CellValue::Text { value } if value.is_empty())
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
    criteria::matching_indices(pairs, ctx)?
        .into_iter()
        .map(|i| {
            let (r, c) = (i / cols, i % cols);
            value_area.get(ctx, r, c)
        })
        .collect::<Result<Vec<_>, _>>()
        .map(|values| {
            values
                .into_iter()
                .filter_map(|value| match value {
                    CellValue::Number { value } => Some(value),
                    _ => None,
                })
                .collect()
        })
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
