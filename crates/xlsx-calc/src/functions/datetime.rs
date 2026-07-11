//! date and time functions on excel 1900-system serials, including the
//! deliberate 1900 leap-year bug (serial 60 = phantom 1900-02-29).

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{EvalContext, err, evaluate, num, to_text};
use crate::parser::Expr;

use super::{nth_int, nth_number};

/// the phantom 1900-02-29; serials above it are shifted by one real day.
const PHANTOM: i64 = 60;
/// serial = unix day count + 25568 below the phantom day (serial 1 = 1900-01-01).
const SERIAL_OFFSET: i64 = 25_568;

// howard hinnant's civil-date algorithms
/// days since 1970-01-01 for a proleptic gregorian date.
fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let doy = (153 * (if m > 2 { m - 3 } else { m + 9 }) + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

/// (year, month, day) for a day count since 1970-01-01.
fn civil_from_days(z: i64) -> (i64, i64, i64) {
    let z = z + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    (if m <= 2 { y + 1 } else { y }, m, d)
}

/// convert a unix day count to a 1900-system serial, inserting the phantom day.
fn unix_to_serial(unix: i64) -> i64 {
    let base = unix + SERIAL_OFFSET;
    if base < PHANTOM { base } else { base + 1 }
}

/// (year, month, day) for a serial. serial 60 is the phantom 1900-02-29;
/// serial < 1 has no calendar date here (returns None).
fn serial_to_ymd(serial: i64) -> Option<(i64, i64, i64)> {
    if serial < 1 {
        return None;
    }
    if serial == PHANTOM {
        return Some((1900, 2, 29));
    }
    let adjusted = if serial > PHANTOM { serial - 1 } else { serial };
    Some(civil_from_days(adjusted - SERIAL_OFFSET))
}

/// serial for a (year, month, day) under excel's DATE rules: months and day
/// overflow roll over, phantom 1900-02-29 maps to serial 60.
fn date_to_serial(year: i64, month: i64, day: i64) -> i64 {
    if year == 1900 && month == 2 && day == 29 {
        return PHANTOM;
    }
    let mut y = year;
    let mut m = month;
    y += (m - 1).div_euclid(12);
    m = (m - 1).rem_euclid(12) + 1;
    unix_to_serial(days_from_civil(y, m, 1) + (day - 1))
}

fn days_in_month(year: i64, month: i64) -> i64 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if is_leap(year) => 29,
        2 => 28,
        _ => 30,
    }
}

fn is_leap(y: i64) -> bool {
    (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
}

/// DATE(year, month, day). years 0..=1899 are treated as 1900+year (excel).
pub(crate) fn date(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let (y, m, d) = match (
        nth_int(args, ctx, 0),
        nth_int(args, ctx, 1),
        nth_int(args, ctx, 2),
    ) {
        (Ok(y), Ok(m), Ok(d)) => (y, m, d),
        (Err(e), _, _) | (_, Err(e), _) | (_, _, Err(e)) => return err(e),
    };
    let y = if (0..1900).contains(&y) { y + 1900 } else { y };
    let serial = date_to_serial(y, m, d);
    if serial < 0 {
        err(ErrorValue::Num)
    } else {
        num(serial as f64)
    }
}

pub(crate) fn year(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    ymd_part(args, ctx, |(y, _, _)| y)
}

pub(crate) fn month(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    ymd_part(args, ctx, |(_, m, _)| m)
}

pub(crate) fn day(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    ymd_part(args, ctx, |(_, _, d)| d)
}

/// WEEKDAY(serial, [type]): type 1 (default) = 1..7 Sun..Sat; 2/11..17 shift
/// the first day of the week; 3 = 0..6 Mon..Sun.
pub(crate) fn weekday(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() || args.len() > 2 {
        return err(ErrorValue::Value);
    }
    let serial = match nth_number(args, ctx, 0) {
        Ok(n) => n.floor() as i64,
        Err(e) => return err(e),
    };
    let kind = if args.len() == 2 {
        match nth_int(args, ctx, 1) {
            Ok(k) => k,
            Err(e) => return err(e),
        }
    } else {
        1
    };
    // d0: 0=Sat,1=Sun,...,6=Fri (serial 1 = sunday)
    let d0 = serial.rem_euclid(7);
    let result = match kind {
        1 => {
            if d0 == 0 {
                7
            } else {
                d0
            }
        }
        2 | 11 => (d0 + 5).rem_euclid(7) + 1,
        3 => (d0 + 5).rem_euclid(7),
        12..=17 => {
            let first = (2 + (kind - 11)).rem_euclid(7);
            (d0 - first).rem_euclid(7) + 1
        }
        _ => return err(ErrorValue::Num),
    };
    num(result as f64)
}

/// EDATE(start, months): the same day-of-month `months` away, clamped to the
/// target month's length.
pub(crate) fn edate(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    shifted_month(args, ctx, false)
}

/// EOMONTH(start, months): the last day of the month `months` away.
pub(crate) fn eomonth(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    shifted_month(args, ctx, true)
}

/// TODAY(): the injected date, time truncated; no injected clock -> #VALUE!.
pub(crate) fn today(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if !args.is_empty() {
        return err(ErrorValue::Value);
    }
    match ctx.now_serial {
        Some(s) => num(s.floor()),
        None => err(ErrorValue::Value),
    }
}

/// NOW(): the injected date and time. absent clock -> #VALUE!.
pub(crate) fn now(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if !args.is_empty() {
        return err(ErrorValue::Value);
    }
    match ctx.now_serial {
        Some(s) => num(s),
        None => err(ErrorValue::Value),
    }
}

pub(crate) fn hour(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    time_part(args, ctx, |secs| secs / 3600)
}

pub(crate) fn minute(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    time_part(args, ctx, |secs| (secs / 60) % 60)
}

pub(crate) fn second(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    time_part(args, ctx, |secs| secs % 60)
}

/// TIME(hour, minute, second): a fraction of a day in [0, 1); overflow wraps.
pub(crate) fn time(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let (h, m, s) = match (
        nth_number(args, ctx, 0),
        nth_number(args, ctx, 1),
        nth_number(args, ctx, 2),
    ) {
        (Ok(h), Ok(m), Ok(s)) => (h, m, s),
        (Err(e), _, _) | (_, Err(e), _) | (_, _, Err(e)) => return err(e),
    };
    let frac = (h * 3600.0 + m * 60.0 + s) / 86_400.0;
    num(frac.rem_euclid(1.0))
}

/// DATEDIF(start, end, unit): complete intervals between two dates; units are
/// `Y`, `M`, `D`, `YM`, `YD`, `MD`.
pub(crate) fn datedif(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 3 {
        return err(ErrorValue::Value);
    }
    let s1 = match nth_number(args, ctx, 0) {
        Ok(n) => n.floor() as i64,
        Err(e) => return err(e),
    };
    let s2 = match nth_number(args, ctx, 1) {
        Ok(n) => n.floor() as i64,
        Err(e) => return err(e),
    };
    let unit = match to_text(&evaluate(&args[2], ctx)) {
        Ok(s) => s.to_uppercase(),
        Err(e) => return err(e),
    };
    if s2 < s1 {
        return err(ErrorValue::Num);
    }
    let (a, b) = match (serial_to_ymd(s1), serial_to_ymd(s2)) {
        (Some(a), Some(b)) => (a, b),
        _ => return err(ErrorValue::Num),
    };
    let (y1, m1, d1) = a;
    let (y2, m2, d2) = b;
    let value = match unit.as_str() {
        "D" => (s2 - s1) as f64,
        "Y" => {
            let mut years = y2 - y1;
            if (m2, d2) < (m1, d1) {
                years -= 1;
            }
            years as f64
        }
        "M" => complete_months(y1, m1, d1, y2, m2, d2) as f64,
        "YM" => (complete_months(y1, m1, d1, y2, m2, d2) % 12) as f64,
        "MD" => {
            if d2 >= d1 {
                (d2 - d1) as f64
            } else {
                let (py, pm) = if m2 == 1 { (y2 - 1, 12) } else { (y2, m2 - 1) };
                (days_in_month(py, pm) - d1 + d2) as f64
            }
        }
        "YD" => {
            let anchor = if (m1, d1) <= (m2, d2) {
                date_to_serial(y2, m1, d1)
            } else {
                date_to_serial(y2 - 1, m1, d1)
            };
            (s2 - anchor) as f64
        }
        _ => return err(ErrorValue::Value),
    };
    num(value)
}

fn complete_months(y1: i64, m1: i64, d1: i64, y2: i64, m2: i64, d2: i64) -> i64 {
    let mut months = (y2 - y1) * 12 + (m2 - m1);
    if d2 < d1 {
        months -= 1;
    }
    months
}

fn ymd_part(args: &[Expr], ctx: &EvalContext<'_>, pick: fn((i64, i64, i64)) -> i64) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    let serial = match nth_number(args, ctx, 0) {
        Ok(n) => n.floor() as i64,
        Err(e) => return err(e),
    };
    // serial 0 is excel's 1900-01-00: day 0, month 1, year 1900
    if serial == 0 {
        return num(pick((1900, 1, 0)) as f64);
    }
    match serial_to_ymd(serial) {
        Some(ymd) => num(pick(ymd) as f64),
        None => err(ErrorValue::Num),
    }
}

fn shifted_month(args: &[Expr], ctx: &EvalContext<'_>, end_of_month: bool) -> CellValue {
    if args.len() != 2 {
        return err(ErrorValue::Value);
    }
    let serial = match nth_number(args, ctx, 0) {
        Ok(n) => n.floor() as i64,
        Err(e) => return err(e),
    };
    let months = match nth_int(args, ctx, 1) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let (y, m, d) = match serial_to_ymd(serial) {
        Some(v) => v,
        None => return err(ErrorValue::Num),
    };
    let total = (y * 12 + (m - 1)) + months;
    let ty = total.div_euclid(12);
    let tm = total.rem_euclid(12) + 1;
    let target_day = if end_of_month {
        days_in_month(ty, tm)
    } else {
        d.min(days_in_month(ty, tm))
    };
    let serial = date_to_serial(ty, tm, target_day);
    if serial < 0 {
        err(ErrorValue::Num)
    } else {
        num(serial as f64)
    }
}

fn time_part(args: &[Expr], ctx: &EvalContext<'_>, pick: fn(i64) -> i64) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    let serial = match nth_number(args, ctx, 0) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let frac = serial - serial.floor();
    // round to the nearest second to absorb float noise in the day fraction
    let secs = (frac * 86_400.0).round() as i64 % 86_400;
    num(pick(secs) as f64)
}
