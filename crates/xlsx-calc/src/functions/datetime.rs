//! date and time functions on excel 1900-system serials, including the
//! deliberate 1900 leap-year bug (serial 60 = phantom 1900-02-29).

use xlsx_model::{CellValue, ErrorValue};

use crate::eval::{EvalContext, err, evaluate, num, to_text};
use crate::parser::Expr;

use super::{nth_int, nth_number, omitted};

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
    if serial < 0 {
        return None;
    }
    // excel calls serial 0 "january 0, 1900", which is what a blank date cell
    // reads as, so YEAR, MONTH and DATEDIF all answer for it
    if serial == 0 {
        return Some((1900, 1, 0));
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
/// WEEKNUM(serial, [type]): the week containing Jan 1 is week 1, and a week
/// starts on the day `type` names. type 21 is the ISO rule and defers to
/// [`isoweeknum`].
pub(crate) fn weeknum(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() || args.len() > 2 {
        return err(ErrorValue::Value);
    }
    let serial = match serial_arg(args, ctx, 0) {
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
    if kind == 21 {
        return iso_week(serial);
    }
    // the weekday index the week starts on, as an offset from sunday
    let start = match kind {
        1 | 17 => 0,
        2 | 11 => 1,
        12 => 2,
        13 => 3,
        14 => 4,
        15 => 5,
        16 => 6,
        _ => return err(ErrorValue::Num),
    };
    // excel's serial 0 is the phantom day before 1900-01-01, and is week 0
    let year = match serial_to_ymd(serial) {
        Some((year, _, _)) => year,
        None if serial == 0 => 1900,
        None => return err(ErrorValue::Num),
    };
    let jan1 = date_to_serial(year, 1, 1);
    if jan1 < 0 {
        return err(ErrorValue::Num);
    }
    // serial 1 is a sunday, so `(serial - 1) % 7` is the offset from sunday
    let offset = ((jan1 - 1).rem_euclid(7) - start).rem_euclid(7);
    num(((serial - jan1 + offset).div_euclid(7) + 1) as f64)
}

/// ISOWEEKNUM(serial): ISO 8601 — weeks start monday, week 1 holds the first
/// thursday of the year.
pub(crate) fn isoweeknum(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    match serial_arg(args, ctx, 0) {
        Ok(_) if args.len() != 1 => err(ErrorValue::Value),
        Ok(n) => iso_week(n.floor() as i64),
        Err(e) => err(e),
    }
}

fn iso_week(serial: i64) -> CellValue {
    if serial_to_ymd(serial).is_none() {
        return err(ErrorValue::Num);
    }
    // monday=0 .. sunday=6; serial 1 is a sunday, so shift by 1
    let weekday = (serial - 1).rem_euclid(7);
    let monday = (weekday + 6) % 7;
    let thursday = serial - monday + 3;
    let Some((iso_year, _, _)) = serial_to_ymd(thursday) else {
        return err(ErrorValue::Num);
    };
    let jan1 = date_to_serial(iso_year, 1, 1);
    if jan1 < 0 {
        return err(ErrorValue::Num);
    }
    num(((thursday - jan1) / 7 + 1) as f64)
}

pub(crate) fn weekday(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.is_empty() || args.len() > 2 {
        return err(ErrorValue::Value);
    }
    let serial = match serial_arg(args, ctx, 0) {
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

/// a date argument as a serial: excel reads text where a date is wanted, which
/// is what the `MONTH(monthname&1)` trick relies on.
fn serial_arg(args: &[Expr], ctx: &EvalContext<'_>, index: usize) -> Result<f64, ErrorValue> {
    match nth_number(args, ctx, index) {
        Ok(value) => Ok(value),
        Err(error) => match args.get(index).map(|arg| evaluate(arg, ctx)) {
            Some(CellValue::Text { value }) => match parse_date_text(&value, ctx) {
                Some(serial) => Ok(serial as f64),
                None => Err(error),
            },
            _ => Err(error),
        },
    }
}

fn ymd_part(args: &[Expr], ctx: &EvalContext<'_>, pick: fn((i64, i64, i64)) -> i64) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    let serial = match serial_arg(args, ctx, 0) {
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
    let serial = match serial_arg(args, ctx, 0) {
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
    let serial = match serial_arg(args, ctx, 0) {
        Ok(n) => n,
        Err(e) => return err(e),
    };
    let frac = serial - serial.floor();
    // round to the nearest second to absorb float noise in the day fraction
    let secs = (frac * 86_400.0).round() as i64 % 86_400;
    num(pick(secs) as f64)
}

/// DATEVALUE(text): the serial of a date written as text. a trailing time is
/// dropped; a two-digit year below 30 is 2000s, otherwise 1900s.
pub(crate) fn datevalue(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() != 1 {
        return err(ErrorValue::Value);
    }
    let raw = match evaluate(&args[0], ctx) {
        CellValue::Error { value } => return err(value),
        CellValue::Text { value } => value,
        _ => return err(ErrorValue::Value),
    };
    match parse_date_text(&raw, ctx) {
        Some(serial) => num(serial as f64),
        None => err(ErrorValue::Value),
    }
}

/// parse the date part of a textual timestamp to a serial, `None` when it is
/// not a date this locale (US order) recognises.
fn parse_date_text(raw: &str, ctx: &EvalContext<'_>) -> Option<i64> {
    let fields = date_fields(raw)?;
    let (y, m, d) = match fields.as_slice() {
        [a, b, c] => three_fields(a, b, c)?,
        [a, b] => two_fields(a, b)?,
        _ => return None,
    };
    let y = match y {
        Some(y) => y,
        None => current_year(ctx)?,
    };
    if !(1900..=9999).contains(&y) || !(1..=12).contains(&m) {
        return None;
    }
    // excel's phantom 1900-02-29 parses even though 1900 was not a leap year
    if (y, m, d) == (1900, 2, 29) {
        return Some(PHANTOM);
    }
    if d < 1 || d > days_in_month(y, m) {
        return None;
    }
    Some(date_to_serial(y, m, d))
}

/// the date words and numbers of a timestamp: time-of-day chunks, meridiem
/// markers and the separators between fields are stripped.
fn date_fields(raw: &str) -> Option<Vec<String>> {
    let mut fields: Vec<String> = Vec::new();
    let mut current = String::new();
    for chunk in raw.split_whitespace() {
        if chunk.contains(':') {
            break;
        }
        if chunk.eq_ignore_ascii_case("am")
            || chunk.eq_ignore_ascii_case("pm")
            || chunk.eq_ignore_ascii_case("a.m.")
            || chunk.eq_ignore_ascii_case("p.m.")
        {
            continue;
        }
        for ch in chunk.chars() {
            if ch.is_alphanumeric() {
                // a letter meeting a digit starts a new field, so the
                // `MONTH(monthname&1)` trick reads "Nov1" as two
                if current
                    .chars()
                    .last()
                    .is_some_and(|last| last.is_ascii_digit() != ch.is_ascii_digit())
                {
                    fields.push(std::mem::take(&mut current));
                }
                current.push(ch);
            } else if matches!(ch, '/' | '-' | ',' | '.') {
                if !current.is_empty() {
                    fields.push(std::mem::take(&mut current));
                }
            } else {
                return None;
            }
        }
        if !current.is_empty() {
            fields.push(std::mem::take(&mut current));
        }
    }
    (fields.len() == 2 || fields.len() == 3).then_some(fields)
}

/// a three-field date: `yyyy-m-d` when the first field is a four-digit year,
/// otherwise US `m/d/y` with either field allowed to be a month name.
fn three_fields(a: &str, b: &str, c: &str) -> Option<(Option<i64>, i64, i64)> {
    if let Some(month) = month_name(a) {
        return Some((Some(year_field(c)?), month, b.parse().ok()?));
    }
    if let Some(month) = month_name(b) {
        return Some((Some(year_field(c)?), month, a.parse().ok()?));
    }
    if a.len() == 4 && a.chars().all(|ch| ch.is_ascii_digit()) {
        return Some((Some(a.parse().ok()?), b.parse().ok()?, c.parse().ok()?));
    }
    Some((Some(year_field(c)?), a.parse().ok()?, b.parse().ok()?))
}

/// a two-field date: a month and a day in the current year, unless the other
/// field is a year (four digits, or any number a day cannot be).
fn two_fields(a: &str, b: &str) -> Option<(Option<i64>, i64, i64)> {
    if let Some(month) = month_name(a) {
        return Some(match year_only(b) {
            Some(year) => (Some(year), month, 1),
            None => (None, month, b.parse().ok()?),
        });
    }
    if let Some(month) = month_name(b) {
        return Some(match year_only(a) {
            Some(year) => (Some(year), month, 1),
            None => (None, month, a.parse().ok()?),
        });
    }
    let month: i64 = a.parse().ok()?;
    match year_only(b) {
        Some(year) => Some((Some(year), month, 1)),
        None => Some((None, month, b.parse().ok()?)),
    }
}

/// a field that can only be a year: four digits, or a number above 31.
fn year_only(field: &str) -> Option<i64> {
    let value: i64 = field.parse().ok()?;
    (field.len() == 4 || value > 31).then_some(value)
}

/// a year field, windowing two-digit years as excel does: 0..=29 is 2000s.
fn year_field(field: &str) -> Option<i64> {
    let value: i64 = field.parse().ok()?;
    if value < 0 {
        return None;
    }
    Some(match field.len() {
        1 | 2 if value < 30 => 2000 + value,
        1 | 2 => 1900 + value,
        _ => value,
    })
}

fn month_name(field: &str) -> Option<i64> {
    const MONTHS: [&str; 12] = [
        "january",
        "february",
        "march",
        "april",
        "may",
        "june",
        "july",
        "august",
        "september",
        "october",
        "november",
        "december",
    ];
    let lower = field.to_ascii_lowercase();
    if lower.len() < 3 {
        return None;
    }
    MONTHS
        .iter()
        .position(|full| *full == lower || full.starts_with(&lower))
        .map(|index| index as i64 + 1)
}

fn current_year(ctx: &EvalContext<'_>) -> Option<i64> {
    let serial = ctx.now_serial?.floor() as i64;
    serial_to_ymd(serial).map(|(y, _, _)| y)
}

/// YEARFRAC(start, end, [basis]): the fraction of a year between two dates
/// under day-count bases 0..=4. the order of the dates does not matter.
pub(crate) fn yearfrac(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 2 || args.len() > 3 {
        return err(ErrorValue::Value);
    }
    let (start, end) = match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(a), Ok(b)) => (a.floor() as i64, b.floor() as i64),
        (Err(e), _) | (_, Err(e)) => return err(e),
    };
    let basis = if args.len() == 3 && !omitted(&args[2]) {
        match nth_int(args, ctx, 2) {
            Ok(b) => b,
            Err(e) => return err(e),
        }
    } else {
        0
    };
    if start < 0 || end < 0 {
        return err(ErrorValue::Num);
    }
    let (lo, hi) = if start <= end {
        (start, end)
    } else {
        (end, start)
    };
    match year_fraction(lo, hi, basis) {
        Some(value) => num(value),
        None => err(ErrorValue::Num),
    }
}

fn year_fraction(lo: i64, hi: i64, basis: i64) -> Option<f64> {
    let actual = (hi - lo) as f64;
    match basis {
        0 => Some(days_30_360(lo, hi, false)? / 360.0),
        1 => Some(actual / actual_denominator(lo, hi)?),
        2 => Some(actual / 360.0),
        3 => Some(actual / 365.0),
        4 => Some(days_30_360(lo, hi, true)? / 360.0),
        _ => None,
    }
}

/// day count on a 30/360 calendar; `european` clamps both days to 30, the US
/// rule instead folds the 31st and the end of february onto the 30th.
fn days_30_360(lo: i64, hi: i64, european: bool) -> Option<f64> {
    let (y1, m1, mut d1) = serial_to_ymd(lo.max(1))?;
    let (y2, m2, mut d2) = serial_to_ymd(hi.max(1))?;
    if european {
        d1 = d1.min(30);
        d2 = d2.min(30);
    } else {
        if m1 == 2 && d1 == days_in_month(y1, 2) {
            if m2 == 2 && d2 == days_in_month(y2, 2) {
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
    Some(((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) as f64)
}

/// the denominator basis 1 divides by: the length of the single year a short
/// span sits in, or the average year length across the years a long span
/// touches.
fn actual_denominator(lo: i64, hi: i64) -> Option<f64> {
    let (y1, m1, d1) = serial_to_ymd(lo.max(1))?;
    let (y2, m2, d2) = serial_to_ymd(hi.max(1))?;
    let within_a_year = y1 == y2 || (y1 + 1 == y2 && (m1, d1) >= (m2, d2));
    if !within_a_year {
        let years = (y2 - y1 + 1) as f64;
        let days: i64 = (y1..=y2).map(|y| if is_leap(y) { 366 } else { 365 }).sum();
        return Some(days as f64 / years);
    }
    if y1 == y2 && is_leap(y1) {
        return Some(366.0);
    }
    let leap_day = |y: i64| is_leap(y).then(|| date_to_serial(y, 2, 29));
    let covered = [y1, y2]
        .into_iter()
        .filter_map(leap_day)
        .any(|serial| (lo..=hi).contains(&serial));
    Some(if covered { 366.0 } else { 365.0 })
}

/// the weekend pattern of the `.INTL` workday functions: seven flags starting
/// at monday.
#[derive(Clone, Copy)]
struct Weekend([bool; 7]);

impl Weekend {
    /// `true` when the serial falls on a weekend day. serial 1 is a sunday, so
    /// `serial % 7 == 2` is a monday.
    fn covers(&self, serial: i64) -> bool {
        self.0[(serial - 2).rem_euclid(7) as usize]
    }

    fn days(&self) -> usize {
        self.0.iter().filter(|day| **day).count()
    }
}

/// read the weekend argument: a seven-character mask starting at monday, or
/// one of excel's numbered patterns (1..=7 pairs, 11..=17 single days).
fn weekend_of(value: &CellValue) -> Result<Weekend, ErrorValue> {
    if let CellValue::Text { value } = value {
        let flags: Vec<bool> = value.chars().map(|ch| ch == '1').collect();
        if flags.len() != 7 || value.chars().any(|ch| ch != '0' && ch != '1') {
            return Err(ErrorValue::Value);
        }
        let mut mask = [false; 7];
        mask.copy_from_slice(&flags);
        return Ok(Weekend(mask));
    }
    let code = crate::eval::to_number(value)?.trunc() as i64;
    let mut mask = [false; 7];
    match code {
        1..=7 => {
            mask[(code + 4) as usize % 7] = true;
            mask[(code + 5) as usize % 7] = true;
        }
        11..=17 => mask[(code - 12).rem_euclid(7) as usize] = true,
        _ => return Err(ErrorValue::Num),
    }
    Ok(Weekend(mask))
}

fn weekend_arg(args: &[Expr], ctx: &EvalContext<'_>, index: usize) -> Result<Weekend, ErrorValue> {
    match args.get(index) {
        None => weekend_of(&CellValue::Number { value: 1.0 }),
        Some(arg) if omitted(arg) => weekend_of(&CellValue::Number { value: 1.0 }),
        Some(arg) => match evaluate(arg, ctx) {
            CellValue::Empty => weekend_of(&CellValue::Number { value: 1.0 }),
            value => weekend_of(&value),
        },
    }
}

/// the holiday serials of the optional trailing argument.
fn holidays_arg(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    index: usize,
) -> Result<Vec<i64>, ErrorValue> {
    let Some(arg) = args.get(index).filter(|arg| !omitted(arg)) else {
        return Ok(Vec::new());
    };
    let mut out = Vec::new();
    for value in crate::array::evaluate_array(arg, ctx).into_array().cells() {
        match value {
            CellValue::Number { value } => out.push(value.floor() as i64),
            CellValue::Error { value } => return Err(*value),
            CellValue::Text { value } => match crate::eval::parse_num(value) {
                Some(number) => out.push(number.floor() as i64),
                None => return Err(ErrorValue::Value),
            },
            _ => {}
        }
    }
    Ok(out)
}

/// NETWORKDAYS.INTL(start, end, [weekend], [holidays]): working days between
/// two dates, both ends included; a range that runs backwards counts negative.
pub(crate) fn networkdays_intl(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 2 || args.len() > 4 {
        return err(ErrorValue::Value);
    }
    let weekend = match weekend_arg(args, ctx, 2) {
        Ok(w) => w,
        Err(e) => return err(e),
    };
    networkdays_core(args, ctx, weekend, 3)
}

/// NETWORKDAYS(start, end, [holidays]): the saturday/sunday weekend spelled
/// out, so the plain form is the `.INTL` one with its weekend fixed.
pub(crate) fn networkdays(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 2 || args.len() > 3 {
        return err(ErrorValue::Value);
    }
    match default_weekend() {
        Ok(weekend) => networkdays_core(args, ctx, weekend, 2),
        Err(e) => err(e),
    }
}

fn networkdays_core(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    weekend: Weekend,
    holidays_index: usize,
) -> CellValue {
    let (start, end) = match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(a), Ok(b)) => (a.floor() as i64, b.floor() as i64),
        (Err(e), _) | (_, Err(e)) => return err(e),
    };
    let holidays = match holidays_arg(args, ctx, holidays_index) {
        Ok(h) => h,
        Err(e) => return err(e),
    };
    if start < 0 || end < 0 {
        return err(ErrorValue::Num);
    }
    let (lo, hi) = if start <= end {
        (start, end)
    } else {
        (end, start)
    };
    let mut count = workdays_between(lo, hi, weekend);
    let mut seen: Vec<i64> = Vec::new();
    for holiday in holidays {
        if (lo..=hi).contains(&holiday) && !weekend.covers(holiday) && !seen.contains(&holiday) {
            seen.push(holiday);
            count -= 1;
        }
    }
    num(if start <= end {
        count as f64
    } else {
        -count as f64
    })
}

/// non-weekend days in an inclusive serial range, counted by whole weeks plus
/// the remainder so a wide range costs no more than a narrow one.
fn workdays_between(lo: i64, hi: i64, weekend: Weekend) -> i64 {
    let span = hi - lo + 1;
    let weeks = span / 7;
    let mut count = weeks * (7 - weekend.days() as i64);
    for offset in 0..span % 7 {
        if !weekend.covers(lo + weeks * 7 + offset) {
            count += 1;
        }
    }
    count
}

/// WORKDAY.INTL(start, days, [weekend], [holidays]): the date `days` working
/// days from `start`, skipping weekend days and holidays.
pub(crate) fn workday_intl(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 2 || args.len() > 4 {
        return err(ErrorValue::Value);
    }
    let weekend = match weekend_arg(args, ctx, 2) {
        Ok(w) => w,
        Err(e) => return err(e),
    };
    workday_core(args, ctx, weekend, 3)
}

/// WORKDAY(start, days, [holidays]): the `.INTL` form with the weekend fixed
/// to saturday and sunday.
pub(crate) fn workday(args: &[Expr], ctx: &EvalContext<'_>) -> CellValue {
    if args.len() < 2 || args.len() > 3 {
        return err(ErrorValue::Value);
    }
    match default_weekend() {
        Ok(weekend) => workday_core(args, ctx, weekend, 2),
        Err(e) => err(e),
    }
}

fn default_weekend() -> Result<Weekend, ErrorValue> {
    weekend_of(&CellValue::Number { value: 1.0 })
}

fn workday_core(
    args: &[Expr],
    ctx: &EvalContext<'_>,
    weekend: Weekend,
    holidays_index: usize,
) -> CellValue {
    let (start, days) = match (nth_number(args, ctx, 0), nth_number(args, ctx, 1)) {
        (Ok(a), Ok(b)) => (a.floor() as i64, b.trunc() as i64),
        (Err(e), _) | (_, Err(e)) => return err(e),
    };
    let holidays = match holidays_arg(args, ctx, holidays_index) {
        Ok(h) => h,
        Err(e) => return err(e),
    };
    if start < 0 {
        return err(ErrorValue::Num);
    }
    if weekend.days() == 7 {
        return err(ErrorValue::Value);
    }
    let step = if days < 0 { -1 } else { 1 };
    let mut remaining = days.abs();
    let mut current = start;
    // whole weeks land on the same weekday, so only the remainder needs walking
    let per_week = 7 - weekend.days() as i64;
    let weeks = remaining / per_week;
    if weeks > 0 && holidays.is_empty() {
        current += step * weeks * 7;
        remaining -= weeks * per_week;
        if !(0..=MAX_SERIAL).contains(&current) {
            return err(ErrorValue::Num);
        }
    }
    while remaining > 0 {
        current += step;
        if !(0..=MAX_SERIAL).contains(&current) {
            return err(ErrorValue::Num);
        }
        if !weekend.covers(current) && !holidays.contains(&current) {
            remaining -= 1;
        }
    }
    num(current as f64)
}

/// 9999-12-31, the last date excel can hold.
const MAX_SERIAL: i64 = 2_958_465;
