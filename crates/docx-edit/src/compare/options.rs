//! Comparison options and the limits a caller may tighten below the v1 ceilings.

use serde::{Deserialize, Serialize};

/// The token unit text differences are reported in.
#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) enum Granularity {
    /// Unicode word boundaries, keeping punctuation and whitespace tokens.
    #[default]
    Word,
    /// Extended grapheme clusters.
    Char,
}

/// Whether inspection stops at the first blocking condition or keeps collecting diagnostics.
/// Neither returns a redline while a blocking condition exists.
#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) enum UnsupportedPolicy {
    #[default]
    Fail,
    Report,
}

#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(crate) struct LimitsWire {
    pub max_input_bytes: Option<u64>,
    pub max_expanded_bytes: Option<u64>,
    pub max_paragraphs: Option<u64>,
    pub max_text_units: Option<u64>,
    pub max_alignment_cells: Option<u64>,
    pub max_diff_cells: Option<u64>,
    pub max_changes: Option<u64>,
    pub max_diagnostics: Option<u64>,
    pub max_staged_bytes: Option<u64>,
    pub max_result_bytes: Option<u64>,
    pub max_output_bytes: Option<u64>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(crate) struct CompareOptionsWire {
    pub author: String,
    pub date: String,
    #[serde(default)]
    pub granularity: Option<Granularity>,
    #[serde(default)]
    pub unsupported: Option<UnsupportedPolicy>,
    #[serde(default)]
    pub limits: Option<LimitsWire>,
}

/// Every bound a comparison enforces.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct CompareLimits {
    pub max_input_bytes: usize,
    pub max_expanded_bytes: usize,
    pub max_paragraphs: usize,
    pub max_text_units: usize,
    pub max_alignment_cells: usize,
    pub max_diff_cells: usize,
    pub max_changes: usize,
    pub max_diagnostics: usize,
    pub max_staged_bytes: usize,
    pub max_result_bytes: usize,
    pub max_output_bytes: usize,
}

const MIB: usize = 1024 * 1024;

/// The v1 ceilings, which are also the defaults.
pub(crate) const CEILINGS: CompareLimits = CompareLimits {
    max_input_bytes: 32 * MIB,
    max_expanded_bytes: 128 * MIB,
    max_paragraphs: 10_000,
    max_text_units: 1_048_576,
    max_alignment_cells: 250_000,
    max_diff_cells: 4_000_000,
    max_changes: 128,
    max_diagnostics: 256,
    max_staged_bytes: 64 * MIB,
    max_result_bytes: 8 * MIB,
    max_output_bytes: 64 * MIB,
};

/// The smallest `maxResultBytes`, which leaves room for the refusal that reports the limit.
const MIN_RESULT_BYTES: usize = 1024;

/// Paragraphs either side of an unresolved alignment gap may hold.
pub(crate) const MAX_GAP_PARAGRAPHS: usize = 64;
/// UTF-16 units an author name may hold.
const MAX_AUTHOR_UNITS: usize = 255;

#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) struct CompareOptions {
    pub author: String,
    /// The revision date, normalized to UTC (`YYYY-MM-DDTHH:MM:SS[.fff]Z`).
    pub date: String,
    pub granularity: Granularity,
    pub unsupported: UnsupportedPolicy,
    pub limits: CompareLimits,
}

impl CompareOptions {
    /// Validates `wire`; the error explains why the options are unusable.
    pub(crate) fn resolve(wire: CompareOptionsWire) -> Result<Self, String> {
        let author = wire.author;
        if author.trim().is_empty() {
            return Err("author must not be blank".to_owned());
        }
        if author.encode_utf16().count() > MAX_AUTHOR_UNITS {
            return Err(format!(
                "author may hold at most {MAX_AUTHOR_UNITS} UTF-16 units"
            ));
        }
        if author.chars().any(char::is_control) {
            return Err("author must not contain control characters".to_owned());
        }
        let date = normalize_date(&wire.date)?;
        let requested = wire.limits.unwrap_or_default();
        let limit = |name: &str, value: Option<u64>, ceiling: usize| -> Result<usize, String> {
            match value {
                None => Ok(ceiling),
                Some(0) => Err(format!("limits.{name} must be at least 1")),
                Some(value) if value > ceiling as u64 => Err(format!(
                    "limits.{name} may be at most {ceiling}, the v1 ceiling"
                )),
                Some(value) => Ok(value as usize),
            }
        };
        let limits = CompareLimits {
            max_input_bytes: limit(
                "maxInputBytes",
                requested.max_input_bytes,
                CEILINGS.max_input_bytes,
            )?,
            max_expanded_bytes: limit(
                "maxExpandedBytes",
                requested.max_expanded_bytes,
                CEILINGS.max_expanded_bytes,
            )?,
            max_paragraphs: limit(
                "maxParagraphs",
                requested.max_paragraphs,
                CEILINGS.max_paragraphs,
            )?,
            max_text_units: limit(
                "maxTextUnits",
                requested.max_text_units,
                CEILINGS.max_text_units,
            )?,
            max_alignment_cells: limit(
                "maxAlignmentCells",
                requested.max_alignment_cells,
                CEILINGS.max_alignment_cells,
            )?,
            max_diff_cells: limit(
                "maxDiffCells",
                requested.max_diff_cells,
                CEILINGS.max_diff_cells,
            )?,
            max_changes: limit("maxChanges", requested.max_changes, CEILINGS.max_changes)?,
            max_diagnostics: limit(
                "maxDiagnostics",
                requested.max_diagnostics,
                CEILINGS.max_diagnostics,
            )?,
            max_staged_bytes: limit(
                "maxStagedBytes",
                requested.max_staged_bytes,
                CEILINGS.max_staged_bytes,
            )?,
            max_result_bytes: limit(
                "maxResultBytes",
                requested.max_result_bytes,
                CEILINGS.max_result_bytes,
            )?,
            max_output_bytes: limit(
                "maxOutputBytes",
                requested.max_output_bytes,
                CEILINGS.max_output_bytes,
            )?,
        };
        if limits.max_result_bytes < MIN_RESULT_BYTES {
            return Err(format!(
                "limits.maxResultBytes must be at least {MIN_RESULT_BYTES}"
            ));
        }
        Ok(Self {
            author,
            date,
            granularity: wire.granularity.unwrap_or_default(),
            unsupported: wire.unsupported.unwrap_or_default(),
            limits,
        })
    }
}

/// Days since 1970-01-01 of a proleptic Gregorian date.
fn days_from_civil(year: i64, month: i64, day: i64) -> i64 {
    let year = if month <= 2 { year - 1 } else { year };
    let era = year.div_euclid(400);
    let year_of_era = year - era * 400;
    let month_index = (month + 9) % 12;
    let day_of_year = (153 * month_index + 2) / 5 + day - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    era * 146_097 + day_of_era - 719_468
}

fn civil_from_days(days: i64) -> (i64, i64, i64) {
    let days = days + 719_468;
    let era = days.div_euclid(146_097);
    let day_of_era = days - era * 146_097;
    let year_of_era =
        (day_of_era - day_of_era / 1460 + day_of_era / 36_524 - day_of_era / 146_096) / 365;
    let day_of_year = day_of_era - (365 * year_of_era + year_of_era / 4 - year_of_era / 100);
    let month_index = (5 * day_of_year + 2) / 153;
    let day = day_of_year - (153 * month_index + 2) / 5 + 1;
    let month = if month_index < 10 {
        month_index + 3
    } else {
        month_index - 9
    };
    let year = year_of_era + era * 400 + i64::from(month <= 2);
    (year, month, day)
}

fn days_in_month(year: i64, month: i64) -> i64 {
    match month {
        2 if (year % 4 == 0 && year % 100 != 0) || year % 400 == 0 => 29,
        2 => 28,
        4 | 6 | 9 | 11 => 30,
        _ => 31,
    }
}

/// Validates an RFC 3339 timestamp with an explicit offset and returns it in UTC, keeping any
/// fractional seconds without trailing zeros.
pub(crate) fn normalize_date(value: &str) -> Result<String, String> {
    let invalid = || {
        format!(
            "date {value:?} is not an RFC 3339 timestamp with an explicit offset, such as 2024-01-31T09:30:00Z"
        )
    };
    let bytes = value.as_bytes();
    let digits = |range: std::ops::Range<usize>| -> Result<i64, String> {
        let slice = bytes.get(range).ok_or_else(invalid)?;
        if slice.is_empty() || !slice.iter().all(u8::is_ascii_digit) {
            return Err(invalid());
        }
        std::str::from_utf8(slice)
            .ok()
            .and_then(|text| text.parse().ok())
            .ok_or_else(invalid)
    };
    let separator = |index: usize, allowed: &[u8]| {
        bytes
            .get(index)
            .filter(|byte| allowed.contains(byte))
            .map(|_| ())
            .ok_or_else(invalid)
    };
    let year = digits(0..4)?;
    separator(4, b"-")?;
    let month = digits(5..7)?;
    separator(7, b"-")?;
    let day = digits(8..10)?;
    separator(10, b"Tt")?;
    let hour = digits(11..13)?;
    separator(13, b":")?;
    let minute = digits(14..16)?;
    separator(16, b":")?;
    let second = digits(17..19)?;
    let mut cursor = 19;
    let mut fraction = String::new();
    if bytes.get(cursor) == Some(&b'.') {
        cursor += 1;
        while let Some(byte) = bytes.get(cursor).filter(|byte| byte.is_ascii_digit()) {
            fraction.push(char::from(*byte));
            cursor += 1;
        }
        if fraction.is_empty() || fraction.len() > 9 {
            return Err(invalid());
        }
    }
    let offset_minutes = match bytes.get(cursor) {
        Some(b'Z' | b'z') if cursor + 1 == bytes.len() => 0,
        Some(sign @ (b'+' | b'-')) if cursor + 6 == bytes.len() => {
            let hours = digits(cursor + 1..cursor + 3)?;
            separator(cursor + 3, b":")?;
            let minutes = digits(cursor + 4..cursor + 6)?;
            if hours > 23 || minutes > 59 {
                return Err(invalid());
            }
            let total = hours * 60 + minutes;
            if *sign == b'-' { -total } else { total }
        }
        _ => return Err(invalid()),
    };
    if !(1..=12).contains(&month)
        || day < 1
        || day > days_in_month(year, month)
        || hour > 23
        || minute > 59
        || second > 59
    {
        return Err(invalid());
    }
    let local = days_from_civil(year, month, day) * 1440 + hour * 60 + minute;
    let utc = local - offset_minutes;
    let (year, month, day) = civil_from_days(utc.div_euclid(1440));
    let minute_of_day = utc.rem_euclid(1440);
    if !(1..=9999).contains(&year) {
        return Err(invalid());
    }
    let fraction = fraction.trim_end_matches('0');
    let fraction = if fraction.is_empty() {
        String::new()
    } else {
        format!(".{fraction}")
    };
    Ok(format!(
        "{year:04}-{month:02}-{day:02}T{:02}:{:02}:{second:02}{fraction}Z",
        minute_of_day / 60,
        minute_of_day % 60
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn wire(author: &str, date: &str) -> CompareOptionsWire {
        CompareOptionsWire {
            author: author.to_owned(),
            date: date.to_owned(),
            granularity: None,
            unsupported: None,
            limits: None,
        }
    }

    #[test]
    fn dates_normalize_to_utc() {
        assert_eq!(
            normalize_date("2024-01-31T09:30:00Z").unwrap(),
            "2024-01-31T09:30:00Z"
        );
        assert_eq!(
            normalize_date("2024-03-01T01:15:00+02:30").unwrap(),
            "2024-02-29T22:45:00Z"
        );
        assert_eq!(
            normalize_date("2023-12-31T23:59:59.1200-00:30").unwrap(),
            "2024-01-01T00:29:59.12Z"
        );
        assert_eq!(
            normalize_date("2024-05-01t10:00:00.000z").unwrap(),
            "2024-05-01T10:00:00Z"
        );
        for bad in [
            "2024-01-31",
            "2024-01-31T09:30:00",
            "2024-02-30T00:00:00Z",
            "2023-02-29T00:00:00Z",
            "2024-01-31T24:00:00Z",
            "2024-01-31T09:30:60Z",
            "2024-01-31T09:30:00+24:00",
            "2024-01-31T09:30:00.Z",
            "2024-01-31 09:30:00Z",
            "0000-01-01T00:00:00Z",
            "0001-01-01T00:00:00+00:01",
            "９９９９-01-01T00:00:00Z",
        ] {
            assert!(normalize_date(bad).is_err(), "{bad}");
        }
    }

    #[test]
    fn options_require_an_author_and_stay_within_the_ceilings() {
        let resolved = CompareOptions::resolve(wire("Reviewer", "2024-01-31T09:30:00Z")).unwrap();
        assert_eq!(resolved.limits, CEILINGS);
        assert_eq!(resolved.granularity, Granularity::Word);
        assert_eq!(resolved.unsupported, UnsupportedPolicy::Fail);
        assert!(CompareOptions::resolve(wire("  ", "2024-01-31T09:30:00Z")).is_err());
        assert!(CompareOptions::resolve(wire("A\nB", "2024-01-31T09:30:00Z")).is_err());
        let mut tightened = wire("A", "2024-01-31T09:30:00Z");
        tightened.limits = Some(LimitsWire {
            max_changes: Some(3),
            ..LimitsWire::default()
        });
        assert_eq!(
            CompareOptions::resolve(tightened)
                .unwrap()
                .limits
                .max_changes,
            3
        );
        for limits in [
            LimitsWire {
                max_changes: Some(129),
                ..LimitsWire::default()
            },
            LimitsWire {
                max_diagnostics: Some(0),
                ..LimitsWire::default()
            },
        ] {
            let mut options = wire("A", "2024-01-31T09:30:00Z");
            options.limits = Some(limits);
            assert!(CompareOptions::resolve(options).is_err());
        }
    }
}
