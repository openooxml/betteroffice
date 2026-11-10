//! Integer request fields whose values the parsed document model holds as
//! `f64`. A host that forwards the parsed model through serde writes those
//! values as `1.0`, so these fields accept a whole-number float as the
//! integer and reject a fractional or out-of-range one, naming the field.

use serde::de::Error as _;
use serde::{Deserialize, Deserializer};
use serde_json::Number;

pub(crate) trait WholeNumber: Sized {
    const KIND: &'static str;
    fn from_number(number: &Number) -> Option<Self>;
}

const TWO_POW_63: f64 = 9_223_372_036_854_775_808.0;
const TWO_POW_64: f64 = 18_446_744_073_709_551_616.0;

impl WholeNumber for i64 {
    const KIND: &'static str = "an integer";
    fn from_number(number: &Number) -> Option<Self> {
        number.as_i64().or_else(|| {
            let value = number.as_f64()?;
            (value.fract() == 0.0 && (-TWO_POW_63..TWO_POW_63).contains(&value))
                .then_some(value as i64)
        })
    }
}

impl WholeNumber for u64 {
    const KIND: &'static str = "a nonnegative integer";
    fn from_number(number: &Number) -> Option<Self> {
        number.as_u64().or_else(|| {
            let value = number.as_f64()?;
            (value.fract() == 0.0 && (0.0..TWO_POW_64).contains(&value)).then_some(value as u64)
        })
    }
}

fn convert<T: WholeNumber, E: serde::de::Error>(number: &Number, field: &str) -> Result<T, E> {
    T::from_number(number)
        .ok_or_else(|| E::custom(format!("{field} must be {}, got {number}", T::KIND)))
}

pub(crate) fn required<'de, T: WholeNumber, D: Deserializer<'de>>(
    deserializer: D,
    field: &str,
) -> Result<T, D::Error> {
    convert(&Number::deserialize(deserializer)?, field)
}

pub(crate) fn optional<'de, T: WholeNumber, D: Deserializer<'de>>(
    deserializer: D,
    field: &str,
) -> Result<Option<T>, D::Error> {
    Option::<Number>::deserialize(deserializer)?
        .map(|number| convert(&number, field))
        .transpose()
}

macro_rules! whole_fields {
    ($($name:ident: $kind:ident $ty:ty = $field:literal;)*) => {
        $(
            pub(crate) fn $name<'de, D: serde::Deserializer<'de>>(
                deserializer: D,
            ) -> Result<whole_fields!(@ty $kind $ty), D::Error> {
                $crate::whole_number::$kind(deserializer, $field)
            }
        )*
    };
    (@ty required $ty:ty) => { $ty };
    (@ty optional $ty:ty) => { Option<$ty> };
}

whole_fields! {
    note_id: required i64 = "note id";
    note_num_start: optional i64 = "numStart";
    footnote_columns: optional u64 = "footnoteColumns";
    page_start: optional u64 = "page numbering start";
    relative_height: optional u64 = "relativeHeight";
    grid_before: optional u64 = "gridBefore";
    grid_after: optional u64 = "gridAfter";
    control_id: optional i64 = "controlId";
    page_restart: optional u64 = "restart";
}

#[cfg(test)]
mod tests {
    use serde::Deserialize;

    #[derive(Debug, Deserialize)]
    struct Signed {
        #[serde(deserialize_with = "super::note_id")]
        id: i64,
    }

    #[derive(Debug, Deserialize)]
    struct Unsigned {
        #[serde(default, deserialize_with = "super::footnote_columns")]
        columns: Option<u64>,
    }

    fn signed(json: &str) -> Result<i64, String> {
        serde_json::from_str::<Signed>(json).map(|value| value.id).map_err(|error| error.to_string())
    }

    fn unsigned(json: &str) -> Result<Option<u64>, String> {
        serde_json::from_str::<Unsigned>(json)
            .map(|value| value.columns)
            .map_err(|error| error.to_string())
    }

    #[test]
    fn whole_floats_read_as_the_integer() {
        assert_eq!(signed(r#"{"id":1.0}"#), Ok(1));
        assert_eq!(signed(r#"{"id":-3.0}"#), Ok(-3));
        assert_eq!(signed(r#"{"id":0.0}"#), Ok(0));
        assert_eq!(unsigned(r#"{"columns":2.0}"#), Ok(Some(2)));
        assert_eq!(unsigned(r#"{"columns":null}"#), Ok(None));
        assert_eq!(unsigned("{}"), Ok(None));
    }

    #[test]
    fn fractional_negative_and_out_of_range_values_are_rejected_by_name() {
        for json in [r#"{"id":1.5}"#, r#"{"id":1e300}"#, r#"{"id":-1e300}"#] {
            let error = signed(json).unwrap_err();
            assert!(error.contains("note id must be an integer"), "{error}");
        }
        for json in [r#"{"columns":-1.0}"#, r#"{"columns":-1}"#, r#"{"columns":0.5}"#, r#"{"columns":1e30}"#] {
            let error = unsigned(json).unwrap_err();
            assert!(error.contains("footnoteColumns must be a nonnegative integer"), "{error}");
        }
        assert!(signed(r#"{"id":"1"}"#).is_err());
    }

    #[test]
    fn every_js_safe_integer_reads_the_same_as_integer_and_as_whole_float() {
        const MAX_SAFE: i64 = (1 << 53) - 1;
        let mut state = 0x9e37_79b9_7f4a_7c15_u64;
        let mut samples = vec![0, 1, -1, MAX_SAFE, -MAX_SAFE, MAX_SAFE - 1, 1 << 31, -(1 << 31)];
        for _ in 0..10_000 {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            samples.push((state % (2 * MAX_SAFE as u64 + 1)) as i64 - MAX_SAFE);
        }
        for value in samples {
            assert_eq!(signed(&format!(r#"{{"id":{value}}}"#)), Ok(value));
            assert_eq!(signed(&format!(r#"{{"id":{value}.0}}"#)), Ok(value));
            assert_eq!(signed(&format!(r#"{{"id":{}}}"#, serde_json::to_string(&(value as f64)).unwrap())), Ok(value));
            if value >= 0 {
                let expected = Ok(Some(value as u64));
                assert_eq!(unsigned(&format!(r#"{{"columns":{value}}}"#)), expected);
                assert_eq!(unsigned(&format!(r#"{{"columns":{value}.0}}"#)), expected);
            }
        }
    }
}
