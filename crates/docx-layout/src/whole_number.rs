//! Integer request fields whose values the parsed document model holds as
//! `f64`. A host that forwards the parsed model through serde writes those
//! values as `1.0`, so these fields accept a whole-number float as the
//! integer and reject a fractional or out-of-range one, naming the field.

use std::fmt;
use std::marker::PhantomData;

use serde::Deserializer;
use serde::de::{Error, Visitor};

pub(crate) trait WholeNumber: Sized + fmt::Display {
    const KIND: &'static str;
    fn from_i64(value: i64) -> Option<Self>;
    fn from_u64(value: u64) -> Option<Self>;
    fn from_f64(value: f64) -> Option<Self>;
    fn deserialize_typed<'de, D: Deserializer<'de>, V: Visitor<'de, Value = Self>>(
        deserializer: D,
        visitor: V,
    ) -> Result<Self, D::Error>;
}

const TWO_POW_63: f64 = 9_223_372_036_854_775_808.0;
const TWO_POW_64: f64 = 18_446_744_073_709_551_616.0;

impl WholeNumber for i64 {
    const KIND: &'static str = "an integer";
    fn from_i64(value: i64) -> Option<Self> {
        Some(value)
    }
    fn from_u64(value: u64) -> Option<Self> {
        i64::try_from(value).ok()
    }
    fn from_f64(value: f64) -> Option<Self> {
        (value.fract() == 0.0 && (-TWO_POW_63..TWO_POW_63).contains(&value)).then_some(value as i64)
    }
    fn deserialize_typed<'de, D: Deserializer<'de>, V: Visitor<'de, Value = Self>>(
        deserializer: D,
        visitor: V,
    ) -> Result<Self, D::Error> {
        if deserializer.is_human_readable() {
            deserializer.deserialize_any(visitor)
        } else {
            deserializer.deserialize_i64(visitor)
        }
    }
}

impl WholeNumber for u64 {
    const KIND: &'static str = "a nonnegative integer";
    fn from_i64(value: i64) -> Option<Self> {
        u64::try_from(value).ok()
    }
    fn from_u64(value: u64) -> Option<Self> {
        Some(value)
    }
    fn from_f64(value: f64) -> Option<Self> {
        (value.fract() == 0.0 && (0.0..TWO_POW_64).contains(&value)).then_some(value as u64)
    }
    fn deserialize_typed<'de, D: Deserializer<'de>, V: Visitor<'de, Value = Self>>(
        deserializer: D,
        visitor: V,
    ) -> Result<Self, D::Error> {
        if deserializer.is_human_readable() {
            deserializer.deserialize_any(visitor)
        } else {
            deserializer.deserialize_u64(visitor)
        }
    }
}

struct Whole<T> {
    field: &'static str,
    kind: PhantomData<T>,
}

impl<T: WholeNumber> Whole<T> {
    fn reject<E: Error>(&self, value: impl fmt::Display) -> E {
        E::custom(format!("{} must be {}, got {value}", self.field, T::KIND))
    }
}

impl<'de, T: WholeNumber> Visitor<'de> for Whole<T> {
    type Value = T;
    fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
        write!(formatter, "{} as {}", self.field, T::KIND)
    }
    fn visit_i64<E: Error>(self, value: i64) -> Result<T, E> {
        T::from_i64(value).ok_or_else(|| self.reject(value))
    }
    fn visit_u64<E: Error>(self, value: u64) -> Result<T, E> {
        T::from_u64(value).ok_or_else(|| self.reject(value))
    }
    fn visit_f64<E: Error>(self, value: f64) -> Result<T, E> {
        T::from_f64(value).ok_or_else(|| self.reject(value))
    }
}

struct OptionalWhole<T>(Whole<T>);

impl<'de, T: WholeNumber> Visitor<'de> for OptionalWhole<T> {
    type Value = Option<T>;
    fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
        self.0.expecting(formatter)?;
        formatter.write_str(" or null")
    }
    fn visit_none<E: Error>(self) -> Result<Option<T>, E> {
        Ok(None)
    }
    fn visit_unit<E: Error>(self) -> Result<Option<T>, E> {
        Ok(None)
    }
    fn visit_some<D: Deserializer<'de>>(self, deserializer: D) -> Result<Option<T>, D::Error> {
        T::deserialize_typed(deserializer, self.0).map(Some)
    }
}

pub(crate) fn required<'de, T: WholeNumber, D: Deserializer<'de>>(
    deserializer: D,
    field: &'static str,
) -> Result<T, D::Error> {
    T::deserialize_typed(
        deserializer,
        Whole {
            field,
            kind: PhantomData,
        },
    )
}

pub(crate) fn optional<'de, T: WholeNumber, D: Deserializer<'de>>(
    deserializer: D,
    field: &'static str,
) -> Result<Option<T>, D::Error> {
    deserializer.deserialize_option(OptionalWhole(Whole {
        field,
        kind: PhantomData,
    }))
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
        serde_json::from_str::<Signed>(json)
            .map(|value| value.id)
            .map_err(|error| error.to_string())
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
        for json in [
            r#"{"columns":-1.0}"#,
            r#"{"columns":-1}"#,
            r#"{"columns":0.5}"#,
            r#"{"columns":1e30}"#,
        ] {
            let error = unsigned(json).unwrap_err();
            assert!(
                error.contains("footnoteColumns must be a nonnegative integer"),
                "{error}"
            );
        }
        assert!(signed(r#"{"id":"1"}"#).is_err());
    }

    #[test]
    fn every_js_safe_integer_reads_the_same_as_integer_and_as_whole_float() {
        const MAX_SAFE: i64 = (1 << 53) - 1;
        let mut state = 0x9e37_79b9_7f4a_7c15_u64;
        let mut samples = vec![
            0,
            1,
            -1,
            MAX_SAFE,
            -MAX_SAFE,
            MAX_SAFE - 1,
            1 << 31,
            -(1 << 31),
        ];
        for _ in 0..10_000 {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            samples.push((state % (2 * MAX_SAFE as u64 + 1)) as i64 - MAX_SAFE);
        }
        for value in samples {
            assert_eq!(signed(&format!(r#"{{"id":{value}}}"#)), Ok(value));
            assert_eq!(signed(&format!(r#"{{"id":{value}.0}}"#)), Ok(value));
            assert_eq!(
                signed(&format!(
                    r#"{{"id":{}}}"#,
                    serde_json::to_string(&(value as f64)).unwrap()
                )),
                Ok(value)
            );
            if value >= 0 {
                let expected = Ok(Some(value as u64));
                assert_eq!(unsigned(&format!(r#"{{"columns":{value}}}"#)), expected);
                assert_eq!(unsigned(&format!(r#"{{"columns":{value}.0}}"#)), expected);
            }
        }
    }

    #[derive(Debug, Deserialize)]
    struct Inner {
        #[serde(default, deserialize_with = "super::grid_before")]
        grid_before: Option<u64>,
        #[serde(deserialize_with = "super::note_id")]
        id: i64,
    }

    #[derive(Debug, Deserialize)]
    struct Flattened {
        #[serde(flatten)]
        inner: Inner,
    }

    #[test]
    fn flattened_fields_read_whole_floats_and_integers() {
        let read = |json: &str| {
            serde_json::from_str::<Flattened>(json)
                .map(|value| (value.inner.grid_before, value.inner.id))
        };
        assert_eq!(
            read(r#"{"grid_before":2.0,"id":3.0}"#).unwrap(),
            (Some(2), 3)
        );
        assert_eq!(read(r#"{"grid_before":2,"id":-3}"#).unwrap(), (Some(2), -3));
        assert_eq!(read(r#"{"id":3}"#).unwrap(), (None, 3));
        let error = read(r#"{"grid_before":2.5,"id":3}"#)
            .unwrap_err()
            .to_string();
        assert!(
            error.contains("gridBefore must be a nonnegative integer, got 2.5"),
            "{error}"
        );
    }

    struct Binary(Option<i64>);

    impl<'de> serde::Deserializer<'de> for Binary {
        type Error = serde::de::value::Error;

        fn deserialize_any<V: serde::de::Visitor<'de>>(
            self,
            _: V,
        ) -> Result<V::Value, Self::Error> {
            Err(serde::de::Error::custom("deserialize_any is not supported"))
        }

        fn deserialize_i64<V: serde::de::Visitor<'de>>(
            self,
            visitor: V,
        ) -> Result<V::Value, Self::Error> {
            visitor.visit_i64(self.0.expect("a value"))
        }

        fn deserialize_u64<V: serde::de::Visitor<'de>>(
            self,
            visitor: V,
        ) -> Result<V::Value, Self::Error> {
            visitor.visit_u64(self.0.expect("a value") as u64)
        }

        fn deserialize_option<V: serde::de::Visitor<'de>>(
            self,
            visitor: V,
        ) -> Result<V::Value, Self::Error> {
            match self.0 {
                None => visitor.visit_none(),
                Some(_) => visitor.visit_some(self),
            }
        }

        fn is_human_readable(&self) -> bool {
            false
        }

        serde::forward_to_deserialize_any! {
            bool i8 i16 i32 i128 u8 u16 u32 u128 f32 f64 char str string bytes byte_buf unit
            unit_struct newtype_struct seq tuple tuple_struct map struct enum identifier ignored_any
        }
    }

    #[test]
    fn binary_formats_read_typed_integers_without_deserialize_any() {
        assert_eq!(super::note_id(Binary(Some(7))).unwrap(), 7);
        assert_eq!(super::footnote_columns(Binary(Some(3))).unwrap(), Some(3));
        assert_eq!(super::footnote_columns(Binary(None)).unwrap(), None);
        assert_eq!(super::note_num_start(Binary(Some(-2))).unwrap(), Some(-2));
    }
}
