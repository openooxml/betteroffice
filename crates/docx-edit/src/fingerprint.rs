//! Streaming content fingerprints for typed layout values.
//!
//! [`fingerprint_without_positions`] walks a value through its `Serialize`
//! impl and feeds every scalar to a 64-bit SipHash, so no intermediate JSON
//! tree is built. Object keys named `pmStart`, `pmEnd`, `docStart` and
//! `docEnd` are skipped at every depth: those absolute document positions
//! shift when text is edited earlier in the story without changing what the
//! value measures or paints. Fingerprints are only compared against others
//! from the same session.

use std::fmt;
use std::hash::{DefaultHasher, Hasher as _};

use serde::Serialize;
use serde::ser::{self, Serializer};

const POSITION_KEYS: [&str; 4] = ["pmStart", "pmEnd", "docStart", "docEnd"];

const TAG_NULL: u64 = 1;
const TAG_FALSE: u64 = 2;
const TAG_TRUE: u64 = 3;
const TAG_I64: u64 = 4;
const TAG_U64: u64 = 5;
const TAG_F64: u64 = 6;
const TAG_STRING: u64 = 7;
const TAG_ARRAY: u64 = 8;
const TAG_OBJECT: u64 = 9;
const TAG_END: u64 = 10;
const TAG_KEY: u64 = 11;
const TAG_BYTES: u64 = 12;

/// Fingerprint of `value` with absolute document positions left out.
pub(crate) fn fingerprint_without_positions<T: Serialize + ?Sized>(
    value: &T,
) -> Result<u64, String> {
    let mut hasher = Hasher::<false>::new();
    value.serialize(&mut hasher).map_err(|error| error.0)?;
    Ok(hasher.finish())
}

/// Fingerprint of `value`'s positions relative to `origin`, the enclosing block's
/// pm position a retained display shifts by; absolute without one.
pub(crate) fn positions_fingerprint<T: Serialize + ?Sized>(
    value: &T,
    origin: Option<f64>,
) -> Result<u64, String> {
    let mut hasher = Hasher::<true>::new();
    hasher.origin = origin;
    value.serialize(&mut hasher).map_err(|error| error.0)?;
    Ok(hasher.finish())
}

fn position_key(key: &str) -> Option<usize> {
    POSITION_KEYS.iter().position(|position| *position == key)
}

/// SipHash-1-3 with fixed keys: deterministic within a process, which is all
/// a session-local fingerprint needs.
struct Hasher<const RELATIVE_POSITIONS: bool = false> {
    inner: DefaultHasher,
    position_family: Option<usize>,
    origin: Option<f64>,
}

impl<const RELATIVE_POSITIONS: bool> Hasher<RELATIVE_POSITIONS> {
    fn new() -> Self {
        Self {
            inner: DefaultHasher::new(),
            position_family: None,
            origin: None,
        }
    }

    fn word(&mut self, word: u64) {
        if !RELATIVE_POSITIONS {
            self.inner.write_u64(word);
        }
    }

    fn bytes(&mut self, bytes: &[u8]) {
        if !RELATIVE_POSITIONS {
            self.inner.write_u64(bytes.len() as u64);
            self.inner.write(bytes);
        }
    }

    fn relative_position(&mut self, value: f64) {
        if self.position_family.is_some() {
            let relative = self.origin.map_or(value, |origin| value - origin);
            self.inner.write_u64(TAG_F64);
            self.inner.write_u64(if relative == 0.0 {
                0
            } else {
                relative.to_bits()
            });
        }
    }

    fn finish(&self) -> u64 {
        self.inner.finish()
    }
}

#[derive(Debug)]
struct Error(String);

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

impl std::error::Error for Error {}

impl ser::Error for Error {
    fn custom<T: fmt::Display>(message: T) -> Self {
        Self(message.to_string())
    }
}

/// Map keys are hashed on their own first, so a position key can be recognized
/// (and dropped with its value) whatever type serialized it.
struct KeyProbe {
    key: Hasher,
    position: Option<usize>,
}

impl KeyProbe {
    fn new() -> Self {
        Self {
            key: Hasher::new(),
            position: None,
        }
    }
}

struct Container<'a, const RELATIVE_POSITIONS: bool> {
    hasher: &'a mut Hasher<RELATIVE_POSITIONS>,
    skip_value: bool,
    position: Option<usize>,
}

impl<const RELATIVE_POSITIONS: bool> Container<'_, RELATIVE_POSITIONS> {
    fn field<T: Serialize + ?Sized>(&mut self, key: &str, value: &T) -> Result<(), Error> {
        let position = position_key(key);
        if RELATIVE_POSITIONS {
            self.position = position;
            return self.value(value);
        }
        if position.is_some() {
            return Ok(());
        }
        self.hasher.word(TAG_KEY);
        self.hasher.bytes(key.as_bytes());
        value.serialize(&mut *self.hasher)
    }

    fn key<T: Serialize + ?Sized>(&mut self, key: &T) -> Result<(), Error> {
        let mut probe = KeyProbe::new();
        key.serialize(&mut probe)?;
        self.position = probe.position;
        self.skip_value = !RELATIVE_POSITIONS && probe.position.is_some();
        if !RELATIVE_POSITIONS && probe.position.is_none() {
            self.hasher.word(TAG_KEY);
            self.hasher.word(probe.key.finish());
        }
        Ok(())
    }

    fn value<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        if std::mem::take(&mut self.skip_value) {
            return Ok(());
        }
        if RELATIVE_POSITIONS {
            let position = self.position.take();
            if let Some(key) = position {
                self.hasher.inner.write_u64(TAG_KEY);
                self.hasher.inner.write_u64(key as u64);
            }
            let previous_family = self.hasher.position_family;
            self.hasher.position_family = position.map(|key| key / 2);
            let result = value.serialize(&mut *self.hasher);
            self.hasher.position_family = previous_family;
            return result;
        }
        value.serialize(&mut *self.hasher)
    }

    fn end(self) -> Result<(), Error> {
        self.hasher.word(TAG_END);
        Ok(())
    }
}

impl<'a, const RELATIVE_POSITIONS: bool> Serializer for &'a mut Hasher<RELATIVE_POSITIONS> {
    type Ok = ();
    type Error = Error;
    type SerializeSeq = Container<'a, RELATIVE_POSITIONS>;
    type SerializeTuple = Container<'a, RELATIVE_POSITIONS>;
    type SerializeTupleStruct = Container<'a, RELATIVE_POSITIONS>;
    type SerializeTupleVariant = Container<'a, RELATIVE_POSITIONS>;
    type SerializeMap = Container<'a, RELATIVE_POSITIONS>;
    type SerializeStruct = Container<'a, RELATIVE_POSITIONS>;
    type SerializeStructVariant = Container<'a, RELATIVE_POSITIONS>;

    fn serialize_bool(self, value: bool) -> Result<(), Error> {
        self.word(if value { TAG_TRUE } else { TAG_FALSE });
        Ok(())
    }
    fn serialize_i8(self, value: i8) -> Result<(), Error> {
        self.serialize_i64(value.into())
    }
    fn serialize_i16(self, value: i16) -> Result<(), Error> {
        self.serialize_i64(value.into())
    }
    fn serialize_i32(self, value: i32) -> Result<(), Error> {
        self.serialize_i64(value.into())
    }
    fn serialize_i64(self, value: i64) -> Result<(), Error> {
        if RELATIVE_POSITIONS {
            self.relative_position(value as f64);
            return Ok(());
        }
        self.word(TAG_I64);
        self.word(value as u64);
        Ok(())
    }
    fn serialize_u8(self, value: u8) -> Result<(), Error> {
        self.serialize_u64(value.into())
    }
    fn serialize_u16(self, value: u16) -> Result<(), Error> {
        self.serialize_u64(value.into())
    }
    fn serialize_u32(self, value: u32) -> Result<(), Error> {
        self.serialize_u64(value.into())
    }
    fn serialize_u64(self, value: u64) -> Result<(), Error> {
        if RELATIVE_POSITIONS {
            self.relative_position(value as f64);
            return Ok(());
        }
        self.word(TAG_U64);
        self.word(value);
        Ok(())
    }
    fn serialize_f32(self, value: f32) -> Result<(), Error> {
        self.serialize_f64(value.into())
    }
    fn serialize_f64(self, value: f64) -> Result<(), Error> {
        if RELATIVE_POSITIONS {
            self.relative_position(value);
            return Ok(());
        }
        if value.is_finite() {
            self.word(TAG_F64);
            self.word(value.to_bits());
        } else {
            self.word(TAG_NULL);
        }
        Ok(())
    }
    fn serialize_char(self, value: char) -> Result<(), Error> {
        self.serialize_str(value.encode_utf8(&mut [0; 4]))
    }
    fn serialize_str(self, value: &str) -> Result<(), Error> {
        self.word(TAG_STRING);
        self.bytes(value.as_bytes());
        Ok(())
    }
    fn serialize_bytes(self, value: &[u8]) -> Result<(), Error> {
        self.word(TAG_BYTES);
        self.bytes(value);
        Ok(())
    }
    fn serialize_none(self) -> Result<(), Error> {
        if RELATIVE_POSITIONS && self.position_family.is_some() {
            self.inner.write_u64(TAG_NULL);
        } else {
            self.word(TAG_NULL);
        }
        Ok(())
    }
    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<(), Error> {
        value.serialize(self)
    }
    fn serialize_unit(self) -> Result<(), Error> {
        self.serialize_none()
    }
    fn serialize_unit_struct(self, _: &'static str) -> Result<(), Error> {
        self.serialize_unit()
    }
    fn serialize_unit_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
    ) -> Result<(), Error> {
        self.serialize_str(variant)
    }
    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        value: &T,
    ) -> Result<(), Error> {
        value.serialize(self)
    }
    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        value: &T,
    ) -> Result<(), Error> {
        self.word(TAG_OBJECT);
        let mut container = Container {
            hasher: self,
            skip_value: false,
            position: None,
        };
        container.field(variant, value)?;
        container.end()
    }
    fn serialize_seq(self, _: Option<usize>) -> Result<Container<'a, RELATIVE_POSITIONS>, Error> {
        self.word(TAG_ARRAY);
        Ok(Container {
            hasher: self,
            skip_value: false,
            position: None,
        })
    }
    fn serialize_tuple(self, len: usize) -> Result<Container<'a, RELATIVE_POSITIONS>, Error> {
        self.serialize_seq(Some(len))
    }
    fn serialize_tuple_struct(
        self,
        _: &'static str,
        len: usize,
    ) -> Result<Container<'a, RELATIVE_POSITIONS>, Error> {
        self.serialize_seq(Some(len))
    }
    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        len: usize,
    ) -> Result<Container<'a, RELATIVE_POSITIONS>, Error> {
        self.word(TAG_OBJECT);
        self.word(TAG_KEY);
        self.bytes(variant.as_bytes());
        self.serialize_seq(Some(len))
    }
    fn serialize_map(self, _: Option<usize>) -> Result<Container<'a, RELATIVE_POSITIONS>, Error> {
        self.word(TAG_OBJECT);
        Ok(Container {
            hasher: self,
            skip_value: false,
            position: None,
        })
    }
    fn serialize_struct(
        self,
        _: &'static str,
        len: usize,
    ) -> Result<Container<'a, RELATIVE_POSITIONS>, Error> {
        self.serialize_map(Some(len))
    }
    fn serialize_struct_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        len: usize,
    ) -> Result<Container<'a, RELATIVE_POSITIONS>, Error> {
        self.word(TAG_OBJECT);
        self.word(TAG_KEY);
        self.bytes(variant.as_bytes());
        self.serialize_map(Some(len))
    }
}

impl<const RELATIVE_POSITIONS: bool> ser::SerializeSeq for Container<'_, RELATIVE_POSITIONS> {
    type Ok = ();
    type Error = Error;
    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        value.serialize(&mut *self.hasher)
    }
    fn end(self) -> Result<(), Error> {
        Container::end(self)
    }
}

impl<const RELATIVE_POSITIONS: bool> ser::SerializeTuple for Container<'_, RELATIVE_POSITIONS> {
    type Ok = ();
    type Error = Error;
    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        value.serialize(&mut *self.hasher)
    }
    fn end(self) -> Result<(), Error> {
        Container::end(self)
    }
}

impl<const RELATIVE_POSITIONS: bool> ser::SerializeTupleStruct
    for Container<'_, RELATIVE_POSITIONS>
{
    type Ok = ();
    type Error = Error;
    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        value.serialize(&mut *self.hasher)
    }
    fn end(self) -> Result<(), Error> {
        Container::end(self)
    }
}

impl<const RELATIVE_POSITIONS: bool> ser::SerializeTupleVariant
    for Container<'_, RELATIVE_POSITIONS>
{
    type Ok = ();
    type Error = Error;
    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        value.serialize(&mut *self.hasher)
    }
    fn end(self) -> Result<(), Error> {
        self.hasher.word(TAG_END);
        Container::end(self)
    }
}

impl<const RELATIVE_POSITIONS: bool> ser::SerializeMap for Container<'_, RELATIVE_POSITIONS> {
    type Ok = ();
    type Error = Error;
    fn serialize_key<T: Serialize + ?Sized>(&mut self, key: &T) -> Result<(), Error> {
        self.key(key)
    }
    fn serialize_value<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        self.value(value)
    }
    fn end(self) -> Result<(), Error> {
        Container::end(self)
    }
}

impl<const RELATIVE_POSITIONS: bool> ser::SerializeStruct for Container<'_, RELATIVE_POSITIONS> {
    type Ok = ();
    type Error = Error;
    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), Error> {
        self.field(key, value)
    }
    fn end(self) -> Result<(), Error> {
        Container::end(self)
    }
}

impl<const RELATIVE_POSITIONS: bool> ser::SerializeStructVariant
    for Container<'_, RELATIVE_POSITIONS>
{
    type Ok = ();
    type Error = Error;
    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), Error> {
        self.field(key, value)
    }
    fn end(self) -> Result<(), Error> {
        self.hasher.word(TAG_END);
        Container::end(self)
    }
}

impl Serializer for &mut KeyProbe {
    type Ok = ();
    type Error = Error;
    type SerializeSeq = ser::Impossible<(), Error>;
    type SerializeTuple = ser::Impossible<(), Error>;
    type SerializeTupleStruct = ser::Impossible<(), Error>;
    type SerializeTupleVariant = ser::Impossible<(), Error>;
    type SerializeMap = ser::Impossible<(), Error>;
    type SerializeStruct = ser::Impossible<(), Error>;
    type SerializeStructVariant = ser::Impossible<(), Error>;

    fn serialize_str(self, value: &str) -> Result<(), Error> {
        self.position = position_key(value);
        self.key.serialize_str(value)
    }
    fn serialize_bool(self, value: bool) -> Result<(), Error> {
        self.key.serialize_bool(value)
    }
    fn serialize_i8(self, value: i8) -> Result<(), Error> {
        self.key.serialize_i8(value)
    }
    fn serialize_i16(self, value: i16) -> Result<(), Error> {
        self.key.serialize_i16(value)
    }
    fn serialize_i32(self, value: i32) -> Result<(), Error> {
        self.key.serialize_i32(value)
    }
    fn serialize_i64(self, value: i64) -> Result<(), Error> {
        self.key.serialize_i64(value)
    }
    fn serialize_u8(self, value: u8) -> Result<(), Error> {
        self.key.serialize_u8(value)
    }
    fn serialize_u16(self, value: u16) -> Result<(), Error> {
        self.key.serialize_u16(value)
    }
    fn serialize_u32(self, value: u32) -> Result<(), Error> {
        self.key.serialize_u32(value)
    }
    fn serialize_u64(self, value: u64) -> Result<(), Error> {
        self.key.serialize_u64(value)
    }
    fn serialize_f32(self, value: f32) -> Result<(), Error> {
        self.key.serialize_f32(value)
    }
    fn serialize_f64(self, value: f64) -> Result<(), Error> {
        self.key.serialize_f64(value)
    }
    fn serialize_char(self, value: char) -> Result<(), Error> {
        self.serialize_str(value.encode_utf8(&mut [0; 4]))
    }
    fn serialize_bytes(self, value: &[u8]) -> Result<(), Error> {
        self.key.serialize_bytes(value)
    }
    fn serialize_none(self) -> Result<(), Error> {
        self.key.serialize_none()
    }
    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<(), Error> {
        value.serialize(self)
    }
    fn serialize_unit(self) -> Result<(), Error> {
        self.key.serialize_unit()
    }
    fn serialize_unit_struct(self, _: &'static str) -> Result<(), Error> {
        self.key.serialize_unit()
    }
    fn serialize_unit_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
    ) -> Result<(), Error> {
        self.serialize_str(variant)
    }
    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        value: &T,
    ) -> Result<(), Error> {
        value.serialize(self)
    }
    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: &T,
    ) -> Result<(), Error> {
        Err(Error("unsupported map key".to_owned()))
    }
    fn serialize_seq(self, _: Option<usize>) -> Result<Self::SerializeSeq, Error> {
        Err(Error("unsupported map key".to_owned()))
    }
    fn serialize_tuple(self, _: usize) -> Result<Self::SerializeTuple, Error> {
        Err(Error("unsupported map key".to_owned()))
    }
    fn serialize_tuple_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleStruct, Error> {
        Err(Error("unsupported map key".to_owned()))
    }
    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleVariant, Error> {
        Err(Error("unsupported map key".to_owned()))
    }
    fn serialize_map(self, _: Option<usize>) -> Result<Self::SerializeMap, Error> {
        Err(Error("unsupported map key".to_owned()))
    }
    fn serialize_struct(self, _: &'static str, _: usize) -> Result<Self::SerializeStruct, Error> {
        Err(Error("unsupported map key".to_owned()))
    }
    fn serialize_struct_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStructVariant, Error> {
        Err(Error("unsupported map key".to_owned()))
    }
}

#[cfg(test)]
mod tests {
    use serde::Serialize;
    use serde_json::json;

    use super::{fingerprint_without_positions as fingerprint, positions_fingerprint};

    #[derive(Serialize)]
    #[serde(rename_all = "camelCase")]
    struct Formatting {
        bold: bool,
        doc_start: Option<f64>,
    }

    #[derive(Serialize)]
    #[serde(rename_all = "camelCase")]
    struct Run {
        #[serde(flatten)]
        fmt: Formatting,
        text: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        pm_start: Option<f64>,
        #[serde(skip_serializing_if = "Option::is_none")]
        width: Option<f64>,
        #[serde(skip_serializing_if = "Option::is_none")]
        height: Option<f64>,
        extra: serde_json::Value,
    }

    fn run(text: &str, position: f64) -> Run {
        Run {
            fmt: Formatting {
                bold: false,
                doc_start: Some(position),
            },
            text: text.to_owned(),
            pm_start: Some(position),
            width: Some(10.0),
            height: None,
            extra: json!({ "pmEnd": position + 4.0, "nested": [{ "docEnd": position }] }),
        }
    }

    #[test]
    fn absolute_positions_do_not_change_the_fingerprint() {
        assert_eq!(
            fingerprint(&run("text", 1.0)),
            fingerprint(&run("text", 90.0))
        );
    }

    #[test]
    fn positions_ignore_a_shift_of_the_whole_block() {
        let base = positions_fingerprint(&run("text", 1.0), Some(1.0)).unwrap();
        assert_eq!(
            base,
            positions_fingerprint(&run("other text", 90.0), Some(90.0)).unwrap()
        );
        let mut doc_moved = run("text", 1.0);
        doc_moved.fmt.doc_start = Some(101.0);
        doc_moved.extra["nested"][0]["docEnd"] = json!(101.0);
        assert_ne!(base, positions_fingerprint(&doc_moved, Some(1.0)).unwrap());

        // Without an enclosing origin positions stay absolute, nested pm ones too.
        let nested = |position: f64| json!({ "docStart": position, "inner": [{ "pmStart": position, "pmEnd": position + 1.0 }] });
        assert_ne!(
            positions_fingerprint(&nested(100.0), None).unwrap(),
            positions_fingerprint(&nested(200.0), None).unwrap()
        );
    }

    #[test]
    fn positions_detect_moves_within_the_block() {
        let base = positions_fingerprint(&run("text", 1.0), Some(1.0)).unwrap();
        let mut moved = run("text", 1.0);
        moved.extra["pmEnd"] = json!(6.0);
        assert_ne!(base, positions_fingerprint(&moved, Some(1.0)).unwrap());

        let mut moved = run("text", 1.0);
        moved.extra["nested"][0]["docEnd"] = json!(2.0);
        assert_ne!(base, positions_fingerprint(&moved, Some(1.0)).unwrap());

        let mut missing = run("text", 1.0);
        missing.fmt.doc_start = None;
        assert_ne!(base, positions_fingerprint(&missing, Some(1.0)).unwrap());
    }

    #[test]
    fn content_changes_the_fingerprint() {
        let base = fingerprint(&run("text", 1.0)).unwrap();
        assert_ne!(base, fingerprint(&run("texT", 1.0)).unwrap());

        let mut bold = run("text", 1.0);
        bold.fmt.bold = true;
        assert_ne!(base, fingerprint(&bold).unwrap());

        let mut wider = run("text", 1.0);
        wider.width = Some(10.5);
        assert_ne!(base, fingerprint(&wider).unwrap());

        let mut moved = run("text", 1.0);
        moved.height = moved.width.take();
        assert_ne!(base, fingerprint(&moved).unwrap());

        let mut extra = run("text", 1.0);
        extra.extra["nested"][0]["kind"] = json!("x");
        assert_ne!(base, fingerprint(&extra).unwrap());
    }

    #[test]
    fn crafted_word_patterns_do_not_collide() {
        assert_ne!(
            fingerprint(&run("AAAAAAAAAAAAAAAA", 1.0)).unwrap(),
            fingerprint(&run(">*!!!!!!)=vfdO4y", 1.0)).unwrap()
        );
        assert_ne!(
            fingerprint(&json!([10.0, 20.0, 30.0])).unwrap(),
            fingerprint(&json!([20.0, 10.0, 30.0])).unwrap()
        );
    }

    #[test]
    fn sequence_boundaries_are_part_of_the_fingerprint() {
        assert_ne!(
            fingerprint(&json!([["a"], []])).unwrap(),
            fingerprint(&json!([[], ["a"]])).unwrap()
        );
        assert_ne!(
            fingerprint(&json!({ "a": "bc" })).unwrap(),
            fingerprint(&json!({ "ab": "c" })).unwrap()
        );
    }
}
