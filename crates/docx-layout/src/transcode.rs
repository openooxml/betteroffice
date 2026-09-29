//! Typed-to-typed conversion through serde without a JSON tree.
//!
//! The display compiler reads its own input types, which mirror the resident
//! pagination types field for field in their JSON shape. [`transcode`] records
//! the source's `Serialize` output as a flat token tape and replays it into the
//! target's `Deserialize`, with exactly the semantics of
//! `serde_json::to_value` followed by an integral-number normalization and
//! `serde_json::from_value`:
//!
//! - a non-finite float becomes null;
//! - a finite float with no fractional part reads back as an integer;
//! - a map key written twice keeps its last value;
//! - unit variants are strings and other variants single-key objects.
//!
//! The tape allocates one token vector and one string buffer, both reused
//! across calls through [`Transcoder`], instead of one heap node per value.

use std::fmt;

use serde::de::value::StrDeserializer;
use serde::de::{self, DeserializeOwned, DeserializeSeed, Visitor};
use serde::ser::{self, Serialize};

/// Convert `value` into `U` as a JSON round trip would, without the JSON.
pub(crate) fn transcode<T: Serialize + ?Sized, U: DeserializeOwned>(
    value: &T,
) -> Result<U, String> {
    Transcoder::default().convert(value)
}

/// Reusable tape for converting many values in a row.
#[derive(Default)]
pub(crate) struct Transcoder {
    tape: Tape,
}

impl Transcoder {
    pub(crate) fn convert<T: Serialize + ?Sized, U: DeserializeOwned>(
        &mut self,
        value: &T,
    ) -> Result<U, String> {
        self.tape.clear();
        value.serialize(&mut self.tape).map_err(|error| error.0)?;
        let mut reader = Reader {
            tape: &self.tape,
            pos: 0,
        };
        let output = U::deserialize(&mut reader).map_err(|error| error.0)?;
        if reader.pos != self.tape.tokens.len() {
            return Err("trailing values after transcoded input".to_owned());
        }
        Ok(output)
    }
}

#[derive(Clone, Copy, Debug)]
enum Token {
    Null,
    Bool(bool),
    I64(i64),
    U64(u64),
    F64(f64),
    Str {
        start: u32,
        len: u32,
    },
    /// A field or variant name.
    Static(&'static str),
    /// `end` is the index just past the matching [`Token::End`].
    Seq {
        len: u32,
        end: u32,
    },
    Map {
        len: u32,
        end: u32,
    },
    End,
    /// A map entry overwritten by a later write of the same key; skip the
    /// next `n` tokens.
    Dead(u32),
}

#[derive(Default)]
struct Tape {
    tokens: Vec<Token>,
    text: String,
    open: Vec<Open>,
}

struct Open {
    token: usize,
    len: u32,
    /// Token index of each live entry's key, for maps.
    keys: Vec<usize>,
}

#[derive(Debug)]
pub(crate) struct Error(String);

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

impl de::Error for Error {
    fn custom<T: fmt::Display>(message: T) -> Self {
        Self(message.to_string())
    }
}

fn u32_len(value: usize, what: &str) -> Result<u32, Error> {
    u32::try_from(value).map_err(|_| Error(format!("transcoded {what} exceeds u32")))
}

impl Tape {
    fn clear(&mut self) {
        self.tokens.clear();
        self.text.clear();
        self.open.clear();
    }

    fn push(&mut self, token: Token) {
        self.tokens.push(token);
    }

    fn push_str(&mut self, value: &str) -> Result<(), Error> {
        let start = u32_len(self.text.len(), "string buffer")?;
        let len = u32_len(value.len(), "string")?;
        self.text.push_str(value);
        self.push(Token::Str { start, len });
        Ok(())
    }

    fn str_at(&self, index: usize) -> &str {
        match self.tokens[index] {
            Token::Str { start, len } => &self.text[start as usize..(start + len) as usize],
            Token::Static(value) => value,
            _ => "",
        }
    }

    fn open_seq(&mut self) {
        self.open.push(Open {
            token: self.tokens.len(),
            len: 0,
            keys: Vec::new(),
        });
        self.push(Token::Seq { len: 0, end: 0 });
    }

    fn open_map(&mut self) {
        self.open.push(Open {
            token: self.tokens.len(),
            len: 0,
            keys: Vec::new(),
        });
        self.push(Token::Map { len: 0, end: 0 });
    }

    fn element(&mut self) {
        if let Some(open) = self.open.last_mut() {
            open.len += 1;
        }
    }

    fn static_key(&mut self, key: &'static str) {
        let index = self.tokens.len();
        self.push(Token::Static(key));
        self.key(index);
    }

    /// Record a map key; an earlier entry with the same key is tombstoned so
    /// the last write wins, as it does when inserting into a JSON object.
    fn key(&mut self, key_token: usize) {
        let Some(open) = self.open.last() else {
            return;
        };
        let key = self.str_at(key_token);
        let earlier = open
            .keys
            .iter()
            .position(|&index| self.str_at(index) == key);
        let open = self.open.last_mut().expect("open container");
        if let Some(slot) = earlier {
            let earlier_token = open.keys.remove(slot);
            let next = open.keys.get(slot).copied().unwrap_or(key_token);
            self.tokens[earlier_token] = Token::Dead((next - earlier_token - 1) as u32);
        } else {
            open.len += 1;
        }
        open.keys.push(key_token);
    }

    fn close(&mut self) -> Result<(), Error> {
        let open = self
            .open
            .pop()
            .ok_or_else(|| Error("unbalanced transcoded container".to_owned()))?;
        self.push(Token::End);
        let end = u32_len(self.tokens.len(), "tape")?;
        self.tokens[open.token] = match self.tokens[open.token] {
            Token::Seq { .. } => Token::Seq { len: open.len, end },
            _ => Token::Map { len: open.len, end },
        };
        Ok(())
    }
}

// ---------------------------------------------------------------------------
// recording
// ---------------------------------------------------------------------------

impl<'a> ser::Serializer for &'a mut Tape {
    type Ok = ();
    type Error = Error;
    type SerializeSeq = Self;
    type SerializeTuple = Self;
    type SerializeTupleStruct = Self;
    type SerializeTupleVariant = Self;
    type SerializeMap = Self;
    type SerializeStruct = Self;
    type SerializeStructVariant = Self;

    fn serialize_bool(self, value: bool) -> Result<(), Error> {
        self.push(Token::Bool(value));
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
        self.push(Token::I64(value));
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
        self.push(Token::U64(value));
        Ok(())
    }
    fn serialize_i128(self, value: i128) -> Result<(), Error> {
        if let Ok(value) = i64::try_from(value) {
            self.serialize_i64(value)
        } else if let Ok(value) = u64::try_from(value) {
            self.serialize_u64(value)
        } else {
            Err(Error("number out of range".to_owned()))
        }
    }
    fn serialize_u128(self, value: u128) -> Result<(), Error> {
        u64::try_from(value)
            .map_err(|_| Error("number out of range".to_owned()))
            .and_then(|value| self.serialize_u64(value))
    }
    fn serialize_f32(self, value: f32) -> Result<(), Error> {
        self.serialize_f64(value.into())
    }
    fn serialize_f64(self, value: f64) -> Result<(), Error> {
        self.push(if value.is_finite() {
            Token::F64(value)
        } else {
            Token::Null
        });
        Ok(())
    }
    fn serialize_char(self, value: char) -> Result<(), Error> {
        self.push_str(value.encode_utf8(&mut [0; 4]))
    }
    fn serialize_str(self, value: &str) -> Result<(), Error> {
        self.push_str(value)
    }
    fn serialize_bytes(self, value: &[u8]) -> Result<(), Error> {
        self.open_seq();
        for byte in value {
            self.element();
            self.push(Token::U64((*byte).into()));
        }
        self.close()
    }
    fn serialize_none(self) -> Result<(), Error> {
        self.push(Token::Null);
        Ok(())
    }
    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<(), Error> {
        value.serialize(self)
    }
    fn serialize_unit(self) -> Result<(), Error> {
        self.push(Token::Null);
        Ok(())
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
        self.push(Token::Static(variant));
        Ok(())
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
        self.open_map();
        self.static_key(variant);
        value.serialize(&mut *self)?;
        self.close()
    }
    fn serialize_seq(self, _: Option<usize>) -> Result<Self, Error> {
        self.open_seq();
        Ok(self)
    }
    fn serialize_tuple(self, _: usize) -> Result<Self, Error> {
        self.open_seq();
        Ok(self)
    }
    fn serialize_tuple_struct(self, _: &'static str, _: usize) -> Result<Self, Error> {
        self.open_seq();
        Ok(self)
    }
    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        _: usize,
    ) -> Result<Self, Error> {
        self.open_map();
        self.static_key(variant);
        self.open_seq();
        Ok(self)
    }
    fn serialize_map(self, _: Option<usize>) -> Result<Self, Error> {
        self.open_map();
        Ok(self)
    }
    fn serialize_struct(self, _: &'static str, _: usize) -> Result<Self, Error> {
        self.open_map();
        Ok(self)
    }
    fn serialize_struct_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        _: usize,
    ) -> Result<Self, Error> {
        self.open_map();
        self.static_key(variant);
        self.open_map();
        Ok(self)
    }
}

impl ser::SerializeSeq for &mut Tape {
    type Ok = ();
    type Error = Error;
    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        self.element();
        value.serialize(&mut **self)
    }
    fn end(self) -> Result<(), Error> {
        self.close()
    }
}

impl ser::SerializeTuple for &mut Tape {
    type Ok = ();
    type Error = Error;
    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        ser::SerializeSeq::serialize_element(self, value)
    }
    fn end(self) -> Result<(), Error> {
        self.close()
    }
}

impl ser::SerializeTupleStruct for &mut Tape {
    type Ok = ();
    type Error = Error;
    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        ser::SerializeSeq::serialize_element(self, value)
    }
    fn end(self) -> Result<(), Error> {
        self.close()
    }
}

impl ser::SerializeTupleVariant for &mut Tape {
    type Ok = ();
    type Error = Error;
    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        ser::SerializeSeq::serialize_element(self, value)
    }
    fn end(self) -> Result<(), Error> {
        self.close()?;
        self.close()
    }
}

impl ser::SerializeMap for &mut Tape {
    type Ok = ();
    type Error = Error;
    fn serialize_key<T: Serialize + ?Sized>(&mut self, key: &T) -> Result<(), Error> {
        let index = self.tokens.len();
        key.serialize(MapKey { tape: self })?;
        self.key(index);
        Ok(())
    }
    fn serialize_value<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        value.serialize(&mut **self)
    }
    fn end(self) -> Result<(), Error> {
        self.close()
    }
}

impl ser::SerializeStruct for &mut Tape {
    type Ok = ();
    type Error = Error;
    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), Error> {
        self.static_key(key);
        value.serialize(&mut **self)
    }
    fn end(self) -> Result<(), Error> {
        self.close()
    }
}

impl ser::SerializeStructVariant for &mut Tape {
    type Ok = ();
    type Error = Error;
    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), Error> {
        self.static_key(key);
        value.serialize(&mut **self)
    }
    fn end(self) -> Result<(), Error> {
        self.close()?;
        self.close()
    }
}

/// Map keys become strings, as `serde_json::to_value` spells them.
struct MapKey<'a> {
    tape: &'a mut Tape,
}

fn key_must_be_a_string() -> Error {
    Error("key must be a string".to_owned())
}

impl ser::Serializer for MapKey<'_> {
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
        self.tape.push_str(value)
    }
    fn serialize_char(self, value: char) -> Result<(), Error> {
        self.tape.push_str(value.encode_utf8(&mut [0; 4]))
    }
    fn serialize_bool(self, value: bool) -> Result<(), Error> {
        self.tape.push_str(if value { "true" } else { "false" })
    }
    fn serialize_i8(self, value: i8) -> Result<(), Error> {
        self.tape.push_str(&value.to_string())
    }
    fn serialize_i16(self, value: i16) -> Result<(), Error> {
        self.tape.push_str(&value.to_string())
    }
    fn serialize_i32(self, value: i32) -> Result<(), Error> {
        self.tape.push_str(&value.to_string())
    }
    fn serialize_i64(self, value: i64) -> Result<(), Error> {
        self.tape.push_str(&value.to_string())
    }
    fn serialize_u8(self, value: u8) -> Result<(), Error> {
        self.tape.push_str(&value.to_string())
    }
    fn serialize_u16(self, value: u16) -> Result<(), Error> {
        self.tape.push_str(&value.to_string())
    }
    fn serialize_u32(self, value: u32) -> Result<(), Error> {
        self.tape.push_str(&value.to_string())
    }
    fn serialize_u64(self, value: u64) -> Result<(), Error> {
        self.tape.push_str(&value.to_string())
    }
    fn serialize_f32(self, value: f32) -> Result<(), Error> {
        self.serialize_f64(value.into())
    }
    fn serialize_f64(self, value: f64) -> Result<(), Error> {
        if !value.is_finite() {
            return Err(Error(
                "float key must be finite (got NaN or +/-inf)".to_owned(),
            ));
        }
        let key = serde_json::to_string(&value).map_err(|error| Error(error.to_string()))?;
        self.tape.push_str(&key)
    }
    fn serialize_bytes(self, _: &[u8]) -> Result<(), Error> {
        Err(key_must_be_a_string())
    }
    fn serialize_none(self) -> Result<(), Error> {
        Err(key_must_be_a_string())
    }
    fn serialize_some<T: Serialize + ?Sized>(self, _: &T) -> Result<(), Error> {
        Err(key_must_be_a_string())
    }
    fn serialize_unit(self) -> Result<(), Error> {
        Err(key_must_be_a_string())
    }
    fn serialize_unit_struct(self, _: &'static str) -> Result<(), Error> {
        Err(key_must_be_a_string())
    }
    fn serialize_unit_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
    ) -> Result<(), Error> {
        self.tape.push_str(variant)
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
        Err(key_must_be_a_string())
    }
    fn serialize_seq(self, _: Option<usize>) -> Result<Self::SerializeSeq, Error> {
        Err(key_must_be_a_string())
    }
    fn serialize_tuple(self, _: usize) -> Result<Self::SerializeTuple, Error> {
        Err(key_must_be_a_string())
    }
    fn serialize_tuple_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleStruct, Error> {
        Err(key_must_be_a_string())
    }
    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleVariant, Error> {
        Err(key_must_be_a_string())
    }
    fn serialize_map(self, _: Option<usize>) -> Result<Self::SerializeMap, Error> {
        Err(key_must_be_a_string())
    }
    fn serialize_struct(self, _: &'static str, _: usize) -> Result<Self::SerializeStruct, Error> {
        Err(key_must_be_a_string())
    }
    fn serialize_struct_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStructVariant, Error> {
        Err(key_must_be_a_string())
    }
}

// ---------------------------------------------------------------------------
// replay
// ---------------------------------------------------------------------------

struct Reader<'t> {
    tape: &'t Tape,
    pos: usize,
}

impl<'t> Reader<'t> {
    fn token(&self) -> Result<Token, Error> {
        self.tape
            .tokens
            .get(self.pos)
            .copied()
            .ok_or_else(|| Error("transcoded input ended early".to_owned()))
    }

    fn text(&self, start: u32, len: u32) -> &'t str {
        &self.tape.text[start as usize..(start + len) as usize]
    }

    /// The string token at the cursor, consumed.
    fn string(&mut self) -> Result<&'t str, Error> {
        let value = match self.token()? {
            Token::Str { start, len } => self.text(start, len),
            Token::Static(value) => value,
            _ => return Err(Error("transcoded map key is not a string".to_owned())),
        };
        self.pos += 1;
        Ok(value)
    }

    fn skip_dead(&mut self) {
        while let Some(Token::Dead(skip)) = self.tape.tokens.get(self.pos) {
            self.pos += 1 + *skip as usize;
        }
    }

    fn skip_value(&mut self) -> Result<(), Error> {
        match self.token()? {
            Token::Seq { end, .. } | Token::Map { end, .. } => self.pos = end as usize,
            Token::End | Token::Dead(_) => {
                return Err(Error("malformed transcoded input".to_owned()));
            }
            _ => self.pos += 1,
        }
        Ok(())
    }
}

/// A finite float with no fractional part reads back as an integer, the
/// normalization the display input applied to its JSON tree.
fn visit_number<'de, V: Visitor<'de>>(value: f64, visitor: V) -> Result<V::Value, Error> {
    if value.fract() == 0.0 && value >= i64::MIN as f64 && value <= i64::MAX as f64 {
        visit_signed(value as i64, visitor)
    } else {
        visitor.visit_f64(value)
    }
}

/// `serde_json::Number::from(i64)` stores non-negative values as unsigned.
fn visit_signed<'de, V: Visitor<'de>>(value: i64, visitor: V) -> Result<V::Value, Error> {
    if value < 0 {
        visitor.visit_i64(value)
    } else {
        visitor.visit_u64(value as u64)
    }
}

impl<'de> de::Deserializer<'de> for &mut Reader<'_> {
    type Error = Error;

    fn deserialize_any<V: Visitor<'de>>(self, visitor: V) -> Result<V::Value, Error> {
        match self.token()? {
            Token::Null => {
                self.pos += 1;
                visitor.visit_unit()
            }
            Token::Bool(value) => {
                self.pos += 1;
                visitor.visit_bool(value)
            }
            Token::I64(value) => {
                self.pos += 1;
                visit_signed(value, visitor)
            }
            Token::U64(value) => {
                self.pos += 1;
                visitor.visit_u64(value)
            }
            Token::F64(value) => {
                self.pos += 1;
                visit_number(value, visitor)
            }
            Token::Str { start, len } => {
                self.pos += 1;
                visitor.visit_string(self.text(start, len).to_owned())
            }
            Token::Static(value) => {
                self.pos += 1;
                visitor.visit_string(value.to_owned())
            }
            Token::Seq { len, end } => {
                self.pos += 1;
                let value = visitor.visit_seq(SeqReader {
                    reader: &mut *self,
                    remaining: len,
                })?;
                if self.pos + 1 != end as usize {
                    return Err(de::Error::invalid_length(
                        len as usize,
                        &"fewer elements in array",
                    ));
                }
                self.pos = end as usize;
                Ok(value)
            }
            Token::Map { len, end } => {
                self.pos += 1;
                let value = visitor.visit_map(MapReader {
                    reader: &mut *self,
                    remaining: len,
                })?;
                self.skip_dead();
                if self.pos + 1 != end as usize {
                    return Err(de::Error::invalid_length(
                        len as usize,
                        &"fewer elements in map",
                    ));
                }
                self.pos = end as usize;
                Ok(value)
            }
            Token::End | Token::Dead(_) => Err(Error("malformed transcoded input".to_owned())),
        }
    }

    fn deserialize_option<V: Visitor<'de>>(self, visitor: V) -> Result<V::Value, Error> {
        if let Token::Null = self.token()? {
            self.pos += 1;
            visitor.visit_none()
        } else {
            visitor.visit_some(self)
        }
    }

    fn deserialize_newtype_struct<V: Visitor<'de>>(
        self,
        _: &'static str,
        visitor: V,
    ) -> Result<V::Value, Error> {
        visitor.visit_newtype_struct(self)
    }

    fn deserialize_enum<V: Visitor<'de>>(
        self,
        _: &'static str,
        _: &'static [&'static str],
        visitor: V,
    ) -> Result<V::Value, Error> {
        let unit = match self.token()? {
            Token::Str { start, len } => Some(self.text(start, len)),
            Token::Static(value) => Some(value),
            _ => None,
        };
        if let Some(variant) = unit {
            self.pos += 1;
            return visitor.visit_enum(StrDeserializer::<Error>::new(variant));
        }
        match self.token()? {
            Token::Map { len: 1, end } => {
                self.pos += 1;
                self.skip_dead();
                let variant = self.string()?;
                let value = visitor.visit_enum(EnumReader {
                    reader: &mut *self,
                    variant,
                })?;
                if self.pos + 1 != end as usize {
                    return Err(Error("malformed transcoded enum".to_owned()));
                }
                self.pos = end as usize;
                Ok(value)
            }
            Token::Map { .. } => Err(de::Error::invalid_value(
                de::Unexpected::Map,
                &"map with a single key",
            )),
            _ => Err(de::Error::invalid_type(
                de::Unexpected::Other("non-enum value"),
                &"string or map",
            )),
        }
    }

    fn deserialize_ignored_any<V: Visitor<'de>>(self, visitor: V) -> Result<V::Value, Error> {
        self.skip_value()?;
        visitor.visit_unit()
    }

    serde::forward_to_deserialize_any! {
        bool i8 i16 i32 i64 i128 u8 u16 u32 u64 u128 f32 f64 char str string
        bytes byte_buf unit unit_struct seq tuple tuple_struct map struct
        identifier
    }
}

struct SeqReader<'r, 't> {
    reader: &'r mut Reader<'t>,
    remaining: u32,
}

impl<'de> de::SeqAccess<'de> for SeqReader<'_, '_> {
    type Error = Error;

    fn next_element_seed<T: DeserializeSeed<'de>>(
        &mut self,
        seed: T,
    ) -> Result<Option<T::Value>, Error> {
        if self.remaining == 0 {
            return Ok(None);
        }
        self.remaining -= 1;
        seed.deserialize(&mut *self.reader).map(Some)
    }

    fn size_hint(&self) -> Option<usize> {
        Some(self.remaining as usize)
    }
}

struct MapReader<'r, 't> {
    reader: &'r mut Reader<'t>,
    remaining: u32,
}

impl<'de> de::MapAccess<'de> for MapReader<'_, '_> {
    type Error = Error;

    fn next_key_seed<K: DeserializeSeed<'de>>(
        &mut self,
        seed: K,
    ) -> Result<Option<K::Value>, Error> {
        if self.remaining == 0 {
            return Ok(None);
        }
        self.remaining -= 1;
        self.reader.skip_dead();
        let key = self.reader.string()?;
        seed.deserialize(KeyReader { key }).map(Some)
    }

    fn next_value_seed<V: DeserializeSeed<'de>>(&mut self, seed: V) -> Result<V::Value, Error> {
        seed.deserialize(&mut *self.reader)
    }

    fn size_hint(&self) -> Option<usize> {
        Some(self.remaining as usize)
    }
}

struct EnumReader<'r, 't> {
    reader: &'r mut Reader<'t>,
    variant: &'t str,
}

impl<'de, 'r, 't> de::EnumAccess<'de> for EnumReader<'r, 't> {
    type Error = Error;
    type Variant = &'r mut Reader<'t>;

    fn variant_seed<V: DeserializeSeed<'de>>(
        self,
        seed: V,
    ) -> Result<(V::Value, Self::Variant), Error> {
        let variant = seed.deserialize(StrDeserializer::<Error>::new(self.variant))?;
        Ok((variant, self.reader))
    }
}

impl<'de> de::VariantAccess<'de> for &mut Reader<'_> {
    type Error = Error;

    fn unit_variant(self) -> Result<(), Error> {
        de::Deserialize::deserialize(self)
    }

    fn newtype_variant_seed<T: DeserializeSeed<'de>>(self, seed: T) -> Result<T::Value, Error> {
        seed.deserialize(self)
    }

    fn tuple_variant<V: Visitor<'de>>(self, _: usize, visitor: V) -> Result<V::Value, Error> {
        match self.token()? {
            Token::Seq { len: 0, end } => {
                self.pos = end as usize;
                visitor.visit_unit()
            }
            Token::Seq { .. } => de::Deserializer::deserialize_any(self, visitor),
            _ => Err(de::Error::invalid_type(
                de::Unexpected::Other("non-array value"),
                &"tuple variant",
            )),
        }
    }

    fn struct_variant<V: Visitor<'de>>(
        self,
        _: &'static [&'static str],
        visitor: V,
    ) -> Result<V::Value, Error> {
        match self.token()? {
            Token::Map { .. } => de::Deserializer::deserialize_any(self, visitor),
            _ => Err(de::Error::invalid_type(
                de::Unexpected::Other("non-object value"),
                &"struct variant",
            )),
        }
    }
}

/// Keys of a JSON object: strings that also parse as numbers or booleans when
/// the target asks for one, like `serde_json`'s map-key deserializer.
struct KeyReader<'a> {
    key: &'a str,
}

macro_rules! parse_key {
    ($($method:ident: $ty:ty => $visit:ident),* $(,)?) => {
        $(
            fn $method<V: Visitor<'de>>(self, visitor: V) -> Result<V::Value, Error> {
                visitor.$visit(self.numeric::<$ty>()?)
            }
        )*
    };
}

impl KeyReader<'_> {
    /// A numeric key reads with JSON number syntax, as `serde_json` reads it.
    fn numeric<T: DeserializeOwned>(&self) -> Result<T, Error> {
        if !self
            .key
            .as_bytes()
            .first()
            .is_some_and(|byte| byte.is_ascii_digit() || *byte == b'-')
        {
            return Err(Error(format!("invalid numeric map key {:?}", self.key)));
        }
        serde_json::from_str(self.key).map_err(|error| Error(error.to_string()))
    }
}

impl<'de> de::Deserializer<'de> for KeyReader<'_> {
    type Error = Error;

    fn deserialize_any<V: Visitor<'de>>(self, visitor: V) -> Result<V::Value, Error> {
        visitor.visit_str(self.key)
    }

    fn deserialize_bool<V: Visitor<'de>>(self, visitor: V) -> Result<V::Value, Error> {
        match self.key {
            "true" => visitor.visit_bool(true),
            "false" => visitor.visit_bool(false),
            _ => visitor.visit_str(self.key),
        }
    }

    parse_key! {
        deserialize_i8: i8 => visit_i8,
        deserialize_i16: i16 => visit_i16,
        deserialize_i32: i32 => visit_i32,
        deserialize_i64: i64 => visit_i64,
        deserialize_i128: i128 => visit_i128,
        deserialize_u8: u8 => visit_u8,
        deserialize_u16: u16 => visit_u16,
        deserialize_u32: u32 => visit_u32,
        deserialize_u64: u64 => visit_u64,
        deserialize_u128: u128 => visit_u128,
        deserialize_f32: f32 => visit_f32,
        deserialize_f64: f64 => visit_f64,
    }

    fn deserialize_option<V: Visitor<'de>>(self, visitor: V) -> Result<V::Value, Error> {
        visitor.visit_some(self)
    }

    fn deserialize_newtype_struct<V: Visitor<'de>>(
        self,
        _: &'static str,
        visitor: V,
    ) -> Result<V::Value, Error> {
        visitor.visit_newtype_struct(self)
    }

    fn deserialize_enum<V: Visitor<'de>>(
        self,
        _: &'static str,
        _: &'static [&'static str],
        visitor: V,
    ) -> Result<V::Value, Error> {
        visitor.visit_enum(StrDeserializer::<Error>::new(self.key))
    }

    serde::forward_to_deserialize_any! {
        char str string bytes byte_buf unit unit_struct seq
        tuple tuple_struct map struct identifier ignored_any
    }
}

#[cfg(test)]
mod tests {
    use serde::de::DeserializeOwned;
    use serde::{Deserialize, Serialize};
    use serde_json::{Number, Value, json};

    use super::transcode;

    fn normalize(value: &mut Value) {
        match value {
            Value::Array(values) => values.iter_mut().for_each(normalize),
            Value::Object(fields) => fields.values_mut().for_each(normalize),
            Value::Number(number) if !number.is_i64() && !number.is_u64() => {
                if let Some(float) = number.as_f64()
                    && float.fract() == 0.0
                    && float >= i64::MIN as f64
                    && float <= i64::MAX as f64
                {
                    *number = Number::from(float as i64);
                }
            }
            _ => {}
        }
    }

    fn via_json<T: Serialize, U: DeserializeOwned>(value: &T) -> Result<U, String> {
        let mut tree = serde_json::to_value(value).map_err(|error| error.to_string())?;
        normalize(&mut tree);
        serde_json::from_value(tree).map_err(|error| error.to_string())
    }

    fn assert_same_value<T: Serialize>(value: &T) {
        let expected: Value = via_json(value).expect("json round trip");
        let actual: Value = transcode(value).expect("transcode");
        assert_eq!(actual, expected);
    }

    #[derive(Serialize)]
    #[serde(rename_all = "camelCase")]
    struct Inner {
        size: f64,
        #[serde(skip_serializing_if = "Option::is_none")]
        label: Option<String>,
    }

    #[derive(Serialize)]
    #[serde(rename_all = "camelCase")]
    struct Outer {
        #[serde(flatten)]
        inner: Inner,
        size: f64,
        count: u32,
        ratio: f64,
        missing: Option<f64>,
        infinite: f64,
        tags: Vec<&'static str>,
        shape: Shape,
        kinds: Vec<Kind>,
        raw: Value,
    }

    #[derive(Serialize, Deserialize, Debug, PartialEq)]
    #[serde(tag = "kind", rename_all = "camelCase")]
    enum Shape {
        Rect { w: f64, h: f64 },
        Dot,
    }

    #[derive(Serialize, Deserialize, Debug, PartialEq)]
    enum Kind {
        Unit,
        Newtype(u32),
        Tuple(u8, String),
        Struct { at: i64 },
    }

    fn outer() -> Outer {
        Outer {
            inner: Inner {
                size: 1.5,
                label: Some("inner".to_owned()),
            },
            size: 12.0,
            count: 7,
            ratio: 0.25,
            missing: None,
            infinite: f64::INFINITY,
            tags: vec!["a", "b"],
            shape: Shape::Rect { w: 3.0, h: -0.0 },
            kinds: vec![
                Kind::Unit,
                Kind::Newtype(4),
                Kind::Tuple(1, "x".to_owned()),
                Kind::Struct { at: -2 },
            ],
            raw: json!({ "z": [1.0, 2.5, -3.0, null, true], "a": { "b": "c" } }),
        }
    }

    #[test]
    fn matches_the_json_round_trip_as_a_value() {
        assert_same_value(&outer());
        assert_same_value(&json!([[], {}, "", 0.0, 1e300, -1e300, u64::MAX, i64::MIN]));
    }

    #[derive(Deserialize, Debug, PartialEq)]
    #[serde(rename_all = "camelCase")]
    struct Target {
        size: f64,
        count: u64,
        ratio: f64,
        missing: Option<f64>,
        infinite: Option<f64>,
        label: String,
        shape: Shape,
        kinds: Vec<Kind>,
        #[serde(default)]
        absent: Vec<u32>,
        raw: Value,
    }

    #[test]
    fn matches_the_json_round_trip_into_typed_targets() {
        let expected: Target = via_json(&outer()).expect("json round trip");
        let actual: Target = transcode(&outer()).expect("transcode");
        assert_eq!(actual, expected);
        // The later `size` write wins over the flattened one.
        assert_eq!(actual.size, 12.0);
    }

    #[test]
    fn integral_floats_read_back_as_integers() {
        let count: u32 = transcode(&3.0_f64).unwrap();
        assert_eq!(count, 3);
        assert!(transcode::<_, u32>(&3.5_f64).is_err());
        assert_eq!(transcode::<_, Value>(&-0.0_f64).unwrap(), json!(0));
    }

    fn assert_same_outcome<T: Serialize, U: DeserializeOwned + PartialEq + std::fmt::Debug>(
        value: &T,
    ) {
        let expected = via_json::<T, U>(value);
        let actual = transcode::<T, U>(value);
        match (&expected, &actual) {
            (Ok(expected), Ok(actual)) => assert_eq!(actual, expected),
            (Err(_), Err(_)) => {}
            _ => panic!("json round trip {expected:?} but transcoded {actual:?}"),
        }
    }

    #[derive(Serialize)]
    struct Renamed {
        #[serde(rename = "x")]
        first: u32,
        #[serde(rename = "x")]
        second: u32,
    }

    #[test]
    fn rejects_and_accepts_what_the_json_round_trip_does() {
        assert_same_outcome::<_, Kind>(&json!({ "Struct": [1] }));
        assert_same_outcome::<_, Kind>(&json!({ "Tuple": [] }));
        assert_same_outcome::<_, Kind>(&json!({ "Tuple": [1, "a"] }));
        assert_same_outcome::<_, Kind>(&json!({ "Struct": { "at": 3.0 } }));
        assert_same_outcome::<_, Kind>(&json!("Newtype"));
        for key in ["+1", "01", "1", "-2", "1.5", "x"] {
            assert_same_outcome::<_, std::collections::BTreeMap<i32, u8>>(&json!({ key: 1 }));
        }
        assert_same_outcome::<_, std::collections::BTreeMap<String, u8>>(
            &[(1.5_f64, 2_u8)]
                .into_iter()
                .map(|(key, value)| (key.to_string(), value))
                .collect::<std::collections::BTreeMap<_, _>>(),
        );
        assert_same_outcome::<_, Value>(&Renamed {
            first: 1,
            second: 2,
        });
        assert_same_outcome::<_, Value>(&7_i128);
        assert_same_outcome::<_, Value>(&u128::MAX);
    }

    #[test]
    fn map_keys_round_trip_as_strings() {
        let map: std::collections::BTreeMap<u32, &str> = [(1, "a"), (20, "b")].into();
        assert_same_value(&map);
        let back: std::collections::BTreeMap<u32, String> = transcode(&map).unwrap();
        assert_eq!(back[&20], "b");
    }
}
