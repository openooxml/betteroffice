//! Streaming content fingerprints for typed layout values.
//!
//! [`fingerprint_without_positions`] walks a value through its `Serialize`
//! impl and feeds every scalar to two independently seeded 64-bit foldhash
//! lanes, a 128-bit fingerprint, so no intermediate JSON tree is built. Object
//! keys named `pmStart`, `pmEnd`, `docStart` and `docEnd` are skipped at every
//! depth: those absolute document positions shift when text is edited earlier
//! in the story without changing what the value measures or paints.
//! Fingerprints are only compared against others from the same session.

use std::fmt;
use std::hash::Hasher as _;

use foldhash::quality::FoldHasher;
use serde::Serialize;
use serde::ser::{self, Serializer};

pub(crate) type Fingerprint = u128;

struct Seeds {
    shared: [foldhash::SharedSeed; 2],
    per_hasher: [u64; 2],
}

fn seeds() -> &'static Seeds {
    static SEEDS: std::sync::OnceLock<Seeds> = std::sync::OnceLock::new();
    SEEDS.get_or_init(|| {
        let [a, b, c, d] = seed_words();
        Seeds {
            shared: [a, b].map(foldhash::SharedSeed::from_u64),
            per_hasher: [c, d],
        }
    })
}

#[cfg(not(all(
    target_family = "wasm",
    target_os = "unknown",
    not(all(feature = "wasm", target_arch = "wasm32"))
)))]
fn seed_words() -> [u64; 4] {
    let [a, b] = crate::identity::entropy();
    let [c, d] = crate::identity::entropy();
    [a, b, c, d]
}

/// Without the `wasm` feature, `wasm32-unknown-unknown` has no entropy source.
#[cfg(all(
    target_family = "wasm",
    target_os = "unknown",
    not(all(feature = "wasm", target_arch = "wasm32"))
))]
fn seed_words() -> [u64; 4] {
    [
        0x243f_6a88_85a3_08d3,
        0x1319_8a2e_0370_7344,
        0xa409_3822_299f_31d0,
        0x082e_fa98_ec4e_6c89,
    ]
}

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
) -> Result<Fingerprint, String> {
    let mut hasher = Hasher::new();
    value.serialize(&mut hasher).map_err(|error| error.0)?;
    Ok(hasher.finish())
}

/// Fingerprint of `value` with its absolute document positions included, so it
/// also tells where each part of the value sits.
pub(crate) fn fingerprint_with_positions<T: Serialize + ?Sized>(
    value: &T,
) -> Result<Fingerprint, String> {
    let mut hasher = Hasher::new();
    hasher.keep_positions = true;
    value.serialize(&mut hasher).map_err(|error| error.0)?;
    Ok(hasher.finish())
}

fn is_position_key(key: &str) -> bool {
    POSITION_KEYS.contains(&key)
}

/// Two independently seeded 64-bit foldhash lanes. The seeds are drawn once per
/// process: deterministic within it, which is all a session-local fingerprint
/// needs, and unknown to the content being hashed.
struct Hasher {
    a: FoldHasher<'static>,
    b: FoldHasher<'static>,
    keep_positions: bool,
}

impl Hasher {
    fn new() -> Self {
        let seeds = seeds();
        Self {
            a: FoldHasher::with_seed(seeds.per_hasher[0], &seeds.shared[0]),
            b: FoldHasher::with_seed(seeds.per_hasher[1], &seeds.shared[1]),
            keep_positions: false,
        }
    }

    fn word(&mut self, word: u64) {
        self.a.write_u64(word);
        self.b.write_u64(word);
    }

    fn bytes(&mut self, bytes: &[u8]) {
        self.word(bytes.len() as u64);
        self.a.write(bytes);
        self.b.write(bytes);
    }

    fn finish(&self) -> Fingerprint {
        (self.a.finish() as u128) << 64 | self.b.finish() as u128
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
    position: bool,
}

impl KeyProbe {
    fn new() -> Self {
        Self {
            key: Hasher::new(),
            position: false,
        }
    }
}

struct Container<'a> {
    hasher: &'a mut Hasher,
    skip_value: bool,
}

impl<'a> Container<'a> {
    fn field<T: Serialize + ?Sized>(&mut self, key: &str, value: &T) -> Result<(), Error> {
        if !self.hasher.keep_positions && is_position_key(key) {
            return Ok(());
        }
        self.hasher.word(TAG_KEY);
        self.hasher.bytes(key.as_bytes());
        value.serialize(&mut *self.hasher)
    }

    fn key<T: Serialize + ?Sized>(&mut self, key: &T) -> Result<(), Error> {
        let mut probe = KeyProbe::new();
        key.serialize(&mut probe)?;
        let skip = probe.position && !self.hasher.keep_positions;
        self.skip_value = skip;
        if !skip {
            self.hasher.word(TAG_KEY);
            let key = probe.key.finish();
            self.hasher.word(key as u64);
            self.hasher.word((key >> 64) as u64);
        }
        Ok(())
    }

    fn value<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        if std::mem::take(&mut self.skip_value) {
            return Ok(());
        }
        value.serialize(&mut *self.hasher)
    }

    fn end(self) -> Result<(), Error> {
        self.hasher.word(TAG_END);
        Ok(())
    }
}

impl<'a> Serializer for &'a mut Hasher {
    type Ok = ();
    type Error = Error;
    type SerializeSeq = Container<'a>;
    type SerializeTuple = Container<'a>;
    type SerializeTupleStruct = Container<'a>;
    type SerializeTupleVariant = Container<'a>;
    type SerializeMap = Container<'a>;
    type SerializeStruct = Container<'a>;
    type SerializeStructVariant = Container<'a>;

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
        self.word(TAG_U64);
        self.word(value);
        Ok(())
    }
    fn serialize_f32(self, value: f32) -> Result<(), Error> {
        self.serialize_f64(value.into())
    }
    fn serialize_f64(self, value: f64) -> Result<(), Error> {
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
        self.word(TAG_NULL);
        Ok(())
    }
    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<(), Error> {
        value.serialize(self)
    }
    fn serialize_unit(self) -> Result<(), Error> {
        self.word(TAG_NULL);
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
        };
        container.field(variant, value)?;
        container.end()
    }
    fn serialize_seq(self, _: Option<usize>) -> Result<Container<'a>, Error> {
        self.word(TAG_ARRAY);
        Ok(Container {
            hasher: self,
            skip_value: false,
        })
    }
    fn serialize_tuple(self, len: usize) -> Result<Container<'a>, Error> {
        self.serialize_seq(Some(len))
    }
    fn serialize_tuple_struct(self, _: &'static str, len: usize) -> Result<Container<'a>, Error> {
        self.serialize_seq(Some(len))
    }
    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        len: usize,
    ) -> Result<Container<'a>, Error> {
        self.word(TAG_OBJECT);
        self.word(TAG_KEY);
        self.bytes(variant.as_bytes());
        self.serialize_seq(Some(len))
    }
    fn serialize_map(self, _: Option<usize>) -> Result<Container<'a>, Error> {
        self.word(TAG_OBJECT);
        Ok(Container {
            hasher: self,
            skip_value: false,
        })
    }
    fn serialize_struct(self, _: &'static str, len: usize) -> Result<Container<'a>, Error> {
        self.serialize_map(Some(len))
    }
    fn serialize_struct_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        len: usize,
    ) -> Result<Container<'a>, Error> {
        self.word(TAG_OBJECT);
        self.word(TAG_KEY);
        self.bytes(variant.as_bytes());
        self.serialize_map(Some(len))
    }
}

impl ser::SerializeSeq for Container<'_> {
    type Ok = ();
    type Error = Error;
    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        value.serialize(&mut *self.hasher)
    }
    fn end(self) -> Result<(), Error> {
        Container::end(self)
    }
}

impl ser::SerializeTuple for Container<'_> {
    type Ok = ();
    type Error = Error;
    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        value.serialize(&mut *self.hasher)
    }
    fn end(self) -> Result<(), Error> {
        Container::end(self)
    }
}

impl ser::SerializeTupleStruct for Container<'_> {
    type Ok = ();
    type Error = Error;
    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        value.serialize(&mut *self.hasher)
    }
    fn end(self) -> Result<(), Error> {
        Container::end(self)
    }
}

impl ser::SerializeTupleVariant for Container<'_> {
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

impl ser::SerializeMap for Container<'_> {
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

impl ser::SerializeStruct for Container<'_> {
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

impl ser::SerializeStructVariant for Container<'_> {
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
        self.position = is_position_key(value);
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

    use super::fingerprint_without_positions as fingerprint;

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
    fn a_positioned_fingerprint_tells_where_each_part_sits() {
        use super::fingerprint_with_positions as positioned;
        let base = positioned(&run("text", 1.0)).unwrap();
        assert_eq!(base, positioned(&run("text", 1.0)).unwrap());
        assert_ne!(base, positioned(&run("text", 2.0)).unwrap());
        // The same position owned by another part of the value.
        let owner = |first: Option<f64>, second: Option<f64>| json!([{ "text": "a", "pmStart": first }, { "text": "b", "pmStart": second }]);
        assert_ne!(
            positioned(&owner(Some(101.0), None)).unwrap(),
            positioned(&owner(None, Some(101.0))).unwrap()
        );
        assert_eq!(
            fingerprint(&owner(Some(101.0), None)).unwrap(),
            fingerprint(&owner(None, Some(101.0))).unwrap()
        );
    }

    #[test]
    fn absolute_positions_do_not_change_the_fingerprint() {
        assert_eq!(
            fingerprint(&run("text", 1.0)),
            fingerprint(&run("text", 90.0))
        );
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
