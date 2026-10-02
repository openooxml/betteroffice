//! Direct serialization into ordered JSON values.

use std::cell::Cell;
use std::fmt;

use serde::Serialize;
use serde::ser::{
    Impossible, SerializeMap, SerializeSeq, SerializeStruct, SerializeStructVariant,
    SerializeTuple, SerializeTupleStruct, SerializeTupleVariant,
};

use super::{OrderedValue, SOURCE_ORDINAL};

const MAX_DIRECT_DEPTH: u32 = 100;
const MAX_PREALLOCATED_BYTES: usize = 1024 * 1024;

thread_local! {
    static FAILED: Cell<bool> = const { Cell::new(false) };
}

/// Builds an ordered JSON value; falls back (`None`) on any serializer error.
pub(super) fn ordered_value<T: Serialize + ?Sized>(value: &T) -> Option<OrderedValue> {
    let outer = FAILED.replace(false);
    let ordered = value.serialize(ValueSerializer { depth: 0 }).ok();
    if FAILED.replace(outer) { None } else { ordered }
}

#[derive(Debug)]
struct Error(());

impl Error {
    fn new() -> Self {
        FAILED.set(true);
        Self(())
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("unsupported ordered JSON serialization")
    }
}

impl std::error::Error for Error {}

impl serde::ser::Error for Error {
    fn custom<T: fmt::Display>(_message: T) -> Self {
        Self::new()
    }
}

fn float_value<T: Serialize>(value: T) -> Result<OrderedValue, Error> {
    let token = serde_json::to_string(&value).map_err(|_| Error::new())?;
    serde_json::from_str(&token).map_err(|_| Error::new())
}

fn preallocated<T>(len: usize) -> Vec<T> {
    Vec::with_capacity(len.min(MAX_PREALLOCATED_BYTES / std::mem::size_of::<T>().max(1)))
}

fn variant_value(variant: &'static str, value: OrderedValue) -> OrderedValue {
    if variant == SOURCE_ORDINAL {
        OrderedValue::Object(Vec::new())
    } else {
        OrderedValue::Object(vec![(variant.to_owned(), value)])
    }
}

#[derive(Clone, Copy)]
struct ValueSerializer {
    depth: u32,
}

impl ValueSerializer {
    fn nested(self, levels: u32) -> Result<Self, Error> {
        let depth = self.depth + levels;
        if depth > MAX_DIRECT_DEPTH {
            return Err(Error::new());
        }
        Ok(Self { depth })
    }
}

impl serde::Serializer for ValueSerializer {
    type Ok = OrderedValue;
    type Error = Error;
    type SerializeSeq = ArraySerializer;
    type SerializeTuple = ArraySerializer;
    type SerializeTupleStruct = ArraySerializer;
    type SerializeTupleVariant = ArraySerializer;
    type SerializeMap = ObjectSerializer;
    type SerializeStruct = ObjectSerializer;
    type SerializeStructVariant = ObjectSerializer;

    fn serialize_bool(self, value: bool) -> Result<Self::Ok, Error> {
        Ok(OrderedValue::Bool(value))
    }

    fn serialize_i8(self, value: i8) -> Result<Self::Ok, Error> {
        self.serialize_i64(value.into())
    }

    fn serialize_i16(self, value: i16) -> Result<Self::Ok, Error> {
        self.serialize_i64(value.into())
    }

    fn serialize_i32(self, value: i32) -> Result<Self::Ok, Error> {
        self.serialize_i64(value.into())
    }

    fn serialize_i64(self, value: i64) -> Result<Self::Ok, Error> {
        if value >= 0 {
            self.serialize_u64(value as u64)
        } else {
            Ok(OrderedValue::Number(value.into()))
        }
    }

    fn serialize_i128(self, value: i128) -> Result<Self::Ok, Error> {
        if value >= 0 {
            self.serialize_u64(value.try_into().map_err(|_| Error::new())?)
        } else {
            self.serialize_i64(value.try_into().map_err(|_| Error::new())?)
        }
    }

    fn serialize_u8(self, value: u8) -> Result<Self::Ok, Error> {
        self.serialize_u64(value.into())
    }

    fn serialize_u16(self, value: u16) -> Result<Self::Ok, Error> {
        self.serialize_u64(value.into())
    }

    fn serialize_u32(self, value: u32) -> Result<Self::Ok, Error> {
        self.serialize_u64(value.into())
    }

    fn serialize_u64(self, value: u64) -> Result<Self::Ok, Error> {
        Ok(OrderedValue::Number(value.into()))
    }

    fn serialize_u128(self, value: u128) -> Result<Self::Ok, Error> {
        self.serialize_u64(value.try_into().map_err(|_| Error::new())?)
    }

    fn serialize_f32(self, value: f32) -> Result<Self::Ok, Error> {
        float_value(value)
    }

    fn serialize_f64(self, value: f64) -> Result<Self::Ok, Error> {
        float_value(value)
    }

    fn serialize_char(self, value: char) -> Result<Self::Ok, Error> {
        Ok(OrderedValue::String(value.to_string()))
    }

    fn serialize_str(self, value: &str) -> Result<Self::Ok, Error> {
        Ok(OrderedValue::String(value.to_owned()))
    }

    fn serialize_bytes(self, value: &[u8]) -> Result<Self::Ok, Error> {
        self.nested(1)?;
        Ok(OrderedValue::Array(
            value
                .iter()
                .map(|byte| OrderedValue::Number(u64::from(*byte).into()))
                .collect(),
        ))
    }

    fn serialize_none(self) -> Result<Self::Ok, Error> {
        Ok(OrderedValue::Null)
    }

    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<Self::Ok, Error> {
        value.serialize(self)
    }

    fn serialize_unit(self) -> Result<Self::Ok, Error> {
        Ok(OrderedValue::Null)
    }

    fn serialize_unit_struct(self, _name: &'static str) -> Result<Self::Ok, Error> {
        Ok(OrderedValue::Null)
    }

    fn serialize_unit_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
    ) -> Result<Self::Ok, Error> {
        self.serialize_str(variant)
    }

    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _name: &'static str,
        value: &T,
    ) -> Result<Self::Ok, Error> {
        value.serialize(self)
    }

    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
        value: &T,
    ) -> Result<Self::Ok, Error> {
        Ok(variant_value(variant, value.serialize(self.nested(1)?)?))
    }

    fn serialize_seq(self, len: Option<usize>) -> Result<ArraySerializer, Error> {
        Ok(ArraySerializer::new(len, None, self.nested(1)?.depth))
    }

    fn serialize_tuple(self, len: usize) -> Result<ArraySerializer, Error> {
        Ok(ArraySerializer::new(Some(len), None, self.nested(1)?.depth))
    }

    fn serialize_tuple_struct(
        self,
        _name: &'static str,
        len: usize,
    ) -> Result<ArraySerializer, Error> {
        Ok(ArraySerializer::new(Some(len), None, self.nested(1)?.depth))
    }

    fn serialize_tuple_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
        len: usize,
    ) -> Result<ArraySerializer, Error> {
        Ok(ArraySerializer::new(
            Some(len),
            Some(variant),
            self.nested(2)?.depth,
        ))
    }

    fn serialize_map(self, len: Option<usize>) -> Result<ObjectSerializer, Error> {
        Ok(ObjectSerializer::new(len, None, self.nested(1)?.depth))
    }

    fn serialize_struct(self, name: &'static str, len: usize) -> Result<ObjectSerializer, Error> {
        if name.starts_with("$serde_json::private::") {
            return Err(Error::new());
        }
        Ok(ObjectSerializer::new(
            Some(len),
            None,
            self.nested(1)?.depth,
        ))
    }

    fn serialize_struct_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
        len: usize,
    ) -> Result<ObjectSerializer, Error> {
        Ok(ObjectSerializer::new(
            Some(len),
            Some(variant),
            self.nested(2)?.depth,
        ))
    }

    fn is_human_readable(&self) -> bool {
        true
    }
}

struct ArraySerializer {
    values: Vec<OrderedValue>,
    empty: bool,
    variant: Option<&'static str>,
    depth: u32,
}

impl ArraySerializer {
    fn new(len: Option<usize>, variant: Option<&'static str>, depth: u32) -> Self {
        Self {
            values: preallocated(len.unwrap_or(0)),
            empty: len == Some(0),
            variant,
            depth,
        }
    }

    fn push<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        if self.empty {
            return Err(Error::new());
        }
        self.values
            .push(value.serialize(ValueSerializer { depth: self.depth })?);
        Ok(())
    }

    fn finish(self) -> OrderedValue {
        let value = OrderedValue::Array(self.values);
        match self.variant {
            Some(variant) => variant_value(variant, value),
            None => value,
        }
    }
}

impl SerializeSeq for ArraySerializer {
    type Ok = OrderedValue;
    type Error = Error;

    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        self.push(value)
    }

    fn end(self) -> Result<Self::Ok, Error> {
        Ok(self.finish())
    }
}

impl SerializeTuple for ArraySerializer {
    type Ok = OrderedValue;
    type Error = Error;

    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        self.push(value)
    }

    fn end(self) -> Result<Self::Ok, Error> {
        Ok(self.finish())
    }
}

impl SerializeTupleStruct for ArraySerializer {
    type Ok = OrderedValue;
    type Error = Error;

    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        self.push(value)
    }

    fn end(self) -> Result<Self::Ok, Error> {
        Ok(self.finish())
    }
}

impl SerializeTupleVariant for ArraySerializer {
    type Ok = OrderedValue;
    type Error = Error;

    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        self.push(value)
    }

    fn end(self) -> Result<Self::Ok, Error> {
        Ok(self.finish())
    }
}

struct ObjectSerializer {
    entries: Vec<(String, OrderedValue)>,
    key: Option<String>,
    empty: bool,
    variant: Option<&'static str>,
    depth: u32,
}

impl ObjectSerializer {
    fn new(len: Option<usize>, variant: Option<&'static str>, depth: u32) -> Self {
        Self {
            entries: preallocated(len.unwrap_or(0)),
            key: None,
            empty: len == Some(0),
            variant,
            depth,
        }
    }

    fn push<T: Serialize + ?Sized>(&mut self, key: String, value: &T) -> Result<(), Error> {
        if self.empty {
            return Err(Error::new());
        }
        let value = value.serialize(ValueSerializer { depth: self.depth })?;
        if key != SOURCE_ORDINAL {
            self.entries.push((key, value));
        }
        Ok(())
    }

    fn finish(self) -> Result<OrderedValue, Error> {
        if self.key.is_some() {
            return Err(Error::new());
        }
        let value = OrderedValue::Object(self.entries);
        Ok(match self.variant {
            Some(variant) => variant_value(variant, value),
            None => value,
        })
    }
}

impl SerializeMap for ObjectSerializer {
    type Ok = OrderedValue;
    type Error = Error;

    fn serialize_key<T: Serialize + ?Sized>(&mut self, key: &T) -> Result<(), Error> {
        if self.empty || self.key.is_some() {
            return Err(Error::new());
        }
        self.key = Some(key.serialize(KeySerializer)?);
        Ok(())
    }

    fn serialize_value<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        let key = self.key.take().ok_or_else(Error::new)?;
        self.push(key, value)
    }

    fn end(self) -> Result<Self::Ok, Error> {
        self.finish()
    }
}

impl SerializeStruct for ObjectSerializer {
    type Ok = OrderedValue;
    type Error = Error;

    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), Error> {
        self.push(key.to_owned(), value)
    }

    fn end(self) -> Result<Self::Ok, Error> {
        self.finish()
    }
}

impl SerializeStructVariant for ObjectSerializer {
    type Ok = OrderedValue;
    type Error = Error;

    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), Error> {
        self.push(key.to_owned(), value)
    }

    fn end(self) -> Result<Self::Ok, Error> {
        self.finish()
    }
}

struct KeySerializer;

impl serde::Serializer for KeySerializer {
    type Ok = String;
    type Error = Error;
    type SerializeSeq = Impossible<String, Error>;
    type SerializeTuple = Impossible<String, Error>;
    type SerializeTupleStruct = Impossible<String, Error>;
    type SerializeTupleVariant = Impossible<String, Error>;
    type SerializeMap = Impossible<String, Error>;
    type SerializeStruct = Impossible<String, Error>;
    type SerializeStructVariant = Impossible<String, Error>;

    fn serialize_bool(self, _value: bool) -> Result<String, Error> {
        Err(Error::new())
    }

    fn serialize_i8(self, value: i8) -> Result<String, Error> {
        Ok(value.to_string())
    }

    fn serialize_i16(self, value: i16) -> Result<String, Error> {
        Ok(value.to_string())
    }

    fn serialize_i32(self, value: i32) -> Result<String, Error> {
        Ok(value.to_string())
    }

    fn serialize_i64(self, value: i64) -> Result<String, Error> {
        Ok(value.to_string())
    }

    fn serialize_i128(self, value: i128) -> Result<String, Error> {
        Ok(value.to_string())
    }

    fn serialize_u8(self, value: u8) -> Result<String, Error> {
        Ok(value.to_string())
    }

    fn serialize_u16(self, value: u16) -> Result<String, Error> {
        Ok(value.to_string())
    }

    fn serialize_u32(self, value: u32) -> Result<String, Error> {
        Ok(value.to_string())
    }

    fn serialize_u64(self, value: u64) -> Result<String, Error> {
        Ok(value.to_string())
    }

    fn serialize_u128(self, value: u128) -> Result<String, Error> {
        Ok(value.to_string())
    }

    fn serialize_f32(self, _value: f32) -> Result<String, Error> {
        Err(Error::new())
    }

    fn serialize_f64(self, _value: f64) -> Result<String, Error> {
        Err(Error::new())
    }

    fn serialize_char(self, value: char) -> Result<String, Error> {
        Ok(value.to_string())
    }

    fn serialize_str(self, value: &str) -> Result<String, Error> {
        Ok(value.to_owned())
    }

    fn serialize_bytes(self, _value: &[u8]) -> Result<String, Error> {
        Err(Error::new())
    }

    fn serialize_none(self) -> Result<String, Error> {
        Err(Error::new())
    }

    fn serialize_some<T: Serialize + ?Sized>(self, _value: &T) -> Result<String, Error> {
        Err(Error::new())
    }

    fn serialize_unit(self) -> Result<String, Error> {
        Err(Error::new())
    }

    fn serialize_unit_struct(self, _name: &'static str) -> Result<String, Error> {
        Err(Error::new())
    }

    fn serialize_unit_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        variant: &'static str,
    ) -> Result<String, Error> {
        Ok(variant.to_owned())
    }

    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _name: &'static str,
        value: &T,
    ) -> Result<String, Error> {
        value.serialize(self)
    }

    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _name: &'static str,
        _variant_index: u32,
        _variant: &'static str,
        _value: &T,
    ) -> Result<String, Error> {
        Err(Error::new())
    }

    fn serialize_seq(self, _len: Option<usize>) -> Result<Self::SerializeSeq, Error> {
        Err(Error::new())
    }

    fn serialize_tuple(self, _len: usize) -> Result<Self::SerializeTuple, Error> {
        Err(Error::new())
    }

    fn serialize_tuple_struct(
        self,
        _name: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeTupleStruct, Error> {
        Err(Error::new())
    }

    fn serialize_tuple_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        _variant: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeTupleVariant, Error> {
        Err(Error::new())
    }

    fn serialize_map(self, _len: Option<usize>) -> Result<Self::SerializeMap, Error> {
        Err(Error::new())
    }

    fn serialize_struct(
        self,
        _name: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeStruct, Error> {
        Err(Error::new())
    }

    fn serialize_struct_variant(
        self,
        _name: &'static str,
        _variant_index: u32,
        _variant: &'static str,
        _len: usize,
    ) -> Result<Self::SerializeStructVariant, Error> {
        Err(Error::new())
    }

    fn is_human_readable(&self) -> bool {
        true
    }
}

#[cfg(test)]
mod tests {
    use std::collections::{BTreeMap, HashMap};

    use serde::Serialize;
    use serde::ser::{SerializeMap, SerializeSeq, SerializeStruct};
    use serde_json::{Map, Value, json};

    use super::super::source_ordered_value;
    use super::{OrderedValue, SOURCE_ORDINAL, ordered_value};

    const DEPTHS: [u32; 7] = [99, 100, 101, 127, 128, 129, 200];

    fn round_trip<T: Serialize + ?Sized>(value: &T) -> Result<OrderedValue, String> {
        let token = serde_json::to_string(value).map_err(|error| error.to_string())?;
        serde_json::from_str(&token).map_err(|error| error.to_string())
    }

    fn assert_source_round_trip<T: Serialize + ?Sized>(value: &T) {
        assert_eq!(source_ordered_value(value), round_trip(value));
    }

    fn assert_depth_round_trip<T: Serialize + ?Sized>(value: &T, depth: u32) {
        let expected = round_trip(value);
        assert_eq!(expected.is_ok(), depth < 128, "JSON depth {depth}");
        assert_eq!(source_ordered_value(value), expected, "JSON depth {depth}");
    }

    fn comparable(value: &OrderedValue) -> Value {
        match value {
            OrderedValue::Null => json!(["null"]),
            OrderedValue::Bool(value) => json!(["bool", value]),
            OrderedValue::Number(value) => json!(["number", value.to_string()]),
            OrderedValue::String(value) => json!(["string", value]),
            OrderedValue::Array(values) => {
                json!(["array", values.iter().map(comparable).collect::<Vec<_>>()])
            }
            OrderedValue::Object(entries) => json!([
                "object",
                entries
                    .iter()
                    .map(|(key, value)| (key, comparable(value)))
                    .collect::<Vec<_>>()
            ]),
        }
    }

    fn assert_round_trip<T: Serialize + ?Sized>(value: &T) -> OrderedValue {
        let direct = ordered_value(value).expect("direct serialization is supported");
        let round_trip = round_trip(value).unwrap();
        assert_eq!(comparable(&direct), comparable(&round_trip));
        direct
    }

    #[test]
    fn json_value_matches_round_trip() {
        let value = json!({
            "z": [null, true, false, -5, 5, 0.1, 1e21, -0.0, u64::MAX],
            "sourceOrdinal": 17,
            "a": {
                "sourceOrdinal": 23,
                "text": "quotes: \"; slash: \\; newline: \n; unicode: 🦀",
                "nested": [{"sourceOrdinal": 42, "value": "kept"}]
            }
        });
        assert_round_trip(&value);
    }

    #[allow(non_snake_case)]
    #[derive(Serialize)]
    struct Nested {
        z: bool,
        sourceOrdinal: u32,
        a: Option<String>,
    }

    #[allow(non_snake_case)]
    #[derive(Serialize)]
    enum Variant {
        Unit,
        Newtype(Nested),
        Tuple(i32, Option<String>),
        Struct {
            z: u64,
            sourceOrdinal: u32,
            a: Nested,
        },
        #[serde(rename = "sourceOrdinal")]
        Ordinal(Box<Variant>),
    }

    #[derive(Serialize)]
    struct Sample {
        nested: Nested,
        some: Option<Nested>,
        none: Option<String>,
        floats32: [f32; 2],
        floats64: [f64; 6],
        negative: i32,
        positive: i32,
        maximum: u64,
        integer_keys: BTreeMap<u32, String>,
        string_keys: HashMap<String, i32>,
        variants: Vec<Variant>,
        bytes: Vec<u8>,
    }

    fn nested() -> Nested {
        Nested {
            z: true,
            sourceOrdinal: 42,
            a: Some("nested".to_owned()),
        }
    }

    #[test]
    fn structs_and_variants_match_round_trip() {
        let sample = Sample {
            nested: nested(),
            some: Some(nested()),
            none: None,
            floats32: [0.1, 1.0e-7],
            floats64: [0.1, 1e21, -0.0, f64::NAN, f64::INFINITY, f64::NEG_INFINITY],
            negative: -5,
            positive: 5,
            maximum: u64::MAX,
            integer_keys: BTreeMap::from([(2, "two".to_owned()), (1, "one".to_owned())]),
            string_keys: HashMap::from([("z".to_owned(), -5), ("a".to_owned(), 5)]),
            variants: vec![
                Variant::Unit,
                Variant::Newtype(nested()),
                Variant::Tuple(-5, None),
                Variant::Struct {
                    z: u64::MAX,
                    sourceOrdinal: 99,
                    a: nested(),
                },
                Variant::Ordinal(Box::new(Variant::Unit)),
            ],
            bytes: vec![0, 1, 127, 255],
        };
        assert_round_trip(&sample);
    }

    struct DuplicateMap;

    impl Serialize for DuplicateMap {
        fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
            let mut map = serializer.serialize_map(Some(5))?;
            map.serialize_entry("z", &1)?;
            map.serialize_entry(SOURCE_ORDINAL, &2)?;
            map.serialize_entry("a", &3)?;
            map.serialize_entry("z", &4)?;
            map.serialize_entry(SOURCE_ORDINAL, &5)?;
            map.end()
        }
    }

    #[test]
    fn map_preserves_order_and_duplicates() {
        let ordered = assert_round_trip(&DuplicateMap);
        let OrderedValue::Object(entries) = ordered else {
            panic!("map must serialize as an object");
        };
        assert_eq!(
            entries
                .iter()
                .map(|(key, _)| key.as_str())
                .collect::<Vec<_>>(),
            ["z", "a", "z"]
        );
    }

    struct PrivateNumber;

    impl Serialize for PrivateNumber {
        fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
            let mut number = serializer.serialize_struct("$serde_json::private::Number", 1)?;
            number.serialize_field("$serde_json::private::Number", "1")?;
            number.end()
        }
    }

    #[test]
    fn serde_json_private_structs_take_the_round_trip() {
        assert!(ordered_value(&PrivateNumber).is_none());
        assert_source_round_trip(&PrivateNumber);
    }

    struct OversizedHints;

    impl Serialize for OversizedHints {
        fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
            let mut sequence = serializer.serialize_seq(Some(usize::MAX))?;
            sequence.serialize_element(&OversizedMap)?;
            sequence.end()
        }
    }

    struct OversizedMap;

    impl Serialize for OversizedMap {
        fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
            let mut map = serializer.serialize_map(Some(usize::MAX))?;
            map.serialize_entry("a", &1)?;
            map.end()
        }
    }

    #[test]
    fn length_hints_preallocate_a_bounded_capacity() {
        assert!(ordered_value(&OversizedHints).is_some());
        assert_source_round_trip(&OversizedHints);
    }

    #[derive(Serialize)]
    struct Newtype(u32);

    #[derive(Serialize)]
    struct TupleStruct(bool, char);

    #[derive(Serialize)]
    struct UnitStruct;

    struct Bytes;

    impl Serialize for Bytes {
        fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
            serializer.serialize_bytes(&[0, 1, 127, 255])
        }
    }

    struct MapKeys;

    impl Serialize for MapKeys {
        fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
            let mut map = serializer.serialize_map(None)?;
            map.serialize_entry("quotes: \"; newline: \n", &1)?;
            map.serialize_entry(&String::from("owned"), &2)?;
            map.serialize_entry(&'🦀', &3)?;
            map.serialize_entry(&-5i8, &4)?;
            map.serialize_entry(&-5i16, &5)?;
            map.serialize_entry(&-5i32, &6)?;
            map.serialize_entry(&-5i64, &7)?;
            map.serialize_entry(&i128::MIN, &8)?;
            map.serialize_entry(&5u8, &9)?;
            map.serialize_entry(&5u16, &10)?;
            map.serialize_entry(&5u32, &11)?;
            map.serialize_entry(&u64::MAX, &12)?;
            map.serialize_entry(&u128::MAX, &13)?;
            map.serialize_entry(&Newtype(42), &14)?;
            map.serialize_entry(&Variant::Unit, &15)?;
            map.end()
        }
    }

    struct HumanReadable;

    impl Serialize for HumanReadable {
        fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
            if serializer.is_human_readable() {
                serializer.serialize_str("human readable")
            } else {
                serializer.serialize_bytes(&[0])
            }
        }
    }

    struct Failing;

    impl Serialize for Failing {
        fn serialize<S: serde::Serializer>(&self, _serializer: S) -> Result<S::Ok, S::Error> {
            Err(serde::ser::Error::custom("serialization failed"))
        }
    }

    struct MapKey<T>(T);

    impl<T: Serialize> Serialize for MapKey<T> {
        fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
            let mut map = serializer.serialize_map(Some(1))?;
            map.serialize_entry(&self.0, &1)?;
            map.end()
        }
    }

    fn nested_array(depth: u32) -> Value {
        let mut value = Value::Null;
        for _ in 0..depth {
            value = Value::Array(vec![value]);
        }
        value
    }

    fn nested_object(depth: u32) -> Value {
        let mut value = Value::Null;
        for _ in 0..depth {
            value = Value::Object(Map::from_iter([("value".to_owned(), value)]));
        }
        value
    }

    #[derive(Serialize)]
    enum DepthVariant {
        Unit,
        Newtype(Box<DepthVariant>),
        Tuple(Box<DepthVariant>, ()),
        Struct {
            value: Box<DepthVariant>,
        },
        #[serde(rename = "sourceOrdinal")]
        OrdinalNewtype(Box<DepthVariant>),
        #[serde(rename = "sourceOrdinal")]
        OrdinalTuple(Box<DepthVariant>, ()),
        #[serde(rename = "sourceOrdinal")]
        OrdinalStruct {
            value: Box<DepthVariant>,
        },
    }

    fn nested_variant(mut depth: u32) -> DepthVariant {
        let mut value = DepthVariant::Unit;
        let mut index = 0;
        while depth > 0 {
            if depth == 1 || index % 3 == 0 {
                value = DepthVariant::Newtype(Box::new(value));
                depth -= 1;
            } else if index % 3 == 1 {
                value = DepthVariant::Tuple(Box::new(value), ());
                depth -= 2;
            } else {
                value = DepthVariant::Struct {
                    value: Box::new(value),
                };
                depth -= 2;
            }
            index += 1;
        }
        value
    }

    #[test]
    fn nested_json_matches_round_trip() {
        for depth in DEPTHS {
            assert_depth_round_trip(&nested_array(depth), depth);
            assert_depth_round_trip(&nested_object(depth), depth);
        }
    }

    #[test]
    fn nested_variants_match_round_trip() {
        for depth in DEPTHS {
            assert_depth_round_trip(&nested_variant(depth), depth);
        }
    }

    #[test]
    fn discarded_values_match_round_trip() {
        for depth in DEPTHS {
            let value = BTreeMap::from([(SOURCE_ORDINAL, nested_array(depth - 1))]);
            assert_depth_round_trip(&value, depth);
            assert_depth_round_trip(
                &DepthVariant::OrdinalNewtype(Box::new(nested_variant(depth - 1))),
                depth,
            );
            assert_depth_round_trip(
                &DepthVariant::OrdinalTuple(Box::new(nested_variant(depth - 2)), ()),
                depth,
            );
            assert_depth_round_trip(
                &DepthVariant::OrdinalStruct {
                    value: Box::new(nested_variant(depth - 2)),
                },
                depth,
            );
        }
    }

    #[test]
    fn primitives_and_wrappers_match_round_trip() {
        assert_round_trip(&());
        assert_round_trip(&UnitStruct);
        assert_round_trip(&Newtype(5));
        assert_round_trip(&TupleStruct(true, '🦀'));
        assert_round_trip(&Bytes);
        assert_round_trip(&(-5i8, 5i8, -5i16, 5i16, -5i32, 5i32, -5i64, 5i64));
        assert_round_trip(&(5u8, 5u16, 5u32, 5u64));
        assert_round_trip(&(i64::MIN as i128, u64::MAX as i128, u64::MAX as u128));
        assert_round_trip(&BTreeMap::from([('🦀', "crab")]));
        assert_round_trip(&BTreeMap::from([(i128::MIN, "minimum")]));
        assert_round_trip(&BTreeMap::from([(u128::MAX, "maximum")]));
        assert_round_trip(&(f32::NAN, f32::INFINITY, f32::NEG_INFINITY, -0.0f32));
    }

    #[test]
    fn map_keys_match_round_trip() {
        assert_round_trip(&MapKeys);
    }

    #[test]
    fn serializers_are_human_readable() {
        assert_round_trip(&HumanReadable);
        assert!(matches!(
            HumanReadable.serialize(super::KeySerializer),
            Ok(key) if key == "human readable"
        ));
    }

    #[test]
    fn unsupported_values_match_round_trip() {
        assert_source_round_trip(&i128::MIN);
        assert_source_round_trip(&(i64::MIN as i128 - 1));
        assert_source_round_trip(&(u64::MAX as i128 + 1));
        assert_source_round_trip(&(u64::MAX as u128 + 1));
        assert_source_round_trip(&u128::MAX);
        assert_source_round_trip(&BTreeMap::from([(true, 1)]));
        assert_source_round_trip(&BTreeMap::from([(Some("key"), 1)]));
        assert_source_round_trip(&BTreeMap::from([((), 1)]));
        assert_source_round_trip(&BTreeMap::from([(vec![1u8], 1)]));
        assert_source_round_trip(&Failing);
        assert_source_round_trip(&BTreeMap::from([(SOURCE_ORDINAL, u128::MAX)]));
        assert_source_round_trip(&MapKey(0.1f32));
        assert_source_round_trip(&MapKey(0.1f64));
        assert_source_round_trip(&MapKey(Bytes));
        assert_source_round_trip(&MapKey(UnitStruct));
        assert_source_round_trip(&MapKey(TupleStruct(true, 'a')));
        assert_source_round_trip(&MapKey(Variant::Newtype(nested())));
    }

    struct SwallowedError;

    impl Serialize for SwallowedError {
        fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
            let mut map = serializer.serialize_map(None)?;
            let _ = map.serialize_entry(&true, &1);
            map.serialize_entry("fallback", &2)?;
            map.end()
        }
    }

    #[test]
    fn swallowed_errors_fall_back() {
        assert!(ordered_value(&SwallowedError).is_none());
        assert_source_round_trip(&SwallowedError);
        assert_source_round_trip(&[SwallowedError]);
    }

    enum ZeroHint {
        Seq,
        Map,
        Struct,
    }

    impl Serialize for ZeroHint {
        fn serialize<S: serde::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
            match self {
                Self::Seq => {
                    let mut seq = serializer.serialize_seq(Some(0))?;
                    seq.serialize_element(&1)?;
                    seq.end()
                }
                Self::Map => {
                    let mut map = serializer.serialize_map(Some(0))?;
                    map.serialize_entry(SOURCE_ORDINAL, &1)?;
                    map.end()
                }
                Self::Struct => {
                    let mut value = serializer.serialize_struct("ZeroHint", 0)?;
                    value.serialize_field("a", &1)?;
                    value.end()
                }
            }
        }
    }

    #[test]
    fn zero_length_hints_with_children_fall_back() {
        for hint in [ZeroHint::Seq, ZeroHint::Map, ZeroHint::Struct] {
            assert!(ordered_value(&hint).is_none());
            assert_source_round_trip(&hint);
        }
    }
}
