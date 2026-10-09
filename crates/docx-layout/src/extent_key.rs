use std::fmt;

use serde::ser::{self, Serialize};

pub(crate) fn encode<T: Serialize + ?Sized>(output: &mut Vec<u8>, value: &T) -> Result<(), Error> {
    value.serialize(&mut Encoder { output })
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

#[repr(u8)]
enum Tag {
    End,
    Unit,
    None,
    Some,
    False,
    True,
    I64,
    U64,
    F64,
    F32,
    Char,
    Str,
    Bytes,
    Seq,
    Tuple,
    TupleStruct,
    Map,
    Struct,
    UnitVariant,
    NewtypeVariant,
    TupleVariant,
    StructVariant,
    UnitStruct,
    NewtypeStruct,
    I128,
    U128,
}

struct Encoder<'a> {
    output: &'a mut Vec<u8>,
}

impl Encoder<'_> {
    fn tag(&mut self, tag: Tag) {
        self.output.push(tag as u8);
    }

    fn bytes(&mut self, bytes: &[u8]) {
        self.output
            .extend_from_slice(&(bytes.len() as u64).to_le_bytes());
        self.output.extend_from_slice(bytes);
    }

    fn named(&mut self, tag: Tag, name: &str) {
        self.tag(tag);
        self.bytes(name.as_bytes());
    }

    fn variant(&mut self, tag: Tag, index: u32, name: &str) {
        self.tag(tag);
        self.output.extend_from_slice(&index.to_le_bytes());
        self.bytes(name.as_bytes());
    }
}

impl<'a> ser::Serializer for &'a mut Encoder<'_> {
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
        self.tag(if value { Tag::True } else { Tag::False });
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
        self.tag(Tag::I64);
        self.output.extend_from_slice(&value.to_le_bytes());
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
        self.tag(Tag::U64);
        self.output.extend_from_slice(&value.to_le_bytes());
        Ok(())
    }

    fn serialize_i128(self, value: i128) -> Result<(), Error> {
        self.tag(Tag::I128);
        self.output.extend_from_slice(&value.to_le_bytes());
        Ok(())
    }

    fn serialize_u128(self, value: u128) -> Result<(), Error> {
        self.tag(Tag::U128);
        self.output.extend_from_slice(&value.to_le_bytes());
        Ok(())
    }

    fn serialize_f32(self, value: f32) -> Result<(), Error> {
        self.tag(Tag::F32);
        self.output
            .extend_from_slice(&value.to_bits().to_le_bytes());
        Ok(())
    }

    fn serialize_f64(self, value: f64) -> Result<(), Error> {
        self.tag(Tag::F64);
        self.output
            .extend_from_slice(&value.to_bits().to_le_bytes());
        Ok(())
    }

    fn serialize_char(self, value: char) -> Result<(), Error> {
        self.tag(Tag::Char);
        self.output.extend_from_slice(&(value as u32).to_le_bytes());
        Ok(())
    }

    fn serialize_str(self, value: &str) -> Result<(), Error> {
        self.tag(Tag::Str);
        self.bytes(value.as_bytes());
        Ok(())
    }

    fn serialize_bytes(self, value: &[u8]) -> Result<(), Error> {
        self.tag(Tag::Bytes);
        self.bytes(value);
        Ok(())
    }

    fn serialize_none(self) -> Result<(), Error> {
        self.tag(Tag::None);
        Ok(())
    }

    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<(), Error> {
        self.tag(Tag::Some);
        value.serialize(self)
    }

    fn serialize_unit(self) -> Result<(), Error> {
        self.tag(Tag::Unit);
        Ok(())
    }

    fn serialize_unit_struct(self, name: &'static str) -> Result<(), Error> {
        self.named(Tag::UnitStruct, name);
        Ok(())
    }

    fn serialize_unit_variant(
        self,
        _: &'static str,
        index: u32,
        variant: &'static str,
    ) -> Result<(), Error> {
        self.variant(Tag::UnitVariant, index, variant);
        Ok(())
    }

    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        name: &'static str,
        value: &T,
    ) -> Result<(), Error> {
        self.named(Tag::NewtypeStruct, name);
        value.serialize(self)
    }

    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        index: u32,
        variant: &'static str,
        value: &T,
    ) -> Result<(), Error> {
        self.variant(Tag::NewtypeVariant, index, variant);
        value.serialize(self)
    }

    fn serialize_seq(self, _: Option<usize>) -> Result<Self, Error> {
        self.tag(Tag::Seq);
        Ok(self)
    }

    fn serialize_tuple(self, _: usize) -> Result<Self, Error> {
        self.tag(Tag::Tuple);
        Ok(self)
    }

    fn serialize_tuple_struct(self, name: &'static str, _: usize) -> Result<Self, Error> {
        self.named(Tag::TupleStruct, name);
        Ok(self)
    }

    fn serialize_tuple_variant(
        self,
        _: &'static str,
        index: u32,
        variant: &'static str,
        _: usize,
    ) -> Result<Self, Error> {
        self.variant(Tag::TupleVariant, index, variant);
        Ok(self)
    }

    fn serialize_map(self, _: Option<usize>) -> Result<Self, Error> {
        self.tag(Tag::Map);
        Ok(self)
    }

    fn serialize_struct(self, name: &'static str, _: usize) -> Result<Self, Error> {
        self.named(Tag::Struct, name);
        Ok(self)
    }

    fn serialize_struct_variant(
        self,
        _: &'static str,
        index: u32,
        variant: &'static str,
        _: usize,
    ) -> Result<Self, Error> {
        self.variant(Tag::StructVariant, index, variant);
        Ok(self)
    }
}

impl ser::SerializeSeq for &mut Encoder<'_> {
    type Ok = ();
    type Error = Error;

    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), Error> {
        self.tag(Tag::End);
        Ok(())
    }
}

impl ser::SerializeTuple for &mut Encoder<'_> {
    type Ok = ();
    type Error = Error;

    fn serialize_element<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        ser::SerializeSeq::serialize_element(self, value)
    }

    fn end(self) -> Result<(), Error> {
        ser::SerializeSeq::end(self)
    }
}

impl ser::SerializeTupleStruct for &mut Encoder<'_> {
    type Ok = ();
    type Error = Error;

    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        ser::SerializeSeq::serialize_element(self, value)
    }

    fn end(self) -> Result<(), Error> {
        ser::SerializeSeq::end(self)
    }
}

impl ser::SerializeTupleVariant for &mut Encoder<'_> {
    type Ok = ();
    type Error = Error;

    fn serialize_field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        ser::SerializeSeq::serialize_element(self, value)
    }

    fn end(self) -> Result<(), Error> {
        ser::SerializeSeq::end(self)
    }
}

impl ser::SerializeMap for &mut Encoder<'_> {
    type Ok = ();
    type Error = Error;

    fn serialize_key<T: Serialize + ?Sized>(&mut self, key: &T) -> Result<(), Error> {
        key.serialize(&mut **self)
    }

    fn serialize_value<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Error> {
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), Error> {
        ser::SerializeSeq::end(self)
    }
}

impl ser::SerializeStruct for &mut Encoder<'_> {
    type Ok = ();
    type Error = Error;

    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), Error> {
        ser::Serializer::serialize_str(&mut **self, key)?;
        value.serialize(&mut **self)
    }

    fn end(self) -> Result<(), Error> {
        ser::SerializeSeq::end(self)
    }
}

impl ser::SerializeStructVariant for &mut Encoder<'_> {
    type Ok = ();
    type Error = Error;

    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), Error> {
        ser::SerializeStruct::serialize_field(self, key, value)
    }

    fn end(self) -> Result<(), Error> {
        ser::SerializeSeq::end(self)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::BTreeMap;

    fn key<T: Serialize + ?Sized>(value: &T) -> Vec<u8> {
        let mut output = Vec::new();
        encode(&mut output, value).unwrap();
        output
    }

    #[test]
    fn string_boundaries_are_distinct() {
        assert_ne!(key(&("ab", "c")), key(&("a", "bc")));
        assert_ne!(key(&vec!["ab", "c"]), key(&vec!["a", "bc"]));
    }

    #[test]
    fn nested_options_are_distinct() {
        assert_ne!(key(&Some(None::<u8>)), key(&None::<Option<u8>>));
        assert_ne!(key(&None::<u8>), key(&()));
    }

    #[test]
    fn scalar_and_named_types_are_distinct() {
        #[derive(serde::Serialize)]
        struct Unit;
        #[derive(serde::Serialize)]
        struct Newtype(u8);
        #[derive(serde::Serialize)]
        struct Tuple(u8, u8);

        struct Bytes<'a>(&'a [u8]);

        impl Serialize for Bytes<'_> {
            fn serialize<S: ser::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
                serializer.serialize_bytes(self.0)
            }
        }

        assert_ne!(key(&()), key(&Unit));
        assert_ne!(key(&false), key(&true));
        assert_ne!(key(&1_u8), key(&Newtype(1)));
        assert_ne!(key(&(1_u8, 2_u8)), key(&Tuple(1, 2)));
        assert_ne!(key(&'a'), key(&"a"));
        assert_ne!(key(&"a"), key(&Bytes(b"a")));
        assert_ne!(key(&Bytes(b"a")), key(&vec![b'a']));
    }

    #[test]
    fn numeric_types_and_float_bits_are_distinct() {
        assert_ne!(key(&1_u64), key(&1_i64));
        assert_ne!(key(&1_u64), key(&1.0_f64));
        assert_ne!(key(&1_i64), key(&1.0_f64));
        assert_ne!(key(&1.0_f32), key(&1.0_f64));
        assert_ne!(key(&0.0_f64), key(&-0.0_f64));
        assert_ne!(key(&0.0_f32), key(&-0.0_f32));
        assert_ne!(
            key(&f64::from_bits(0x7ff8_0000_0000_0001)),
            key(&f64::from_bits(0x7ff8_0000_0000_0002))
        );
        assert_ne!(
            key(&f32::from_bits(0x7fc0_0001)),
            key(&f32::from_bits(0x7fc0_0002))
        );
        assert_ne!(key(&i128::MIN), key(&i128::MAX));
        assert_ne!(key(&i128::MAX), key(&(i128::MAX as u128)));
        assert_ne!(key(&0_u128), key(&u128::MAX));
    }

    #[test]
    fn sequence_boundaries_are_distinct() {
        assert_ne!(key(&Vec::<()>::new()), key(&vec![()]));
        assert_ne!(key(&vec![vec![1_u8], vec![2]]), key(&vec![vec![1_u8, 2]]));
        assert_ne!(key(&vec![1_u8, 2]), key(&(1_u8, 2_u8)));
    }

    #[test]
    fn optional_struct_fields_are_distinct() {
        #[derive(serde::Serialize)]
        struct Fields {
            #[serde(rename = "", skip_serializing_if = "Option::is_none")]
            a: Option<u8>,
            #[serde(skip_serializing_if = "Option::is_none")]
            b: Option<u8>,
        }

        let absent = Fields { a: None, b: None };
        let a = Fields {
            a: Some(1),
            b: None,
        };
        let b = Fields {
            a: None,
            b: Some(1),
        };
        assert_ne!(key(&absent), key(&a));
        assert_ne!(key(&a), key(&b));
    }

    #[test]
    fn enum_variants_with_equal_payloads_are_distinct() {
        #[derive(serde::Serialize)]
        enum Variant {
            #[serde(rename = "same")]
            A(u8),
            #[serde(rename = "same")]
            B(u8),
            Unit,
            Tuple(u8, u8),
            Struct {
                value: u8,
            },
        }

        assert_ne!(key(&Variant::A(1)), key(&Variant::B(1)));
        let keys = [
            key(&Variant::A(1)),
            key(&Variant::Unit),
            key(&Variant::Tuple(1, 1)),
            key(&Variant::Struct { value: 1 }),
        ];
        for (index, key) in keys.iter().enumerate() {
            assert!(!keys[index + 1..].contains(key));
        }
    }

    #[test]
    fn sequence_length_hints_do_not_change_encoding() {
        struct UnknownLength<'a>(&'a [u8]);

        impl Serialize for UnknownLength<'_> {
            fn serialize<S: ser::Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
                let mut seq = serializer.serialize_seq(None)?;
                for value in self.0 {
                    ser::SerializeSeq::serialize_element(&mut seq, value)?;
                }
                ser::SerializeSeq::end(seq)
            }
        }

        for values in [vec![], vec![1], vec![1, 2]] {
            assert_eq!(key(&values), key(&UnknownLength(&values)));
        }
    }

    #[test]
    fn map_keys_use_the_same_encoding() {
        let a = BTreeMap::from([(("ab", "c"), 1_u8)]);
        let b = BTreeMap::from([(("a", "bc"), 1_u8)]);
        assert_ne!(key(&a), key(&b));
        assert_ne!(
            key(&BTreeMap::from([(1_u64, ())])),
            key(&BTreeMap::from([(1_i64, ())]))
        );
        assert_ne!(key(&BTreeMap::<u8, ()>::new()), key(&vec![(1_u8, ())]));
    }

    #[test]
    fn json_values_use_generic_encoding() {
        let value = serde_json::json!({"runs": [null, true, "quoted\"\n", {"n": 1}]});
        let mut changed = value.clone();
        changed["runs"][3]["n"] = serde_json::json!(2);
        assert_ne!(key(&value), key(&changed));
        assert_eq!(key(&serde_json::json!(-1)), key(&-1_i64));
        assert_eq!(key(&serde_json::json!(1)), key(&1_u64));
        assert_eq!(key(&serde_json::json!(1.0)), key(&1.0_f64));
    }
}
