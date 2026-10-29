#![cfg_attr(not(test), no_main)]

use std::ops::Range;

#[cfg(not(test))]
use libfuzzer_sys::fuzz_target;
use serde::ser::{SerializeMap, SerializeStruct, SerializeStructVariant};
use serde::{Serialize, Serializer};
use xlsx_ops::{Op, OracleOp};

enum Shape {
    Scalar,
    Seq(Vec<Shape>),
    Map(Vec<(String, Shape)>),
}

struct ShapeSerializer;

struct Sequence {
    values: Vec<Shape>,
    variant: Option<&'static str>,
}

struct Object {
    fields: Vec<(String, Shape)>,
    key: Option<String>,
    variant: Option<&'static str>,
}

macro_rules! scalar {
    ($($method:ident($ty:ty)),* $(,)?) => {
        $(
            fn $method(self, _value: $ty) -> serde_json::Result<Shape> {
                Ok(Shape::Scalar)
            }
        )*
    };
}

impl Serializer for ShapeSerializer {
    type Ok = Shape;
    type Error = serde_json::Error;
    type SerializeSeq = Sequence;
    type SerializeTuple = Sequence;
    type SerializeTupleStruct = Sequence;
    type SerializeTupleVariant = Sequence;
    type SerializeMap = Object;
    type SerializeStruct = Object;
    type SerializeStructVariant = Object;

    scalar! {
        serialize_bool(bool), serialize_i8(i8), serialize_i16(i16), serialize_i32(i32),
        serialize_i64(i64), serialize_i128(i128), serialize_u8(u8), serialize_u16(u16),
        serialize_u32(u32), serialize_u64(u64), serialize_u128(u128), serialize_f32(f32),
        serialize_f64(f64), serialize_char(char), serialize_str(&str), serialize_bytes(&[u8]),
    }

    fn serialize_none(self) -> serde_json::Result<Shape> {
        Ok(Shape::Scalar)
    }

    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> serde_json::Result<Shape> {
        value.serialize(self)
    }

    fn serialize_unit(self) -> serde_json::Result<Shape> {
        Ok(Shape::Scalar)
    }

    fn serialize_unit_struct(self, _name: &'static str) -> serde_json::Result<Shape> {
        Ok(Shape::Scalar)
    }

    fn serialize_unit_variant(
        self,
        _name: &'static str,
        _index: u32,
        _variant: &'static str,
    ) -> serde_json::Result<Shape> {
        Ok(Shape::Scalar)
    }

    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _name: &'static str,
        value: &T,
    ) -> serde_json::Result<Shape> {
        value.serialize(self)
    }

    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _name: &'static str,
        _index: u32,
        variant: &'static str,
        value: &T,
    ) -> serde_json::Result<Shape> {
        Ok(Shape::Map(vec![(variant.into(), value.serialize(self)?)]))
    }

    fn serialize_seq(self, _len: Option<usize>) -> serde_json::Result<Sequence> {
        Ok(Sequence {
            values: Vec::new(),
            variant: None,
        })
    }

    fn serialize_tuple(self, len: usize) -> serde_json::Result<Sequence> {
        self.serialize_seq(Some(len))
    }

    fn serialize_tuple_struct(
        self,
        _name: &'static str,
        len: usize,
    ) -> serde_json::Result<Sequence> {
        self.serialize_seq(Some(len))
    }

    fn serialize_tuple_variant(
        self,
        _name: &'static str,
        _index: u32,
        variant: &'static str,
        len: usize,
    ) -> serde_json::Result<Sequence> {
        let mut seq = self.serialize_seq(Some(len))?;
        seq.variant = Some(variant);
        Ok(seq)
    }

    fn serialize_map(self, _len: Option<usize>) -> serde_json::Result<Object> {
        Ok(Object {
            fields: Vec::new(),
            key: None,
            variant: None,
        })
    }

    fn serialize_struct(self, _name: &'static str, len: usize) -> serde_json::Result<Object> {
        self.serialize_map(Some(len))
    }

    fn serialize_struct_variant(
        self,
        _name: &'static str,
        _index: u32,
        variant: &'static str,
        len: usize,
    ) -> serde_json::Result<Object> {
        let mut map = self.serialize_map(Some(len))?;
        map.variant = Some(variant);
        Ok(map)
    }
}

macro_rules! sequence {
    ($trait:ident, $method:ident) => {
        impl serde::ser::$trait for Sequence {
            type Ok = Shape;
            type Error = serde_json::Error;

            fn $method<T: Serialize + ?Sized>(&mut self, value: &T) -> serde_json::Result<()> {
                self.values.push(value.serialize(ShapeSerializer)?);
                Ok(())
            }

            fn end(self) -> serde_json::Result<Shape> {
                let shape = Shape::Seq(self.values);
                Ok(match self.variant {
                    Some(variant) => Shape::Map(vec![(variant.into(), shape)]),
                    None => shape,
                })
            }
        }
    };
}

sequence!(SerializeSeq, serialize_element);
sequence!(SerializeTuple, serialize_element);
sequence!(SerializeTupleStruct, serialize_field);
sequence!(SerializeTupleVariant, serialize_field);

impl SerializeMap for Object {
    type Ok = Shape;
    type Error = serde_json::Error;

    fn serialize_key<T: Serialize + ?Sized>(&mut self, key: &T) -> serde_json::Result<()> {
        let key = serde_json::to_value(key)?;
        self.key = Some(match key.as_str() {
            Some(key) => key.to_owned(),
            None => key.to_string(),
        });
        Ok(())
    }

    fn serialize_value<T: Serialize + ?Sized>(&mut self, value: &T) -> serde_json::Result<()> {
        self.fields.push((
            self.key.take().expect("map key"),
            value.serialize(ShapeSerializer)?,
        ));
        Ok(())
    }

    fn end(self) -> serde_json::Result<Shape> {
        Ok(Shape::Map(self.fields))
    }
}

macro_rules! object {
    ($trait:ident) => {
        impl $trait for Object {
            type Ok = Shape;
            type Error = serde_json::Error;

            fn serialize_field<T: Serialize + ?Sized>(
                &mut self,
                key: &'static str,
                value: &T,
            ) -> serde_json::Result<()> {
                self.fields.push((key.into(), value.serialize(ShapeSerializer)?));
                Ok(())
            }

            fn skip_field(&mut self, key: &'static str) -> serde_json::Result<()> {
                self.fields.push((key.into(), Shape::Scalar));
                Ok(())
            }

            fn end(self) -> serde_json::Result<Shape> {
                let shape = Shape::Map(self.fields);
                Ok(match self.variant {
                    Some(variant) => Shape::Map(vec![(variant.into(), shape)]),
                    None => shape,
                })
            }
        }
    };
}

object!(SerializeStruct);
object!(SerializeStructVariant);

struct Member {
    start: usize,
    name: Option<String>,
    value: Range<usize>,
}

fn whitespace(input: &[u8], mut at: usize) -> usize {
    while input.get(at).is_some_and(u8::is_ascii_whitespace) {
        at += 1;
    }
    at
}

fn end_value(input: &[u8], start: usize) -> Option<usize> {
    let mut at = start;
    let mut depth = 0usize;
    loop {
        match *input.get(at)? {
            b'"' => {
                at += 1;
                loop {
                    match *input.get(at)? {
                        b'\\' => at += 2,
                        b'"' => {
                            at += 1;
                            break;
                        }
                        _ => at += 1,
                    }
                }
                if depth == 0 {
                    return Some(at);
                }
            }
            b'[' | b'{' => {
                depth += 1;
                at += 1;
            }
            b']' | b'}' if depth > 0 => {
                depth -= 1;
                at += 1;
                if depth == 0 {
                    return Some(at);
                }
            }
            byte if depth == 0 => {
                if byte.is_ascii_whitespace() || matches!(byte, b',' | b']' | b'}') {
                    return (at > start).then_some(at);
                }
                at += 1;
                if at == input.len() {
                    return Some(at);
                }
            }
            _ => at += 1,
        }
    }
}

fn root(input: &[u8]) -> Range<usize> {
    let start = whitespace(input, 0);
    start..end_value(input, start).expect("accepted JSON value")
}

fn members(input: &[u8], span: Range<usize>) -> Option<Vec<Member>> {
    let object = match input.get(span.start)? {
        b'{' => true,
        b'[' => false,
        _ => return None,
    };
    let mut at = whitespace(input, span.start + 1);
    let mut fields = Vec::new();
    while at < span.end - 1 {
        let start = at;
        let name = if object {
            let end = end_value(input, at)?;
            let name = serde_json::from_slice::<String>(&input[at..end]).ok()?;
            at = whitespace(input, end);
            if input.get(at) != Some(&b':') {
                return None;
            }
            at = whitespace(input, at + 1);
            Some(name)
        } else {
            None
        };
        let end = end_value(input, at)?;
        fields.push(Member {
            start,
            name,
            value: at..end,
        });
        at = whitespace(input, end);
        if input.get(at) == Some(&b',') {
            at = whitespace(input, at + 1);
        } else if at != span.end - 1 {
            return None;
        }
    }
    Some(fields)
}

fn normalize_nested(
    input: &[u8],
    span: Range<usize>,
    canonical: &serde_json::Value,
    shape: Option<&Shape>,
    replacements: &mut Vec<(Range<usize>, Vec<u8>)>,
) {
    let Some(fields) = members(input, span.clone()) else {
        return;
    };
    if input[span.start] == b'['
        && canonical
            .as_object()
            .is_some_and(|value| value.contains_key("kind") || value.contains_key("type"))
        && fields.first().is_some_and(|field| {
            serde_json::from_slice::<serde_json::Value>(&input[field.value.clone()])
                .is_ok_and(|value| value.is_string() || value.is_number())
        })
    {
        replacements.push((span, serde_json::to_vec(canonical).unwrap()));
        return;
    }
    for (index, field) in fields.into_iter().enumerate() {
        let name = field.name.as_deref().or_else(|| match shape {
            Some(Shape::Map(fields)) => fields.get(index).map(|(name, _)| name.as_str()),
            _ => None,
        });
        let value = match name {
            Some(name) => canonical.get(name),
            None => canonical.get(index),
        };
        let Some(value) = value else {
            continue;
        };
        if matches!(field.name.as_deref(), Some("kind" | "type"))
            && value.is_string()
            && serde_json::from_slice::<serde_json::Value>(&input[field.value.clone()])
                .is_ok_and(|value| value.is_number())
        {
            replacements.push((field.value, serde_json::to_vec(value).unwrap()));
        } else {
            let child = match shape {
                Some(Shape::Map(fields)) => name.and_then(|name| matching_field(fields, name)),
                Some(Shape::Seq(fields)) => fields.get(index),
                _ => None,
            };
            normalize_nested(input, field.value, value, child, replacements);
        }
    }
}

fn normalized_fields(
    input: &[u8],
    span: Range<usize>,
    oracle: &OracleOp,
) -> Vec<(Range<usize>, Vec<u8>)> {
    let canonical = serde_json::to_value(oracle).unwrap();
    let Shape::Map(schema) = oracle.serialize(ShapeSerializer).unwrap() else {
        panic!("operation fields");
    };
    let mut replacements = Vec::new();
    for (index, field) in members(input, span)
        .expect("accepted operation container")
        .into_iter()
        .enumerate()
    {
        let name = field
            .name
            .as_deref()
            .or_else(|| schema.get(index).map(|field| field.0.as_str()));
        if name == Some("type") {
            continue;
        }
        if let Some(value) = name.and_then(|name| canonical.get(name)) {
            let shape = name.and_then(|name| matching_field(&schema, name));
            normalize_nested(input, field.value, value, shape, &mut replacements);
        }
    }
    replacements
}

fn replace_fields(input: &[u8], replacements: Vec<(Range<usize>, Vec<u8>)>) -> Vec<u8> {
    let mut output = Vec::new();
    let mut at = 0;
    for (span, value) in replacements {
        output.extend_from_slice(&input[at..span.start]);
        output.extend(value);
        at = span.end;
    }
    output.extend_from_slice(&input[at..]);
    output
}

fn matching_field<'a>(fields: &'a [(String, Shape)], name: &str) -> Option<&'a Shape> {
    let numeric = name.parse::<u64>().ok().map(|value| value.to_string());
    fields
        .iter()
        .find(|(key, _)| key == name || numeric.as_ref() == Some(key))
        .map(|(_, shape)| shape)
}

fn strip_unknown(input: &[u8], span: Range<usize>, shape: &Shape) -> Vec<u8> {
    let Some(fields) = members(input, span.clone()) else {
        return input[span].to_vec();
    };
    let object = input[span.start] == b'{';
    let mut output = vec![input[span.start]];
    for (index, field) in fields.into_iter().enumerate() {
        let child = match shape {
            Shape::Map(schema) if object => {
                let Some(shape) = field
                    .name
                    .as_deref()
                    .and_then(|name| matching_field(schema, name))
                else {
                    continue;
                };
                Some(shape)
            }
            Shape::Map(schema) => schema.get(index).map(|(_, shape)| shape),
            Shape::Seq(schema) if !object => schema.get(index),
            _ => return input[span].to_vec(),
        };
        if output.len() > 1 {
            output.push(b',');
        }
        if object {
            output.extend_from_slice(&input[field.start..field.value.start]);
        }
        match child {
            Some(shape) => output.extend(strip_unknown(input, field.value, shape)),
            None => output.extend_from_slice(&input[field.value]),
        }
    }
    output.push(if object { b'}' } else { b']' });
    output
}

fn check_op(input: &[u8], direct: serde_json::Result<Op>, oracle: serde_json::Result<OracleOp>) {
    match (&direct, &oracle) {
        (Ok(direct), Ok(oracle)) => {
            assert_eq!(direct, &Op::from(oracle.clone()));
            assert_eq!(serde_json::to_vec(direct).unwrap(), serde_json::to_vec(oracle).unwrap());
        }
        (Err(_), Ok(oracle)) => {
            let fields = normalized_fields(input, root(input), oracle);
            let rebuilt = replace_fields(input, fields);
            assert_eq!(serde_json::from_slice::<Op>(&rebuilt).unwrap(), Op::from(oracle.clone()));
            assert_eq!(
                serde_json::from_str::<Op>(std::str::from_utf8(&rebuilt).unwrap()).unwrap(),
                Op::from(oracle.clone())
            );
            assert_eq!(
                serde_json::from_value::<Op>(serde_json::from_slice(&rebuilt).unwrap()).unwrap(),
                Op::from(oracle.clone())
            );
        }
        (Ok(direct), Err(_)) => {
            let shape = direct.serialize(ShapeSerializer).unwrap();
            let stripped = strip_unknown(input, root(input), &shape);
            assert_eq!(
                Op::from(serde_json::from_slice::<OracleOp>(&stripped).unwrap()),
                *direct
            );
        }
        (Err(_), Err(_)) => {}
    }
    for canonical in [
        direct.as_ref().ok().map(|op| serde_json::to_vec(op).unwrap()),
        oracle.as_ref().ok().map(|op| serde_json::to_vec(op).unwrap()),
    ]
    .into_iter()
    .flatten()
    {
        let decoded = serde_json::from_slice::<Op>(&canonical).unwrap();
        let replayed = Op::from(serde_json::from_slice::<OracleOp>(&canonical).unwrap());
        assert_eq!(decoded, replayed);
        if let Ok(original) = &direct {
            assert_eq!(&decoded, original);
        }
        if let Ok(original) = &oracle {
            assert_eq!(decoded, Op::from(original.clone()));
        }
    }
}

fn check_ops(
    input: &[u8],
    direct: serde_json::Result<Vec<Op>>,
    oracle: serde_json::Result<Vec<OracleOp>>,
) {
    if let (Err(_), Ok(oracle)) = (&direct, &oracle) {
        let entries = members(input, root(input)).expect("accepted operation vector");
        let fields = entries
            .into_iter()
            .zip(oracle)
            .flat_map(|(entry, op)| normalized_fields(input, entry.value, op))
            .collect();
        let rebuilt = replace_fields(input, fields);
        let expected = oracle.iter().cloned().map(Op::from).collect::<Vec<_>>();
        assert_eq!(serde_json::from_slice::<Vec<Op>>(&rebuilt).unwrap(), expected);
        assert_eq!(
            serde_json::from_str::<Vec<Op>>(std::str::from_utf8(&rebuilt).unwrap()).unwrap(),
            expected
        );
        assert_eq!(
            serde_json::from_value::<Vec<Op>>(serde_json::from_slice(&rebuilt).unwrap()).unwrap(),
            expected
        );
    }
    if let (Ok(direct), Err(_)) = (&direct, &oracle) {
        let shape = direct.serialize(ShapeSerializer).unwrap();
        let stripped = strip_unknown(input, root(input), &shape);
        assert_eq!(
            serde_json::from_slice::<Vec<OracleOp>>(&stripped)
                .unwrap()
                .into_iter()
                .map(Op::from)
                .collect::<Vec<_>>(),
            *direct
        );
    }
    let direct = direct.map(|ops| (serde_json::to_vec(&ops).unwrap(), ops));
    let oracle = oracle.map(|ops| {
        (
            serde_json::to_vec(&ops).unwrap(),
            ops.into_iter().map(Op::from).collect::<Vec<_>>(),
        )
    });
    if let (Ok(direct), Ok(oracle)) = (&direct, &oracle) {
        assert_eq!(direct, oracle);
    }
    for (canonical, ops) in [direct.ok(), oracle.ok()].into_iter().flatten() {
        assert_eq!(serde_json::from_slice::<Vec<Op>>(&canonical).unwrap(), ops);
        assert_eq!(
            serde_json::from_slice::<Vec<OracleOp>>(&canonical)
                .unwrap()
                .into_iter()
                .map(Op::from)
                .collect::<Vec<_>>(),
            ops
        );
    }
}

#[cfg(not(test))]
fuzz_target!(|data: &[u8]| {
    check_op(data, serde_json::from_slice(data), serde_json::from_slice(data));
    check_ops(data, serde_json::from_slice(data), serde_json::from_slice(data));
    if let Ok(text) = std::str::from_utf8(data) {
        check_op(data, serde_json::from_str(text), serde_json::from_str(text));
        check_ops(data, serde_json::from_str(text), serde_json::from_str(text));
    }
    if let Ok(value) = serde_json::from_slice::<serde_json::Value>(data) {
        let input = serde_json::to_vec(&value).unwrap();
        check_op(
            &input,
            serde_json::from_value(value.clone()),
            serde_json::from_value(value.clone()),
        );
        check_ops(
            &input,
            serde_json::from_value(value.clone()),
            serde_json::from_value(value),
        );
    }
});

#[cfg(test)]
mod tests {
    use super::*;

    fn normalize(input: &str) -> Vec<u8> {
        let oracle = serde_json::from_str::<OracleOp>(input).unwrap();
        let input = input.as_bytes();
        replace_fields(input, normalized_fields(input, root(input), &oracle))
    }

    #[test]
    fn numeric_nested_tag_is_substituted() {
        let input = concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":{"kind":1,"value":3.50}}}"##,
        );
        let expected = concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":{"kind":"number","value":3.50}}}"##,
        );
        assert_eq!(normalize(input), expected.as_bytes());
    }

    #[test]
    fn ordinary_float_spelling_is_untouched() {
        let input = r##"{"type":"setColWidth","sheet":0,"col":0,"width":0.10}"##;
        let oracle = serde_json::from_str::<OracleOp>(input).unwrap();
        assert!(normalized_fields(input.as_bytes(), root(input.as_bytes()), &oracle).is_empty());
        assert_eq!(normalize(input), input.as_bytes());
    }

    #[test]
    fn positional_nested_enum_is_substituted() {
        let input = concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":[1,3.50]}}"##,
        );
        let expected = concat!(
            r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
            r##""cell":{"value":{"kind":"number","value":3.5}}}"##,
        );
        assert_eq!(normalize(input), expected.as_bytes());
    }

    #[test]
    fn positional_cell_ref_is_untouched() {
        let input = concat!(
            r##"{"type":"setCell","sheet":0,"at":[0,0,false,false],"##,
            r##""cell":{"value":{"kind":"empty"}}}"##,
        );
        let oracle = serde_json::from_str::<OracleOp>(input).unwrap();
        let bytes = input.as_bytes();
        assert!(normalized_fields(bytes, root(bytes), &oracle).is_empty());
        assert_eq!(normalize(input), bytes);
        check_op(bytes, serde_json::from_str(input), Ok(oracle));
    }

    #[test]
    fn fuzz_helpers_cover_matching_and_one_sided_inputs() {
        for input in [
            r##"{"type":"setColWidth","sheet":0,"col":0,"width":0.10}"##,
            concat!(
                r##"{"type":"setCell","sheet":0,"at":{"row":0,"col":0},"##,
                r##""cell":{"value":{"kind":1,"value":3.50}}}"##,
            ),
            r##"{"type":"removeSheet","index":0,"unknown":1e309}"##,
        ] {
            let bytes = input.as_bytes();
            check_op(bytes, serde_json::from_slice(bytes), serde_json::from_slice(bytes));
            let vector = format!("[{input}]");
            let bytes = vector.as_bytes();
            check_ops(bytes, serde_json::from_slice(bytes), serde_json::from_slice(bytes));
        }
    }

    #[test]
    fn untouched_fields_keep_their_raw_text() {
        for input in [
            r##"{
  "s\u0068eet": 0,
  "t\u0079pe": "setCell",
  "at": { "row": 0, "col": 0 },
  "unknown": {"same":0.10,"same":"a\u0062"},
  "unknown": true,
  "cell": {"formula":"\u0031","value":{"kind":1,"value":0.10}}
}"##,
            concat!(
                r##"["setCell", 0, {"row":0,"col":0}, "##,
                r##"{"formula":"\u0031","value":{"kind":1,"value":0.10}}]"##,
            ),
        ] {
            let expected = input.replace(r##""kind":1"##, r##""kind":"number""##);
            assert_eq!(normalize(input), expected.as_bytes());
        }
    }
}
