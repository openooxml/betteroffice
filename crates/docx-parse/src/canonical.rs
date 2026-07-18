//! `docx-document-canonical-v1`, the byte contract shared with TypeScript.

use std::collections::BTreeSet;

use base64::Engine as _;
use serde::Serialize;
use sha2::{Digest, Sha256};
use thiserror::Error;

pub const VERSION: &str = "docx-document-canonical-v1";

/// Values accepted by the deliberately narrow canonical contract.
#[derive(Clone, Debug, PartialEq)]
pub enum CanonicalValue {
    Null,
    Bool(bool),
    String(String),
    Number(f64),
    Array(Vec<CanonicalValue>),
    /// Plain-object entries. Encoding sorts keys by Unicode scalar value.
    Object(Vec<(String, CanonicalValue)>),
    /// Map entries. Encoding preserves insertion order and requires unique keys.
    OrderedMap(Vec<(String, CanonicalValue)>),
    Binary(Vec<u8>),
    /// A validated UTC ISO-8601 string supplied by a typed model adapter.
    Date(String),
}

#[derive(Debug, Error, PartialEq, Eq)]
pub enum CanonicalError {
    #[error("canonical object contains duplicate key {0:?}")]
    DuplicateObjectKey(String),
    #[error("canonical map contains duplicate key {0:?}")]
    DuplicateMapKey(String),
    #[error("canonical JSON cannot encode a non-finite number")]
    NonFiniteNumber,
    #[error("canonical date is not a UTC ISO-8601 value: {0:?}")]
    InvalidDate(String),
    #[error("canonical serialization failed: {0}")]
    Serialization(String),
}

/// Encode the header, one compact JSON value, and the required trailing LF.
pub fn to_canonical_bytes(value: &CanonicalValue) -> Result<Vec<u8>, CanonicalError> {
    let mut output = Vec::new();
    output.extend_from_slice(VERSION.as_bytes());
    output.push(b'\n');
    encode_value(value, &mut output)?;
    output.push(b'\n');
    Ok(output)
}

/// SHA-256 of the exact canonical bytes, rendered as lowercase hexadecimal.
pub fn canonical_sha256(value: &CanonicalValue) -> Result<String, CanonicalError> {
    let bytes = to_canonical_bytes(value)?;
    Ok(format!("{:x}", Sha256::digest(bytes)))
}

/// Convert a typed serde DTO into the strict canonical value vocabulary.
pub fn from_serializable<T: Serialize + ?Sized>(
    value: &T,
) -> Result<CanonicalValue, CanonicalError> {
    let value = serde_json::to_value(value)
        .map_err(|error| CanonicalError::Serialization(error.to_string()))?;
    from_json_value(value)
}

fn from_json_value(value: serde_json::Value) -> Result<CanonicalValue, CanonicalError> {
    match value {
        serde_json::Value::Null => Ok(CanonicalValue::Null),
        serde_json::Value::Bool(value) => Ok(CanonicalValue::Bool(value)),
        serde_json::Value::String(value) => Ok(CanonicalValue::String(value)),
        serde_json::Value::Number(value) => value
            .as_f64()
            .filter(|value| value.is_finite())
            .map(CanonicalValue::Number)
            .ok_or(CanonicalError::NonFiniteNumber),
        serde_json::Value::Array(values) => values
            .into_iter()
            .map(from_json_value)
            .collect::<Result<Vec<_>, _>>()
            .map(CanonicalValue::Array),
        serde_json::Value::Object(entries) => entries
            .into_iter()
            .map(|(key, value)| Ok((key, from_json_value(value)?)))
            .collect::<Result<Vec<_>, CanonicalError>>()
            .map(CanonicalValue::Object),
    }
}

fn encode_value(value: &CanonicalValue, output: &mut Vec<u8>) -> Result<(), CanonicalError> {
    match value {
        CanonicalValue::Null => output.extend_from_slice(b"null"),
        CanonicalValue::Bool(value) => {
            output.extend_from_slice(if *value { b"true" } else { b"false" })
        }
        CanonicalValue::String(value) => encode_string(value, output)?,
        CanonicalValue::Number(value) => encode_number(*value, output)?,
        CanonicalValue::Array(values) => {
            output.push(b'[');
            for (index, value) in values.iter().enumerate() {
                if index > 0 {
                    output.push(b',');
                }
                encode_value(value, output)?;
            }
            output.push(b']');
        }
        CanonicalValue::Object(entries) => encode_object(entries, output)?,
        CanonicalValue::OrderedMap(entries) => {
            ensure_unique(entries, true)?;
            output.extend_from_slice(b"{\"$map\":[");
            for (index, (key, value)) in entries.iter().enumerate() {
                if index > 0 {
                    output.push(b',');
                }
                output.push(b'[');
                encode_string(key, output)?;
                output.push(b',');
                encode_value(value, output)?;
                output.push(b']');
            }
            output.extend_from_slice(b"]}");
        }
        CanonicalValue::Binary(bytes) => {
            let encoded = base64::engine::general_purpose::STANDARD.encode(bytes);
            output.extend_from_slice(b"{\"$binary\":{");
            output.extend_from_slice(b"\"base64\":");
            encode_string(&encoded, output)?;
            output.extend_from_slice(b",\"byteLength\":");
            output.extend_from_slice(bytes.len().to_string().as_bytes());
            output.extend_from_slice(b"}}");
        }
        CanonicalValue::Date(value) => {
            if !is_utc_iso_date(value) {
                return Err(CanonicalError::InvalidDate(value.clone()));
            }
            output.extend_from_slice(b"{\"$date\":");
            encode_string(value, output)?;
            output.push(b'}');
        }
    }
    Ok(())
}

fn encode_object(
    entries: &[(String, CanonicalValue)],
    output: &mut Vec<u8>,
) -> Result<(), CanonicalError> {
    ensure_unique(entries, false)?;
    let mut sorted: Vec<_> = entries.iter().collect();
    // Rust string ordering is UTF-8 lexicographic, which has the same order as
    // Unicode scalar values for valid UTF-8 strings.
    sorted.sort_by(|left, right| left.0.cmp(&right.0));

    output.push(b'{');
    for (index, (key, value)) in sorted.into_iter().enumerate() {
        if index > 0 {
            output.push(b',');
        }
        encode_string(key, output)?;
        output.push(b':');
        encode_value(value, output)?;
    }
    output.push(b'}');
    Ok(())
}

fn ensure_unique(entries: &[(String, CanonicalValue)], map: bool) -> Result<(), CanonicalError> {
    let mut keys = BTreeSet::new();
    for (key, _) in entries {
        if !keys.insert(key) {
            return Err(if map {
                CanonicalError::DuplicateMapKey(key.clone())
            } else {
                CanonicalError::DuplicateObjectKey(key.clone())
            });
        }
    }
    Ok(())
}

fn encode_string(value: &str, output: &mut Vec<u8>) -> Result<(), CanonicalError> {
    let encoded = serde_json::to_string(value)
        .map_err(|error| CanonicalError::Serialization(error.to_string()))?;
    output.extend_from_slice(encoded.as_bytes());
    Ok(())
}

fn encode_number(value: f64, output: &mut Vec<u8>) -> Result<(), CanonicalError> {
    if !value.is_finite() {
        return Err(CanonicalError::NonFiniteNumber);
    }
    if value == 0.0 {
        output.push(b'0');
        return Ok(());
    }
    let mut buffer = ryu_js::Buffer::new();
    output.extend_from_slice(buffer.format(value).as_bytes());
    Ok(())
}

fn is_utc_iso_date(value: &str) -> bool {
    // JavaScript Date#toISOString always has this exact ASCII shape.
    value.len() == 24
        && value.as_bytes()[4] == b'-'
        && value.as_bytes()[7] == b'-'
        && value.as_bytes()[10] == b'T'
        && value.as_bytes()[13] == b':'
        && value.as_bytes()[16] == b':'
        && value.as_bytes()[19] == b'.'
        && value.ends_with('Z')
        && value.bytes().enumerate().all(|(index, byte)| {
            matches!(index, 4 | 7 | 10 | 13 | 16 | 19 | 23) || byte.is_ascii_digit()
        })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn body(value: &CanonicalValue) -> String {
        let encoded = to_canonical_bytes(value).unwrap();
        String::from_utf8(encoded).unwrap()
    }

    #[test]
    fn shared_cross_language_golden() {
        let value = CanonicalValue::Object(vec![
            ("z".into(), CanonicalValue::Number(-0.0)),
            (
                "map".into(),
                CanonicalValue::OrderedMap(vec![
                    ("second".into(), CanonicalValue::Bool(true)),
                    ("first".into(), CanonicalValue::Null),
                ]),
            ),
            ("bytes".into(), CanonicalValue::Binary(vec![0, 255, 16])),
            (
                "date".into(),
                CanonicalValue::Date("2024-02-03T04:05:06.007Z".into()),
            ),
        ]);
        assert_eq!(
            body(&value),
            concat!(
                "docx-document-canonical-v1\n",
                "{\"bytes\":{\"$binary\":{\"base64\":\"AP8Q\",\"byteLength\":3}},",
                "\"date\":{\"$date\":\"2024-02-03T04:05:06.007Z\"},",
                "\"map\":{\"$map\":[[\"second\",true],[\"first\",null]]},\"z\":0}\n"
            )
        );
    }

    #[test]
    fn numbers_match_ecmascript_shortest_round_trip() {
        let values = CanonicalValue::Array(vec![
            CanonicalValue::Number(0.0),
            CanonicalValue::Number(-0.0),
            CanonicalValue::Number(1.0),
            CanonicalValue::Number(0.000001),
            CanonicalValue::Number(0.0000001),
            CanonicalValue::Number(1e20),
            CanonicalValue::Number(1e21),
            CanonicalValue::Number(333333333.33333329),
        ]);
        assert_eq!(
            body(&values),
            "docx-document-canonical-v1\n[0,0,1,0.000001,1e-7,100000000000000000000,1e+21,333333333.3333333]\n"
        );
    }

    #[test]
    fn rejects_non_contract_values() {
        assert_eq!(
            to_canonical_bytes(&CanonicalValue::Number(f64::NAN)),
            Err(CanonicalError::NonFiniteNumber)
        );
        assert_eq!(
            to_canonical_bytes(&CanonicalValue::OrderedMap(vec![
                ("x".into(), CanonicalValue::Null),
                ("x".into(), CanonicalValue::Bool(true)),
            ])),
            Err(CanonicalError::DuplicateMapKey("x".into()))
        );
        assert!(matches!(
            to_canonical_bytes(&CanonicalValue::Date("yesterday".into())),
            Err(CanonicalError::InvalidDate(_))
        ));
    }

    #[test]
    fn converts_typed_serde_values_without_changing_the_contract() {
        #[derive(Serialize)]
        struct Example {
            z: f64,
            a: Vec<bool>,
        }
        let value = from_serializable(&Example {
            z: 2.0,
            a: vec![true, false],
        })
        .unwrap();
        assert_eq!(
            body(&value),
            "docx-document-canonical-v1\n{\"a\":[true,false],\"z\":2}\n"
        );
    }
}
