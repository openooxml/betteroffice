#![cfg_attr(not(feature = "wasm"), allow(dead_code))]

use std::collections::BTreeMap;

use serde::ser::SerializeMap;
use serde::{Serialize, Serializer};
use yrs::Any;

use crate::{SegmentContent, StorySegment};

pub(crate) fn segments_json_string(segments: &[StorySegment]) -> Result<String, serde_json::Error> {
    serde_json::to_string(&SegmentsJson(segments))
}

struct SegmentsJson<'a>(&'a [StorySegment]);
struct SegmentJson<'a>(&'a StorySegment);
struct AttrsJson<'a>(&'a BTreeMap<String, Any>);
struct AnyJson<'a>(&'a Any);

impl Serialize for SegmentsJson<'_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        serializer.collect_seq(self.0.iter().map(SegmentJson))
    }
}

impl Serialize for SegmentJson<'_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        let segment = self.0;
        let len = match &segment.content {
            SegmentContent::Text(_) => 3,
            _ => 4,
        };
        let mut map = serializer.serialize_map(Some(len))?;
        map.serialize_entry("attributes", &AttrsJson(&segment.attributes))?;
        match &segment.content {
            SegmentContent::Text(text) => {
                map.serialize_entry("kind", "text")?;
                map.serialize_entry("text", text)?;
            }
            SegmentContent::Pilcrow(properties) => {
                map.serialize_entry("kind", "pilcrow")?;
                map.serialize_entry("paraId", &properties.para_id)?;
                map.serialize_entry("properties", &AttrsJson(&properties.values))?;
            }
            SegmentContent::OtherEmbed { kind, payload } => {
                map.serialize_entry("embedKind", kind)?;
                map.serialize_entry("kind", "embed")?;
                map.serialize_entry("payload", &AttrsJson(payload))?;
            }
        }
        map.end()
    }
}

impl Serialize for AttrsJson<'_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        serializer.collect_map(self.0.iter().map(|(key, value)| (key, AnyJson(value))))
    }
}

impl Serialize for AnyJson<'_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        match self.0 {
            Any::Array(values) => serializer.collect_seq(values.iter().map(AnyJson)),
            Any::Map(entries) => {
                let mut entries: Vec<_> = entries.iter().collect();
                entries.sort_unstable_by(|(left, _), (right, _)| left.cmp(right));
                serializer.collect_map(
                    entries
                        .into_iter()
                        .map(|(key, value)| (key, AnyJson(value))),
                )
            }
            value => value.serialize(serializer),
        }
    }
}

#[cfg(any(feature = "wasm", test))]
pub(crate) fn segments_json(
    segments: Vec<StorySegment>,
) -> Result<Vec<serde_json::Value>, serde_json::Error> {
    use serde_json::json;

    segments
        .into_iter()
        .map(|segment| {
            let attributes = serde_json::to_value(&segment.attributes)?;
            Ok(match segment.content {
                SegmentContent::Text(text) => {
                    json!({ "kind": "text", "text": text, "attributes": attributes })
                }
                SegmentContent::Pilcrow(properties) => json!({
                    "kind": "pilcrow",
                    "paraId": properties.para_id,
                    "properties": serde_json::to_value(&properties.values)?,
                    "attributes": attributes,
                }),
                SegmentContent::OtherEmbed { kind, payload } => {
                    json!({
                        "kind": "embed",
                        "embedKind": kind,
                        "payload": serde_json::to_value(&payload)?,
                        "attributes": attributes,
                    })
                }
            })
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;
    use std::sync::Arc;

    use yrs::{Map, ReadTxn, Transact};

    use super::*;
    use crate::{EngineSession, ParagraphProperties, seed_from_docx};

    fn assert_parity(segments: &[StorySegment]) {
        assert_eq!(
            segments_json_string(segments).unwrap(),
            serde_json::to_string(&segments_json(segments.to_vec()).unwrap()).unwrap()
        );
    }

    fn any_map(entries: &[(&str, Any)]) -> Any {
        Any::Map(Arc::new(
            entries
                .iter()
                .map(|(key, value)| (key.to_string(), value.clone()))
                .collect(),
        ))
    }

    fn all_contents(values: BTreeMap<String, Any>) -> Vec<StorySegment> {
        [
            SegmentContent::Text("\"\\\0\u{8}\t\n\u{c}\r\u{1f}\u{2028}\u{2029}😀𝄞".into()),
            SegmentContent::Text(String::new()),
            SegmentContent::Pilcrow(ParagraphProperties {
                para_id: "para\"\\\n😀".into(),
                values: values.clone(),
            }),
            SegmentContent::OtherEmbed {
                kind: "kind\"\\\n😀".into(),
                payload: values.clone(),
            },
        ]
        .into_iter()
        .map(|content| StorySegment {
            content,
            attributes: values.clone(),
        })
        .collect()
    }

    #[test]
    fn synthetic_segments_match_value_oracle() {
        let text: String = (0..=31)
            .map(|value| char::from_u32(value).unwrap())
            .chain("\"\\/\u{7f}\u{2028}\u{2029}😀𝄞".chars())
            .collect();
        let scalars = vec![
            Any::Null,
            Any::Undefined,
            Any::Bool(false),
            Any::Bool(true),
            Any::String(text.into()),
            Any::String("".into()),
            Any::Buffer(Arc::from([0, 1, 31, 127, 128, 255])),
            Any::Buffer(Arc::from([])),
            Any::Array(Arc::from([])),
            Any::Map(Arc::new(HashMap::new())),
            Any::BigInt(i64::MIN),
            Any::BigInt(i64::MAX),
            Any::BigInt(0),
            Any::BigInt(-1),
        ];
        let numbers = [
            0.0,
            -0.0,
            1.0,
            -1.0,
            1.5,
            -1.5,
            1e-300,
            -1e300,
            f64::MIN,
            f64::MAX,
            f64::MIN_POSITIVE,
            f64::from_bits(1),
            -f64::from_bits(1),
            9_007_199_254_740_991.0,
            9_007_199_254_740_992.0,
            i64::MIN as f64,
            i64::MAX as f64,
            f64::from_bits((i64::MAX as f64).to_bits() - 1),
            f64::from_bits((i64::MAX as f64).to_bits() + 1),
            f64::NAN,
            f64::INFINITY,
            f64::NEG_INFINITY,
        ];
        let values: Vec<_> = scalars
            .into_iter()
            .chain(numbers.into_iter().map(Any::Number))
            .collect();
        let nested = any_map(&[
            ("z", Any::Array(values.clone().into())),
            (
                "a\"\\\n😀",
                any_map(&[("z", Any::Undefined), ("a", Any::Null)]),
            ),
            ("ä", Any::Bool(true)),
            ("😀", Any::Bool(false)),
        ]);
        let mut attributes: BTreeMap<_, _> = values
            .into_iter()
            .enumerate()
            .map(|(index, value)| (format!("value{index}"), value))
            .collect();
        attributes.insert("nested".into(), nested.clone());
        attributes.insert("array".into(), Any::Array(vec![nested].into()));
        assert_parity(&[]);
        assert_parity(&all_contents(BTreeMap::new()));
        assert_parity(&all_contents(attributes));
    }

    #[test]
    fn nested_maps_and_arrays_match_value_oracle() {
        let mut value = Any::String("\"\\\n😀".into());
        for _ in 0..64 {
            value = any_map(&[
                ("z", Any::Null),
                ("a", Any::Array(vec![Any::Undefined, value].into())),
            ]);
        }
        assert_parity(&all_contents(BTreeMap::from([("nested".into(), value)])));
    }

    #[test]
    fn number_bit_patterns_match_value_oracle() {
        let mut bits = 1_u64;
        let mut numbers = Vec::new();
        for _ in 0..1024 {
            bits ^= bits << 13;
            bits ^= bits >> 7;
            bits ^= bits << 17;
            numbers.push(Any::Number(f64::from_bits(bits)));
        }
        assert_parity(&all_contents(BTreeMap::from([(
            "numbers".into(),
            Any::Array(numbers.into()),
        )])));
    }

    #[test]
    fn every_corpus_story_matches_value_oracle() {
        for (name, bytes) in crate::engine::preview_fixture::corpus() {
            let engine = EngineSession::new(75200);
            seed_from_docx(engine.doc(), bytes).unwrap();
            let stories = {
                let txn = engine.doc().yrs_doc().transact();
                let mut stories: Vec<_> = txn
                    .get_map(crate::STORIES)
                    .unwrap()
                    .keys(&txn)
                    .map(str::to_owned)
                    .collect();
                stories.sort();
                stories
            };
            assert!(!stories.is_empty(), "{name}");
            for story in stories {
                let segments = engine.doc().story_segments(&story).unwrap();
                assert_eq!(
                    segments_json_string(&segments).unwrap(),
                    serde_json::to_string(&segments_json(segments).unwrap()).unwrap(),
                    "{name}: {story}"
                );
            }
        }
    }
}
