use std::collections::{BTreeMap, HashMap};
use std::sync::Arc;

use docx_edit::{ParagraphProperties, SegmentContent, StorySegment, segments_digest};
use yrs::Any;

fn text(value: &str, attributes: BTreeMap<String, Any>) -> StorySegment {
    StorySegment {
        content: SegmentContent::Text(value.to_string()),
        attributes,
    }
}

fn entries<const N: usize>(pairs: [(&str, Any); N]) -> BTreeMap<String, Any> {
    pairs
        .into_iter()
        .map(|(key, value)| (key.to_string(), value))
        .collect()
}

fn any_map<const N: usize>(pairs: [(&str, Any); N]) -> Any {
    Any::Map(Arc::new(
        pairs
            .into_iter()
            .map(|(key, value)| (key.to_string(), value))
            .collect::<HashMap<String, Any>>(),
    ))
}

fn any_array<const N: usize>(values: [Any; N]) -> Any {
    Any::Array(Arc::from(values.to_vec()))
}

fn any_buffer(bytes: &[u8]) -> Any {
    Any::Buffer(Arc::from(bytes))
}

fn any_string(value: &str) -> Any {
    Any::String(Arc::from(value))
}

fn fixed_units() -> Vec<StorySegment> {
    vec![
        text(
            "First line of the sample paragraph",
            entries([
                ("bold", Any::Bool(true)),
                ("size", Any::Number(11.5)),
                ("font", any_string("Serif")),
            ]),
        ),
        StorySegment {
            content: SegmentContent::Pilcrow(ParagraphProperties {
                para_id: "1A2B3C4D".to_string(),
                values: entries([
                    ("pStyle", any_string("Heading1")),
                    ("level", Any::BigInt(-3)),
                    ("keep", Any::Null),
                    ("spare", Any::Undefined),
                    ("tabs", any_array([Any::Number(720.0), Any::Number(1440.0)])),
                ]),
            }),
            attributes: BTreeMap::new(),
        },
        StorySegment {
            content: SegmentContent::OtherEmbed {
                kind: "table".to_string(),
                payload: entries([
                    (
                        "rows",
                        any_array([
                            any_map([
                                ("height", Any::Number(240.0)),
                                ("cells", any_array([any_string("a"), any_string("b")])),
                            ]),
                            any_map([]),
                            any_array([]),
                        ]),
                    ),
                    (
                        "grid",
                        any_map([
                            (
                                "cols",
                                any_array([Any::BigInt(1), Any::BigInt(2), Any::BigInt(3)]),
                            ),
                            ("blob", any_buffer(&[9, 8, 7])),
                            ("flag", Any::Bool(false)),
                        ]),
                    ),
                    ("image", any_buffer(&[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11])),
                    ("empty", any_buffer(&[])),
                ]),
            },
            attributes: entries([("anchor", any_string("inline"))]),
        },
        text("Tail text after the table", BTreeMap::new()),
    ]
}

#[test]
fn segment_digests_are_fixed_across_platforms() {
    assert_eq!(
        segments_digest(&fixed_units()),
        0xc38e897e5bb2b49f479308803311a066
    );
    assert_eq!(segments_digest(&[]), 0xe555b420703a0ad4aacd33b0879bde5f);
    assert_eq!(
        segments_digest(&[text("Tail text after the table", BTreeMap::new())]),
        0xf3bfbf304bc301162957e93132a465a4
    );
}
