//! The language a text primitive carries follows the script of its text.

use docx_layout::display_list::{DisplayList, Primitive, build_display_list_json};

fn primitive_language(text: &str, language: serde_json::Value) -> Option<String> {
    let count = text.chars().count() as i64;
    let input = serde_json::json!({
        "measured": [{
            "block": {
                "kind": "paragraph",
                "id": 0,
                "runs": [{
                    "kind": "text", "text": text, "pmStart": 1, "pmEnd": 1 + count,
                    "fontFamily": "Liberation Sans", "fontSize": 12.0, "language": language
                }],
                "pmStart": 1,
                "pmEnd": 2 + count
            },
            "measure": {
                "kind": "paragraph",
                "totalHeight": 20.0,
                "lines": [{
                    "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": count,
                    "width": 0.0, "ascent": 16.0, "descent": 4.0, "lineHeight": 20.0
                }]
            }
        }],
        "options": {},
        "layout": {
            "pages": [{
                "size": {"w": 816.0, "h": 1056.0},
                "margins": {"top": 96.0, "right": 96.0, "bottom": 96.0, "left": 96.0},
                "number": 1,
                "fragments": [{
                    "kind": "paragraph", "blockId": 0,
                    "x": 96.0, "y": 96.0, "width": 624.0, "height": 20.0,
                    "fromLine": 0, "toLine": 1, "pmStart": 1, "pmEnd": 2 + count
                }]
            }]
        }
    });
    let list: DisplayList =
        serde_json::from_str(&build_display_list_json(&input.to_string()).unwrap()).unwrap();
    list.pages[0]
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::Text(run) => Some(run.attrs.lang.clone()),
            Primitive::GlyphRun(run) => Some(run.attrs.lang.clone()),
            _ => None,
        })
        .expect("a text primitive")
}

#[test]
fn latin_text_carries_its_latin_language_in_an_east_asian_document() {
    for east_asia in ["zh-CN", "ja-JP"] {
        let language = serde_json::json!({"latin": "en-US", "eastAsia": east_asia});
        assert_eq!(
            primitive_language("Annual report", language).as_deref(),
            Some("en-US")
        );
    }
}

#[test]
fn east_asian_text_keeps_its_east_asian_language() {
    let language = serde_json::json!({"latin": "en-US", "eastAsia": "ja-JP"});
    assert_eq!(
        primitive_language("年次報告", language).as_deref(),
        Some("ja-JP")
    );
}

#[test]
fn right_to_left_text_carries_its_bidi_language() {
    let language = serde_json::json!({"latin": "en-US", "eastAsia": "zh-CN", "bidi": "he-IL"});
    assert_eq!(
        primitive_language("שלום", language).as_deref(),
        Some("he-IL")
    );
}

#[test]
fn latin_text_without_a_latin_language_carries_none() {
    let language = serde_json::json!({"eastAsia": "zh-CN"});
    assert_eq!(primitive_language("Annual report", language), None);
}
