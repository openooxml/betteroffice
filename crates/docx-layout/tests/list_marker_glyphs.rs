//! List markers paint from the same font chain as body text when one is supplied.

use docx_layout::display_list::{
    DisplayList, Primitive, build_display_list_json, build_display_list_json_with_fonts,
};
use ooxml_text::FontStore;

const LIBERATION: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn input(with_chain: Option<u32>) -> String {
    let mut input = serde_json::json!({
        "measured": [{
            "block": {
                "kind": "paragraph", "id": 0,
                "runs": [{
                    "kind": "text", "text": "Item", "pmStart": 1, "pmEnd": 5,
                    "fontFamily": "Liberation Sans", "fontSize": 12.0
                }],
                "attrs": {
                    "listMarker": "1.", "listMarkerFontFamily": "Liberation Sans",
                    "listMarkerFontSize": 12.0, "indent": {"left": 48.0, "hanging": 24.0}
                },
                "pmStart": 1, "pmEnd": 6
            },
            "measure": {
                "kind": "paragraph", "totalHeight": 20.0,
                "lines": [{
                    "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 4,
                    "width": 30.0, "ascent": 16.0, "descent": 4.0, "lineHeight": 20.0
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
                    "fromLine": 0, "toLine": 1, "pmStart": 1, "pmEnd": 6
                }]
            }]
        }
    });
    if let Some(id) = with_chain {
        input["fontChains"] = serde_json::json!({ "liberation sans|0|0": [id] });
    }
    input.to_string()
}

fn markers(list: &DisplayList) -> Vec<(&'static str, String)> {
    list.pages[0]
        .primitives
        .iter()
        .filter_map(|primitive| match primitive {
            Primitive::Text(run) if run.attrs.list_marker == Some(true) => {
                Some(("text", run.text.clone()))
            }
            Primitive::GlyphRun(run) if run.attrs.list_marker == Some(true) => {
                Some(("glyphRun", run.text.clone()))
            }
            _ => None,
        })
        .collect()
}

#[test]
fn list_markers_shape_from_the_font_chain_when_one_is_supplied() {
    let mut store = FontStore::new();
    let id = store
        .register(LIBERATION.to_vec())
        .expect("fixture registers");
    let list: DisplayList = serde_json::from_str(
        &build_display_list_json_with_fonts(&input(Some(id.to_u32())), &store).unwrap(),
    )
    .unwrap();
    let markers = markers(&list);
    assert!(!markers.is_empty());
    assert!(
        markers.iter().all(|(kind, _)| *kind == "glyphRun"),
        "{markers:?}"
    );
    assert_eq!(
        markers
            .iter()
            .map(|(_, text)| text.as_str())
            .collect::<String>(),
        "1."
    );
}

#[test]
fn list_markers_stay_text_without_fonts() {
    let list: DisplayList =
        serde_json::from_str(&build_display_list_json(&input(None)).unwrap()).unwrap();
    assert_eq!(markers(&list), vec![("text", "1.".to_owned())]);
}
