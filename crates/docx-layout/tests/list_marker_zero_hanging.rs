//! A measured zero-hanging list marker starts at the first-line indent.

use docx_layout::display_list::{DisplayList, Primitive, build_display_list_json_with_fonts};
use ooxml_text::{FontStore, measure_paragraph_json};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

#[test]
fn zero_hanging_marker_paints_at_first_line_indent_and_text_at_next_tab() {
    let mut store = FontStore::new();
    let font = store.register(FONT.to_vec()).unwrap();
    let font_chains = json!({
        "liberation sans|0|0": [font.to_u32()], "liberation sans|1|0": [font.to_u32()]
    });
    let (fragment_x, left, first_line, width) = (96.0, 12.0, 8.0, 624.0);
    let (number, points, tab_twips) = ("2.1", 16.0, 709.0);
    let block = json!({
        "kind": "paragraph", "id": 0, "pmStart": 1, "pmEnd": 6,
        "runs": [{
            "kind": "text", "text": "Beta", "pmStart": 1, "pmEnd": 5,
            "fontFamily": "Liberation Sans", "fontSize": points, "bold": true
        }],
        "attrs": {
            "listMarker": number, "listMarkerFontFamily": "Liberation Sans",
            "listMarkerFontSize": points, "listMarkerBold": true,
            "defaultTabStopTwips": tab_twips,
            "indent": {"left": left, "firstLine": first_line, "hanging": 0.0}
        }
    });
    let measure: Value = serde_json::from_str(
        &measure_paragraph_json(
            &store,
            &json!({
                "block": block, "maxWidth": width, "fontChains": font_chains,
                "authoritativeShaping": true,
                "defaults": {"fontSize": points, "fontFamily": "Liberation Sans"}
            })
            .to_string(),
        )
        .unwrap(),
    )
    .unwrap();
    assert_eq!(measure["lines"].as_array().unwrap().len(), 1);
    let input = json!({
        "measured": [{"block": block, "measure": measure}],
        "fontChains": font_chains,
        "options": {},
        "layout": {"pages": [{
            "size": {"w": 816.0, "h": 1056.0}, "number": 1,
            "margins": {"top": 96.0, "right": 96.0, "bottom": 96.0, "left": 96.0},
            "fragments": [{
                "kind": "paragraph", "blockId": 0, "x": fragment_x, "y": 96.0,
                "width": width, "height": measure["totalHeight"],
                "fromLine": 0, "toLine": 1, "pmStart": 1, "pmEnd": 6
            }]
        }]}
    });
    let list: DisplayList = serde_json::from_str(
        &build_display_list_json_with_fonts(&input.to_string(), &store).unwrap(),
    )
    .unwrap();
    let extent = |text: &str, marker: Option<bool>| {
        list.pages[0]
            .primitives
            .iter()
            .find_map(|primitive| match primitive {
                Primitive::GlyphRun(run)
                    if !run.text.is_empty()
                        && text.starts_with(run.text.as_str())
                        && run.attrs.list_marker == marker =>
                {
                    Some((
                        run.glyphs[0].x,
                        run.glyphs.iter().map(|glyph| glyph.advance).sum::<f64>(),
                    ))
                }
                Primitive::Text(run) if run.text == text && run.attrs.list_marker == marker => {
                    Some((run.x.as_f64().unwrap(), run.width.as_f64().unwrap()))
                }
                _ => None,
            })
            .unwrap_or_else(|| panic!("no {text:?} run in {:?}", list.pages[0].primitives))
    };
    let (marker_x, marker_width) = extent(number, Some(true));
    let (text_x, _) = extent("Beta", None);
    let tab = tab_twips / 15.0;
    let next_stop = ((left + first_line + marker_width) / tab).floor() + 1.0;
    let expected_text_x = fragment_x + next_stop * tab;
    assert_eq!(marker_x, fragment_x + left + first_line, "marker origin");
    assert!(
        (text_x - expected_text_x).abs() < 0.5,
        "text x {text_x} vs {expected_text_x}"
    );
    assert!(
        text_x >= marker_x + marker_width,
        "text overlaps the marker"
    );
}
