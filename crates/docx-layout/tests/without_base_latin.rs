#![cfg(not(feature = "base-latin"))]

use ooxml_text::{FontId, FontStore};

#[test]
fn feature_off_has_no_reserved_faces_or_implicit_backstop() {
    let mut store = FontStore::new();
    for bold in [false, true] {
        for italic in [false, true] {
            assert!(store.font_bytes(FontId::base_latin(bold, italic)).is_err());
        }
    }
    let font = store
        .register_slice(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
    assert_eq!(store.resolve_with_style(&[], 'A', false, false), None);
    assert_eq!(
        store.resolve_with_style(&[font], '\u{10ffff}', false, false),
        Some(font)
    );
}

#[test]
fn requirements_assign_characters_to_the_selected_slots_and_marker_family() {
    let blocks: Vec<docx_layout::types::LayoutBlock> = serde_json::from_value(serde_json::json!([{
        "kind": "paragraph", "id": "slots", "runs": [{
            "kind": "text", "text": "A\u{f06f}א", "fontFamily": "Custom Logo",
            "fontSlots": {"ascii": "Latin", "hAnsi": "Custom Logo", "cs": "Complex"}
        }, {
            "kind": "text", "text": "B", "complexScript": true, "fontSlots": {"cs": "Complex"}
        }], "attrs": {"listMarker": "\u{f0b7}", "listMarkerFontFamily": "Wingdings"}
    }]))
    .unwrap();
    let requirements = docx_layout::measure_blocks::collect_font_requirements(&blocks, "Calibri");
    let characters = |family: &str| {
        requirements
            .iter()
            .find(|requirement| requirement.family == family)
            .unwrap()
            .characters
            .iter()
            .collect::<String>()
    };
    assert_eq!(characters("Latin"), "A");
    assert_eq!(characters("Custom Logo"), "\u{f06f}");
    assert_eq!(characters("Complex"), "B");
    assert_eq!(characters("Wingdings"), "\u{f0b7}");
}

#[test]
fn requirement_bitmaps_are_bounded_and_keep_transformed_text() {
    let collect = |text: &str| {
        let blocks = serde_json::from_value::<Vec<docx_layout::types::LayoutBlock>>(
            serde_json::json!([{
                "kind": "paragraph", "id": "caps", "runs": [{"kind": "text", "text": text, "allCaps": true,
                    "language": {"latin": "tr-TR"}}]
            }]),
        )
        .unwrap();
        docx_layout::measure_blocks::collect_font_requirements(&blocks, "Requested")
    };
    let short = collect("aiΩ\u{f06f}");
    let long = collect(&"aiΩ\u{f06f}".repeat(10_000));
    assert_eq!(
        serde_json::to_value(&short).unwrap(),
        serde_json::to_value(&long).unwrap()
    );
    let requirement = short
        .iter()
        .find(|entry| entry.family == "Requested")
        .unwrap();
    assert_eq!(
        requirement.characters.iter().collect::<String>(),
        "AİΩ\u{f06f}"
    );
    let value = serde_json::to_value(requirement).unwrap();
    assert!(value.get("characters").is_none());
    assert!(value["characterBits"].as_str().unwrap().len() <= 512);
}

#[test]
fn implicit_hyphen_measurement_and_display_match_the_explicit_chain() {
    use docx_layout::display_list::{Primitive, build_display_list_value_with_fonts};
    use serde_json::{Value, json};

    let mut fonts = FontStore::new();
    let id = fonts
        .register_slice(include_bytes!(
            "../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf"
        ))
        .unwrap();
    let block = json!({"kind": "paragraph", "id": 0, "runs": [{
        "kind": "text", "text": "co\u{2011}operate", "fontFamily": "Requested", "pmStart": 1, "pmEnd": 11
    }]});
    let build = |implicit: bool| {
        let mut chains = json!({"requested|0|0": [id.to_u32()]});
        if implicit {
            chains["@betteroffice/fonts/coverage"] = json!([id.to_u32()]);
        }
        let measure: Value = serde_json::from_str(
            &ooxml_text::measure_paragraph_json(
                &fonts,
                &json!({
                    "block": block, "fontChains": chains, "maxWidth": 340,
                    "defaults": {"fontFamily": "Requested", "fontSize": 11}, "authoritativeShaping": true
                })
                .to_string(),
            )
            .unwrap(),
        )
        .unwrap();
        let mut input = json!({"measured": [{"block": block, "measure": measure}], "fontChains": chains,
            "options": {"pageSize": {"w": 360, "h": 200}, "margins": {"top": 10, "right": 10, "bottom": 10, "left": 10}}});
        input["layout"] =
            serde_json::from_str(&docx_layout::layout_to_json(&input.to_string()).unwrap())
                .unwrap();
        build_display_list_value_with_fonts(&input.to_string(), &fonts).unwrap()
    };
    let explicit = build(false);
    assert!(
        explicit
            .pages
            .iter()
            .flat_map(|page| &page.primitives)
            .any(|primitive| matches!(primitive, Primitive::GlyphRun(_)))
    );
    assert_eq!(build(true), explicit);
}
