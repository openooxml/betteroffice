use ooxml_text::{FontStore, measure_paragraph_json};
use serde_json::json;

const BACKSTOP: &[u8] = include_bytes!("../../base-latin/fonts/BetterOfficeBaseLatin-Regular.ttf");
const PRIMARY: &[u8] = include_bytes!("../tests/fonts/LiberationSans-Regular.ttf");

#[cfg(feature = "base-latin")]
#[test]
fn native_implicit_aliases_require_a_symbol_family() {
    let mut store = FontStore::new();
    let primary = store.register_slice(PRIMARY).unwrap();
    assert_eq!(
        store.resolve_with_family(&[primary], '\u{f06f}', false, false, "Wingdings 2"),
        Some(ooxml_text::FontId::base_latin(false, false))
    );
    assert_eq!(
        store.resolve_with_family(&[primary], '\u{f06f}', false, false, "Custom Logo"),
        Some(primary)
    );
}

#[test]
fn private_use_aliases_require_a_symbol_family_for_registered_faces() {
    let mut store = FontStore::new();
    let primary = store.register_slice(PRIMARY).unwrap();
    let backstop = store.register_slice(BACKSTOP).unwrap();
    let chain = [primary, backstop];
    for ch in [
        '\u{f0b7}', '\u{f0a7}', '\u{f06f}', '\u{f0d8}', '\u{f0fc}', '\u{f076}', '\u{f0a8}',
    ] {
        for family in [
            "Symbol",
            "sYmBoL",
            "Wingdings",
            "WINGDINGS 2",
            "wingdings 3",
        ] {
            assert_eq!(
                store.resolve_with_family(&chain, ch, false, false, family),
                Some(backstop)
            );
        }
        for family in ["Custom Logo", "Arial", "Wingdings 4", ""] {
            assert_eq!(
                store.resolve_with_family(&chain, ch, false, false, family),
                Some(primary)
            );
            assert_eq!(
                store.resolve_with_family(&[backstop], ch, false, false, family),
                None
            );
        }
    }
}

#[test]
fn an_embedded_private_use_face_keeps_its_own_glyphs() {
    let from: Vec<u8> = "BetterOffice Base Latin"
        .encode_utf16()
        .flat_map(u16::to_be_bytes)
        .collect();
    let to: Vec<u8> = "Document PUA Glyph Face"
        .encode_utf16()
        .flat_map(u16::to_be_bytes)
        .collect();
    assert_eq!(from.len(), to.len());
    let mut data = BACKSTOP.to_vec();
    for index in 0..=data.len() - from.len() {
        if data[index..index + from.len()] == from {
            data[index..index + to.len()].copy_from_slice(&to);
        }
    }
    let mut store = FontStore::new();
    let embedded = store.register(data).unwrap();
    let backstop = store.register_slice(BACKSTOP).unwrap();
    assert_eq!(store.backstop_style(embedded), None);
    assert_eq!(
        store.resolve_with_family(
            &[embedded, backstop],
            '\u{f06f}',
            false,
            false,
            "Custom Logo"
        ),
        Some(embedded)
    );
}

#[test]
fn measurement_scopes_aliases_for_text_fields_and_markers() {
    let mut store = FontStore::new();
    let backstop = store.register_slice(BACKSTOP).unwrap().to_u32();
    for family in ["Symbol", "Custom Logo"] {
        for kind in ["text", "field", "marker"] {
            let run = if kind == "field" {
                json!({"kind": "field", "fallback": "\u{f06f}"})
            } else {
                json!({"kind": "text", "text": if kind == "marker" { "A" } else { "\u{f06f}" }})
            };
            let input = json!({
                "block": {"kind": "paragraph", "runs": [run], "attrs": {
                    "listMarker": if kind == "marker" { "\u{f06f}" } else { "" },
                    "listMarkerFontFamily": family
                }},
                "fontChains": {format!("{}|0|0", family.to_lowercase()): [backstop]},
                "defaults": {"fontFamily": family, "fontSize": 11}, "maxWidth": 600
            });
            let measured = measure_paragraph_json(&store, &input.to_string());
            assert_eq!(measured.is_ok(), family == "Symbol", "{kind}: {measured:?}");
        }
    }
}
