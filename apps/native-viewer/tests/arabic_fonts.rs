#![cfg(feature = "docx")]

use ooxml_text::FontId;
use serde_json::json;

#[allow(dead_code)]
mod fonts {
    include!(concat!(env!("CARGO_MANIFEST_DIR"), "/src/fonts.rs"));
}

const NASKH_REGULAR: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../packages/fonts/assets/NotoNaskhArabic-Regular.ttf"
));
const SANS_REGULAR: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../packages/fonts/assets/NotoSansArabic-Regular.ttf"
));
const SANS_BOLD: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../packages/fonts/assets/NotoSansArabic-Bold.ttf"
));

#[test]
fn arabic_fallback_respects_requested_family() {
    for (bold, italic, style) in [
        (false, false, ""),
        (true, false, "bold "),
        (false, true, "italic "),
        (true, true, "italic bold "),
    ] {
        let sans = if bold { SANS_BOLD } else { SANS_REGULAR };
        for (family, expected) in [("Arial", NASKH_REGULAR), ("Noto Sans Arabic", sans)] {
            let key = format!(
                "{}|{}|{}",
                family.to_lowercase(),
                u8::from(bold),
                u8::from(italic)
            );
            let registry = fonts::FontRegistry::load(&json!([{
                "key": key,
                "family": family,
                "bold": bold,
                "italic": italic,
                "scripts": ["arabic"],
            }]))
            .unwrap();
            let font = format!("{style}16px '{family}'");
            let (_, runs) = registry
                .shape_fallback("سلام", &font, true, false, 0.0, 0.0)
                .unwrap();
            assert_eq!(runs.len(), 1, "{font}");
            let id = FontId::from_u32(runs[0].0);
            let actual = registry.store.font_bytes(id).unwrap();
            assert!(
                actual == expected,
                "{font}: resolved {:?}",
                registry.store.declared_families(id).unwrap()
            );
        }
    }
}
