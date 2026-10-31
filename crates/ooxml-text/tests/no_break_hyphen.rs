use ooxml_text::{
    FontStore, ShapeDirection, echo_glyph_set, kern_features, measure_paragraph_json,
    shape_with_direction,
};
use serde_json::{Value, json};

const LIBERATION: &[u8] = include_bytes!("fonts/LiberationSans-Regular.ttf");
const CARLITO: &[u8] = include_bytes!("../../../packages/fonts/assets/Carlito-Regular.ttf");
const INTER: &[u8] = include_bytes!("../../../packages/fonts/assets/Inter-Regular.ttf");

fn measure(store: &FontStore, text: &str, max_width: f64, chain: &[u32]) -> Value {
    let output = measure_paragraph_json(
        store,
        &json!({
            "block": {"kind": "paragraph", "runs": [{"kind": "text", "text": text}]},
            "maxWidth": max_width,
            "fontChains": {"requested|0|0": chain},
            "defaults": {"fontSize": 12.0, "fontFamily": "Requested"},
            "authoritativeShaping": true
        })
        .to_string(),
    )
    .expect("hyphen measures natively");
    serde_json::from_str(&output).unwrap()
}

fn width(measure: &Value) -> f64 {
    measure["lines"][0]["width"].as_f64().unwrap()
}

#[test]
fn nonbreaking_hyphen_uses_the_covered_hyphen_width() {
    let mut store = FontStore::new();
    let font = store.register(LIBERATION.to_vec()).unwrap();
    assert!(!store.covers(font, '\u{2011}').unwrap());
    assert!(store.covers(font, '\u{2010}').unwrap());
    let chain = [font.to_u32()];
    let expected = measure(&store, "co\u{2010}operate", 1000.0, &chain);
    let actual = measure(&store, "co\u{2011}operate", 1000.0, &chain);
    assert_eq!(width(&actual), width(&expected));
    assert_eq!(actual["lines"][0]["tailChar"], 10);
    assert_eq!(
        actual["lines"][0]["clusterAdvances"],
        expected["lines"][0]["clusterAdvances"]
    );
}

#[test]
fn nonbreaking_hyphen_does_not_add_a_wrap_opportunity() {
    let mut store = FontStore::new();
    let font = store.register(LIBERATION.to_vec()).unwrap();
    let chain = [font.to_u32()];
    let expected = measure(&store, "co\u{2010}operate", 1000.0, &chain);
    let max_width = width(&expected) + 1.0;
    let breakable = measure(&store, "aaa co\u{2010}operate", max_width, &chain);
    let nonbreaking = measure(&store, "aaa co\u{2011}operate", max_width, &chain);
    assert_eq!(breakable["lines"][0]["tailChar"], 7);
    assert_eq!(nonbreaking["lines"][0]["tailChar"], 4);
    assert_eq!(nonbreaking["lines"][1]["headChar"], 4);
    assert_eq!(
        nonbreaking["lines"][1]["width"].as_f64().unwrap(),
        width(&expected)
    );
}

#[test]
fn nonbreaking_hyphen_uses_ascii_hyphen_in_carlito() {
    let mut store = FontStore::new();
    let font = store.register(CARLITO.to_vec()).unwrap();
    assert!(!store.covers(font, '\u{2011}').unwrap());
    assert!(!store.covers(font, '\u{2010}').unwrap());
    assert!(store.covers(font, '-').unwrap());
    let chain = [font.to_u32()];
    let expected = measure(&store, "co-operate", 1000.0, &chain);
    let actual = measure(&store, "co\u{2011}operate", 1000.0, &chain);
    assert_eq!(width(&actual), width(&expected));
}

#[test]
fn shaping_keeps_original_byte_clusters_and_prefers_a_native_glyph() {
    for (bytes, stand_in) in [
        (LIBERATION, '\u{2010}'),
        (CARLITO, '-'),
        (INTER, '\u{2011}'),
    ] {
        let mut store = FontStore::new();
        let font = store.register(bytes.to_vec()).unwrap();
        for direction in [ShapeDirection::Ltr, ShapeDirection::Rtl] {
            let text = "co\u{2011}operate";
            let glyphs =
                shape_with_direction(&store, font, text, 16.0, &kern_features(false), direction)
                    .unwrap();
            let hyphen = glyphs.iter().find(|glyph| glyph.cluster == 2).unwrap();
            assert_eq!(
                hyphen.glyph_id,
                u32::from(store.glyph_id(font, stand_in).unwrap().unwrap())
            );
            let mut clusters: Vec<_> = glyphs.iter().map(|glyph| glyph.cluster).collect();
            clusters.sort_unstable();
            assert_eq!(
                clusters,
                text.char_indices()
                    .map(|(byte, _)| byte as u32)
                    .collect::<Vec<_>>()
            );
        }
    }
}

#[test]
fn fallback_chain_prefers_a_stand_in_in_its_first_face() {
    let mut store = FontStore::new();
    let carlito = store.register(CARLITO.to_vec()).unwrap();
    let liberation = store.register(LIBERATION.to_vec()).unwrap();
    let inter = store.register(INTER.to_vec()).unwrap();
    assert_eq!(
        store.resolve(&[carlito, liberation, inter], '\u{2011}'),
        Some(carlito)
    );
    assert_eq!(store.resolve(&[inter, carlito], '\u{2011}'), Some(inter));
    let chain = [carlito.to_u32(), liberation.to_u32(), inter.to_u32()];
    let expected = measure(&store, "co-operate", 1000.0, &chain);
    let actual = measure(&store, "co\u{2011}operate", 1000.0, &chain);
    assert_eq!(width(&actual), width(&expected));
}

#[test]
fn substitute_widths_use_the_stand_in_without_changing_clusters() {
    let mut store = FontStore::new();
    let base = store.register(CARLITO.to_vec()).unwrap();
    let font = store
        .register_named_substitute(base, "Arial Narrow")
        .unwrap();
    let shape = |text| {
        shape_with_direction(
            &store,
            font,
            text,
            16.0,
            &kern_features(false),
            ShapeDirection::Ltr,
        )
        .unwrap()
    };
    let expected = shape("co-operate");
    let actual = shape("co\u{2011}operate");
    assert_eq!(actual[2].glyph_id, expected[2].glyph_id);
    assert_eq!(actual[2].x_advance, expected[2].x_advance);
    assert_eq!(actual[3].cluster, 5);
    let chain = [font.to_u32()];
    let stand_in = width(&measure(&store, "co\u{2011}operate", 1000.0, &chain));
    let hyphen = width(&measure(&store, "co-operate", 1000.0, &chain));
    assert!((stand_in - hyphen).abs() < 1e-3, "{stand_in} vs {hyphen}");
}

#[test]
fn typing_echo_uses_the_same_hyphen_glyph_and_advance() {
    let mut store = FontStore::new();
    let font = store.register(LIBERATION.to_vec()).unwrap();
    let echo = echo_glyph_set(&store, font, 16.0, false, &['\u{2010}', '\u{2011}']).unwrap();
    assert_ne!(echo.glyph_ids[0], 0);
    assert_eq!(echo.glyph_ids[1], echo.glyph_ids[0]);
    assert_eq!(echo.advances[1], echo.advances[0]);
}
