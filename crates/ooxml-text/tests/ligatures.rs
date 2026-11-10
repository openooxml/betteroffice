use ooxml_text::{
    FontStore, MeasuredGlyphRun, ShapeDirection, kern_features, ligature_features,
    measure_paragraph_json, shape_with_direction, unpack_glyph_runs,
};
use serde_json::{Value, json};

const CARLITO: &[u8] = include_bytes!("../../../packages/fonts/assets/Carlito-Regular.ttf");
const ARIAL: &[u8] = include_bytes!("fonts/LiberationSans-Regular.ttf");
const ARABIC: &[u8] = include_bytes!("../../../packages/fonts/assets/NotoNaskhArabic-Regular.ttf");

fn measure(store: &FontStore, text: &str, settings: Option<&str>) -> Vec<MeasuredGlyphRun> {
    let run = match settings {
        Some(settings) => json!({
            "kind": "text", "text": text,
            "wordFeatures": ligature_features(Some(settings), Some(false))
        }),
        None => json!({"kind": "text", "text": text}),
    };
    let output = measure_paragraph_json(
        store,
        &json!({
            "block": {"kind": "paragraph", "runs": [run]},
            "maxWidth": 1000, "fontChains": {"face|0|0": [0]},
            "defaults": {"fontSize": 12, "fontFamily": "Face"},
            "authoritativeShaping": true,
        })
        .to_string(),
    )
    .unwrap();
    let output: Value = serde_json::from_str(&output).unwrap();
    unpack_glyph_runs(output["lines"][0]["glyphRuns"].as_str().unwrap()).unwrap()
}

#[test]
fn carlito_joins_only_when_word_requests_it() {
    let mut store = FontStore::new();
    let font = store.register(CARLITO.to_vec()).unwrap();
    let text = "fi ti";
    let separate = measure(&store, text, Some("none"));
    let joined = measure(&store, text, Some("standardContextual"));
    let count = |runs: &[MeasuredGlyphRun]| runs.iter().map(|run| run.glyphs.len()).sum::<usize>();
    assert_eq!(count(&separate), text.chars().count());
    assert!(count(&joined) < count(&separate));
    assert_eq!(count(&measure(&store, text, None)), count(&joined));
    for (runs, setting) in [(&separate, "none"), (&joined, "standardContextual")] {
        let mut features = kern_features(false);
        features.extend(ligature_features(Some(setting), Some(false)));
        let shaped =
            shape_with_direction(&store, font, text, 16.0, &features, ShapeDirection::Ltr).unwrap();
        let measured: f32 = runs
            .iter()
            .flat_map(|run| &run.glyphs)
            .map(|glyph| glyph.advance)
            .sum();
        let expected: f32 = shaped.iter().map(|glyph| glyph.x_advance).sum();
        assert!((measured - expected).abs() < 0.001);
    }
}

#[test]
fn arial_and_arabic_keep_their_glyphs_and_advances() {
    for (bytes, text) in [(ARIAL, "fi ti"), (ARABIC, "مرحبا بالعالم")] {
        let mut store = FontStore::new();
        store.register(bytes.to_vec()).unwrap();
        let signature = |setting| {
            measure(&store, text, Some(setting))
                .into_iter()
                .flat_map(|run| run.glyphs)
                .map(|glyph| (glyph.id, glyph.advance))
                .collect::<Vec<_>>()
        };
        assert_eq!(signature("none"), signature("standardContextual"));
        let defaults: Vec<_> = measure(&store, text, None)
            .into_iter()
            .flat_map(|run| run.glyphs)
            .map(|glyph| (glyph.id, glyph.advance))
            .collect();
        assert_eq!(signature("none"), defaults);
    }
}
